#!/usr/bin/env node
/**
 * 議事録AI サーバー
 * 音声ファイル → Gemini File API文字起こし → 議事録生成 → PDF/Word出力
 * APIキーはフロントエンドから受け取る（サーバーに埋め込まない）
 */

import express from "express";
import multer from "multer";
import { resolve, dirname } from "path";
import { fileURLToPath } from "url";
import { execFile } from "child_process";
import { writeFileSync, readFileSync, unlinkSync, mkdirSync, existsSync } from "fs";
import XLSX from "xlsx";
import {
  Document, Packer, Paragraph, TextRun, HeadingLevel,
  AlignmentType, BorderStyle, LevelFormat,
  Table, TableRow, TableCell, WidthType, VerticalAlign,
} from "docx";

const __dirname = dirname(fileURLToPath(import.meta.url));
const IS_WINDOWS = process.platform === "win32";
const TMP_DIR = resolve(__dirname, "tmp");
mkdirSync(TMP_DIR, { recursive: true });

const MODEL = "gemini-3-flash-preview";
const MODEL_FALLBACK = "gemini-2.5-flash";
const PORT = process.env.PORT || process.env.GIJIROKU_PORT || 3456;
const MAX_RETRIES = 5;
const RETRY_DELAY = 5000; // 5秒

const app = express();
app.use(express.json({ limit: "10mb" }));
app.use(express.static(resolve(__dirname, "public")));

const upload = multer({
  storage: multer.memoryStorage(),
  limits: { fileSize: 200 * 1024 * 1024 },
});

// MIMEタイプ判定
const AUDIO_MIME = {
  mp3: "audio/mp3", wav: "audio/wav", m4a: "audio/mp4",
  mp4: "audio/mp4", ogg: "audio/ogg", webm: "audio/webm",
  flac: "audio/flac", aac: "audio/aac",
};

// ─── 音声圧縮（ffmpeg） ───
const COMPRESS_THRESHOLD = 5 * 1024 * 1024; // 5MB以上なら圧縮

function compressAudio(buffer, originalName) {
  return new Promise((resolve, reject) => {
    const ts = Date.now();
    const inputPath = `${TMP_DIR}/input-${ts}`;
    const outputPath = `${TMP_DIR}/output-${ts}.mp3`;

    writeFileSync(inputPath, buffer);

    // 32kbps モノラル MP3に圧縮（会議音声に十分な品質）
    execFile("ffmpeg", [
      "-i", inputPath,
      "-ac", "1",          // モノラル
      "-ab", "32k",        // 32kbps
      "-ar", "16000",      // 16kHz
      "-y",                // 上書き
      outputPath,
    ], { timeout: 120_000 }, (err) => {
      // 入力ファイル削除
      try { unlinkSync(inputPath); } catch {}

      if (err) {
        try { unlinkSync(outputPath); } catch {}
        reject(new Error(`ffmpeg圧縮エラー: ${err.message}`));
        return;
      }

      const compressed = readFileSync(outputPath);
      try { unlinkSync(outputPath); } catch {}
      resolve(compressed);
    });
  });
}

// APIキー取得ヘルパー
function getApiKey(req) {
  const key = req.headers["x-api-key"] || req.body?.apiKey;
  if (!key) throw new Error("Gemini APIキーが設定されていません。画面上部でAPIキーを入力してください。");
  return key;
}

// ─── Gemini File API: アップロード → URI取得 ───
async function uploadToGeminiFileAPI(apiKey, buffer, mimeType, displayName) {
  const startUrl = `https://generativelanguage.googleapis.com/upload/v1beta/files?key=${apiKey}`;
  const startRes = await fetch(startUrl, {
    method: "POST",
    headers: {
      "X-Goog-Upload-Protocol": "resumable",
      "X-Goog-Upload-Command": "start",
      "X-Goog-Upload-Header-Content-Length": buffer.length.toString(),
      "X-Goog-Upload-Header-Content-Type": mimeType,
      "Content-Type": "application/json",
    },
    body: JSON.stringify({ file: { displayName } }),
  });

  if (!startRes.ok) {
    const err = await startRes.text();
    throw new Error(`File API 開始エラー (${startRes.status}): ${err}`);
  }

  const uploadUrl = startRes.headers.get("X-Goog-Upload-URL");
  if (!uploadUrl) throw new Error("File API: Upload URLが取得できませんでした");

  const uploadRes = await fetch(uploadUrl, {
    method: "PUT",
    headers: {
      "Content-Length": buffer.length.toString(),
      "X-Goog-Upload-Offset": "0",
      "X-Goog-Upload-Command": "upload, finalize",
    },
    body: buffer,
  });

  if (!uploadRes.ok) {
    const err = await uploadRes.text();
    throw new Error(`File API アップロードエラー (${uploadRes.status}): ${err}`);
  }

  const uploadData = await uploadRes.json();
  const fileUri = uploadData.file?.uri;
  const fileName = uploadData.file?.name;
  if (!fileUri) throw new Error("File API: file URIが取得できませんでした");

  // ACTIVE状態までポーリング
  const checkUrl = `https://generativelanguage.googleapis.com/v1beta/${fileName}?key=${apiKey}`;
  for (let i = 0; i < 60; i++) {
    const checkRes = await fetch(checkUrl);
    const checkData = await checkRes.json();
    if (checkData.state === "ACTIVE") return fileUri;
    if (checkData.state === "FAILED") throw new Error("File API: ファイル処理に失敗しました");
    await new Promise((r) => setTimeout(r, 2000));
  }
  throw new Error("File API: ファイル処理がタイムアウトしました");
}

// ─── Gemini API共通: リトライ + フォールバック ───
async function callGeminiWithRetry(apiKey, model, body, timeoutMs = 60_000) {
  let lastError;
  for (let attempt = 1; attempt <= MAX_RETRIES; attempt++) {
    try {
      const url = `https://generativelanguage.googleapis.com/v1beta/models/${model}:generateContent?key=${apiKey}`;
      const res = await fetch(url, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(body),
        signal: AbortSignal.timeout(timeoutMs),
      });

      if (res.ok) {
        const data = await res.json();
        return data.candidates?.[0]?.content?.parts?.[0]?.text || "";
      }

      const errText = await res.text();

      // 503/429 はリトライ
      if ((res.status === 503 || res.status === 429) && attempt < MAX_RETRIES) {
        const wait = RETRY_DELAY * attempt;
        console.log(`  → ${model} ${res.status} (${attempt}/${MAX_RETRIES}), ${wait / 1000}秒後にリトライ...`);
        await new Promise((r) => setTimeout(r, wait));
        continue;
      }

      throw new Error(`Gemini API エラー (${res.status}, ${model}): ${errText}`);
    } catch (err) {
      lastError = err;
      // ネットワークエラー（fetch failed, other side closed等）もリトライ
      if (attempt < MAX_RETRIES) {
        const wait = RETRY_DELAY * attempt;
        console.log(`  → ${model} エラー: ${err.message} (${attempt}/${MAX_RETRIES}), ${wait / 1000}秒後にリトライ...`);
        await new Promise((r) => setTimeout(r, wait));
        continue;
      }
      throw err;
    }
  }
  throw lastError;
}

// gemini-3-flash-preview 固定（フォールバックなし）
async function callGeminiSmart(apiKey, body, timeoutMs = 60_000) {
  try {
    return await callGeminiWithRetry(apiKey, MODEL, body, timeoutMs);
  } catch (err) {
    console.log(`  → ${MODEL} 失敗, ${MODEL_FALLBACK} にフォールバック...`);
    try {
      return await callGeminiWithRetry(apiKey, MODEL_FALLBACK, body, timeoutMs);
    } catch (err2) {
      if (err2.message.includes("503") || err2.message.includes("UNAVAILABLE"))
        throw new Error("AIサーバーが現在混み合っています。しばらく時間を置いてから再度お試しください。");
      if (err2.message.includes("fetch failed") || err2.message.includes("other side closed"))
        throw new Error("AIサーバーとの接続が切れました。しばらく時間を置いてから再度お試しください。");
      if (err2.message.includes("timeout") || err2.message.includes("aborted"))
        throw new Error("AIサーバーからの応答がありませんでした。しばらく時間を置いてから再度お試しください。");
      throw err2;
    }
  }
}

// ─── Gemini API: テキストのみ ───
async function callGeminiText(apiKey, prompt) {
  return callGeminiSmart(apiKey, {
    contents: [{ parts: [{ text: prompt }] }],
    generationConfig: { temperature: 0.3, maxOutputTokens: 65536 },
  });
}

// ─── Gemini API: File URI参照で文字起こし ───
async function callGeminiWithFileUri(apiKey, fileUri, mimeType, prompt) {
  return callGeminiSmart(apiKey, {
    contents: [{
      parts: [
        { file_data: { mime_type: mimeType, file_uri: fileUri } },
        { text: prompt },
      ],
    }],
    generationConfig: { temperature: 0.1, maxOutputTokens: 65536 },
  }, 120_000); // 2分タイムアウト（リトライ5回で最大約12分）
}

// ─── APIキー検証 ───
app.post("/api/verify-key", async (req, res) => {
  try {
    const apiKey = req.body.apiKey;
    if (!apiKey) return res.status(400).json({ valid: false, error: "APIキーが空です" });

    const url = `https://generativelanguage.googleapis.com/v1beta/models?key=${apiKey}`;
    const checkRes = await fetch(url);
    if (checkRes.ok) {
      res.json({ valid: true });
    } else {
      res.json({ valid: false, error: "無効なAPIキーです" });
    }
  } catch (err) {
    res.json({ valid: false, error: err.message });
  }
});

// ─── ステップ1: 音声ファイル → 文字起こし ───
app.post("/api/transcribe", upload.single("audio"), async (req, res) => {
  req.setTimeout(600_000);
  res.setTimeout(600_000);

  try {
    const apiKey = req.headers["x-api-key"];
    if (!apiKey) return res.status(400).json({ error: "APIキーが設定されていません" });

    if (!req.file) return res.status(400).json({ error: "音声ファイルが指定されていません" });

    const file = req.file;
    let audioBuffer = file.buffer;
    let mimeType = file.mimetype;
    if (mimeType === "audio/mpeg") mimeType = "audio/mp3";
    const ext = file.originalname?.split(".").pop()?.toLowerCase();
    if (ext && AUDIO_MIME[ext]) mimeType = AUDIO_MIME[ext];

    const sizeMB = (file.size / (1024 * 1024)).toFixed(1);
    console.log(`文字起こし開始: ${file.originalname} (${mimeType}, ${sizeMB}MB)`);

    // 15MB以上なら自動圧縮
    if (file.size > COMPRESS_THRESHOLD) {
      console.log(`  → ${sizeMB}MB > 15MB — ffmpegで圧縮中...`);
      audioBuffer = await compressAudio(file.buffer, file.originalname);
      mimeType = "audio/mp3";
      const compressedMB = (audioBuffer.length / (1024 * 1024)).toFixed(1);
      console.log(`  → 圧縮完了: ${sizeMB}MB → ${compressedMB}MB`);
    }

    const fileUri = await uploadToGeminiFileAPI(apiKey, audioBuffer, mimeType, file.originalname);
    console.log(`  → アップロード完了, 文字起こし中...`);

    // 空レスポンス対策: 最大2回リトライ
    let transcript = "";
    for (let attempt = 1; attempt <= 3; attempt++) {
      transcript = await callGeminiWithFileUri(apiKey, fileUri, mimeType, TRANSCRIBE_PROMPT);
      if (transcript.trim().length > 10) break;
      console.log(`  → 空レスポンス (${attempt}/3), リトライ中...`);
      await new Promise((r) => setTimeout(r, 3000));
    }

    if (!transcript.trim()) {
      throw new Error("文字起こし結果が空でした。音声ファイルを確認するか、再度お試しください。");
    }

    console.log(`文字起こし完了: ${transcript.length}文字`);
    res.json({ transcript });
  } catch (err) {
    const detail = err.cause ? `${err.message} (${err.cause.message || err.cause})` : err.message;
    console.error("文字起こしエラー:", detail);
    res.status(500).json({ error: detail });
  }
});

// ─── ステップ2: 議事録生成 ───
app.post("/api/generate", async (req, res) => {
  try {
    const apiKey = getApiKey(req);
    const { transcript, memo, format, meetingTitle, participants, date, koujiInfo } = req.body;
    if (!transcript && !memo) {
      return res.status(400).json({ error: "文字起こしテキストまたはメモを入力してください" });
    }

    let prompt;
    if (format === "kouji") {
      prompt = buildKoujiPrompt({ transcript, memo, koujiInfo });
    } else {
      prompt = buildMinutesPrompt({ transcript, memo, meetingTitle, participants, date });
    }
    const result = await callGeminiText(apiKey, prompt);
    res.json({ minutes: result });
  } catch (err) {
    console.error("生成エラー:", err.message);
    res.status(500).json({ error: err.message });
  }
});

// ─── Word出力（議事録フォーマット） ───
app.post("/api/export/docx", async (req, res) => {
  const { markdown, meetingTitle } = req.body;
  if (!markdown) return res.status(400).json({ error: "Markdownが空です" });

  try {
    const paragraphs = markdownToDocxParagraphs(markdown);
    const doc = new Document({
      styles: {
        default: {
          document: {
            run: { font: "Yu Gothic", size: 21 }, // 10.5pt
            paragraph: { spacing: { line: 260, after: 0 } }, // 行間詰め
          },
        },
        paragraphStyles: [
          {
            id: "Heading1", name: "Heading 1",
            basedOn: "Normal", next: "Normal", quickFormat: true,
            run: { size: 32, bold: true, font: "Yu Gothic" }, // 16pt
            paragraph: {
              spacing: { before: 0, after: 40 },
              alignment: AlignmentType.CENTER,
              outlineLevel: 0,
            },
          },
          {
            id: "Heading2", name: "Heading 2",
            basedOn: "Normal", next: "Normal", quickFormat: true,
            run: { size: 24, bold: true, font: "Yu Gothic" }, // 12pt
            paragraph: {
              spacing: { before: 100, after: 20 },
              outlineLevel: 1,
            },
          },
          {
            id: "Heading3", name: "Heading 3",
            basedOn: "Normal", next: "Normal", quickFormat: true,
            run: { size: 22, bold: true, font: "Yu Gothic" }, // 11pt
            paragraph: {
              spacing: { before: 80, after: 0 },
              outlineLevel: 2,
            },
          },
        ],
      },
      numbering: {
        config: [{
          reference: "bullets",
          levels: [{
            level: 0,
            format: LevelFormat.BULLET,
            text: "\u2022",
            alignment: AlignmentType.LEFT,
            style: { paragraph: { indent: { left: 720, hanging: 360 } } },
          }],
        }],
      },
      sections: [{
        properties: {
          page: {
            size: { width: 11906, height: 16838 }, // A4
            margin: {
              top: 1985,  // 35mm
              right: 1701, // 30mm
              bottom: 1701, // 30mm
              left: 1701,  // 30mm
            },
          },
        },
        children: paragraphs,
      }],
    });

    const buffer = await Packer.toBuffer(doc);
    const dateStr = new Date().toISOString().slice(0, 10).replace(/-/g, "");
    const baseName = (meetingTitle || "議事録") + "_" + dateStr;
    if (IS_WINDOWS) {
      const outputPath = resolve("C:/Users/tekko/Desktop", baseName + ".docx");
      writeFileSync(outputPath, Buffer.from(buffer));
      res.json({ success: true, path: outputPath, filename: baseName + ".docx" });
    } else {
      const filename = baseName + ".docx";
      const enc = encodeURIComponent(filename);
      res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.wordprocessingml.document");
      res.setHeader("Content-Disposition", `attachment; filename="${enc}"; filename*=UTF-8''${enc}`);
      res.send(Buffer.from(buffer));
    }
  } catch (err) {
    console.error("Word生成エラー:", err.message);
    res.status(500).json({ error: `Word生成エラー: ${err.message}` });
  }
});

// ─── Word出力（公共工事 打合記録簿フォーマット） ───
app.post("/api/export/docx-kouji", async (req, res) => {
  const { markdown, koujiInfo = {} } = req.body;
  if (!markdown) return res.status(400).json({ error: "Markdownが空です" });

  try {
    // ヘルパー: セル生成（改行対応）
    const cellBorders = {
      top: { style: BorderStyle.SINGLE, size: 1 },
      bottom: { style: BorderStyle.SINGLE, size: 1 },
      left: { style: BorderStyle.SINGLE, size: 1 },
      right: { style: BorderStyle.SINGLE, size: 1 },
    };
    const cell = (text, opts = {}) => {
      const lines = (text || "").split("\n");
      const paragraphs = lines.map(line => new Paragraph({
        alignment: opts.align || AlignmentType.LEFT,
        children: [new TextRun({
          text: line,
          font: "Yu Gothic",
          size: opts.size || 18,
          bold: opts.bold || false,
        })],
        spacing: { before: 0, after: 0 },
      }));
      return new TableCell({
        width: opts.width ? { size: opts.width, type: WidthType.DXA } : undefined,
        verticalAlign: VerticalAlign.CENTER,
        children: paragraphs,
        columnSpan: opts.colSpan,
        rowSpan: opts.rowSpan,
        borders: cellBorders,
        shading: opts.shading,
      });
    };

    // ヘッダー情報テーブル
    const ki = koujiInfo;
    const headerTable = new Table({
      width: { size: 100, type: WidthType.PERCENTAGE },
      rows: [
        // 回数行
        new TableRow({
          children: [
            cell(ki.kaisu || "第　回", { bold: true, colSpan: 6, align: AlignmentType.CENTER, size: 20 }),
          ],
        }),
        // 年月日・場所
        new TableRow({
          children: [
            cell("年月日", { bold: true, width: 1400 }),
            cell(ki.date || "", { colSpan: 2, width: 3600 }),
            cell("場所", { bold: true, width: 1000 }),
            cell(ki.place || "", { colSpan: 2 }),
          ],
        }),
        // 業務名称
        new TableRow({
          children: [
            cell("業務の名称", { bold: true, width: 1400 }),
            cell(ki.koujiName || "", { colSpan: 2 }),
            cell("打合せ方式", { bold: true, width: 1400 }),
            cell(ki.method || "会議", { colSpan: 2 }),
          ],
        }),
        // 発注機関・会社名
        new TableRow({
          children: [
            cell("発注機関名\n担当部署名", { bold: true }),
            cell(ki.hatchuName || "", { colSpan: 2 }),
            cell("会社名\n（受注者側）", { bold: true }),
            cell(ki.juchuName || "", { colSpan: 2 }),
          ],
        }),
        // 出席者
        new TableRow({
          children: [
            cell("出席者", { bold: true }),
            cell("発注者側", { bold: true, width: 1000 }),
            cell(ki.hatchuMembers || ""),
            cell("", { bold: true }),
            cell("受注者側", { bold: true, width: 1000 }),
            cell(ki.juchuMembers || ""),
          ],
        }),
      ],
    });

    // 内容を発注者側/受注者側に分離してパース
    const contentRows = parseKoujiContent(markdown);

    // 発注者・受注者2列テーブル
    const colLabelRow = new TableRow({
      children: [
        cell("発注者側", { bold: true, align: AlignmentType.CENTER, width: 5000,
          shading: { fill: "E8E8E8" } }),
        cell("受注者側", { bold: true, align: AlignmentType.CENTER,
          shading: { fill: "E8E8E8" } }),
      ],
    });

    const bodyRows = contentRows.map(row => new TableRow({
      children: [
        new TableCell({
          width: { size: 5000, type: WidthType.DXA },
          verticalAlign: VerticalAlign.TOP,
          children: row.hatchu.map(line => new Paragraph({
            children: [new TextRun({ text: line, font: "Yu Gothic", size: 20 })],
            spacing: { before: 20, after: 20 },
            indent: line.startsWith("・") ? { left: 200, hanging: 200 } : undefined,
          })),
          borders: {
            top: { style: BorderStyle.SINGLE, size: 1 },
            bottom: { style: BorderStyle.SINGLE, size: 1 },
            left: { style: BorderStyle.SINGLE, size: 1 },
            right: { style: BorderStyle.SINGLE, size: 1 },
          },
        }),
        new TableCell({
          verticalAlign: VerticalAlign.TOP,
          children: row.juchu.map(line => new Paragraph({
            children: [new TextRun({ text: line, font: "Yu Gothic", size: 20 })],
            spacing: { before: 20, after: 20 },
            indent: line.startsWith("・") ? { left: 200, hanging: 200 } : undefined,
          })),
          borders: {
            top: { style: BorderStyle.SINGLE, size: 1 },
            bottom: { style: BorderStyle.SINGLE, size: 1 },
            left: { style: BorderStyle.SINGLE, size: 1 },
            right: { style: BorderStyle.SINGLE, size: 1 },
          },
        }),
      ],
    }));

    const contentTable = new Table({
      width: { size: 100, type: WidthType.PERCENTAGE },
      rows: [colLabelRow, ...bodyRows],
    });

    // 注記
    const note = new Paragraph({
      children: [new TextRun({
        text: "（注）1.内容欄には、打合せ議事内容を記載すること。",
        font: "Yu Gothic", size: 16, italics: true,
      })],
      spacing: { before: 100 },
    });

    const doc = new Document({
      styles: {
        default: {
          document: {
            run: { font: "Yu Gothic", size: 20 },
            paragraph: { spacing: { line: 240, after: 0 } },
          },
        },
      },
      sections: [{
        properties: {
          page: {
            size: { width: 16838, height: 11906 }, // A4横
            margin: { top: 1134, right: 1134, bottom: 1134, left: 1134 }, // 20mm
          },
        },
        children: [headerTable, new Paragraph({ spacing: { before: 200 } }), contentTable, note],
      }],
    });

    const buffer = await Packer.toBuffer(doc);
    const filename = (ki.koujiName || "打合記録") + ".docx";
    const encodedFilename = encodeURIComponent(filename);
    res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.wordprocessingml.document");
    res.setHeader("Content-Disposition", `attachment; filename="${encodedFilename}"; filename*=UTF-8''${encodedFilename}`);
    res.send(Buffer.from(buffer));
  } catch (err) {
    console.error("公共工事Word生成エラー:", err.message);
    res.status(500).json({ error: `Word生成エラー: ${err.message}` });
  }
});

// ─── Excel出力（公共工事 Excel COM転記） ───
const KOUJI_TEMPLATE = "X:\\公共工事議事録\\打合記録.xls";
const KOUJI_OUTPUT_DIR = "X:\\公共工事議事録";

app.post("/api/export/xls-kouji", async (req, res) => {
  if (!IS_WINDOWS) {
    return res.status(400).json({ error: "公共工事Excel出力はローカル環境（Windows）でのみ使用できます" });
  }
  const { markdown, koujiInfo = {} } = req.body;
  if (!markdown) return res.status(400).json({ error: "Markdownが空です" });

  try {
    const ki = koujiInfo;
    const dateStr = new Date().toISOString().slice(0, 10).replace(/-/g, "");
    const baseName = (ki.koujiName || "打合記録") + "_" + dateStr;
    const outputPath = resolve(KOUJI_OUTPUT_DIR, baseName + ".xlsx");

    // 協議内容をパース
    const entries = parseKoujiForExcel(markdown);

    // PowerShellスクリプトを生成
    const psScript = buildExcelComScript(ki, entries, KOUJI_TEMPLATE, outputPath);
    const psPath = resolve(TMP_DIR, `kouji-${Date.now()}.ps1`);
    writeFileSync(psPath, "\ufeff" + psScript, "utf-8"); // BOM付きUTF-8

    // PowerShell実行
    await new Promise((resolve, reject) => {
      execFile("powershell", [
        "-ExecutionPolicy", "Bypass",
        "-File", psPath,
      ], { timeout: 30_000 }, (err, stdout, stderr) => {
        try { unlinkSync(psPath); } catch {}
        if (err) {
          console.error("PowerShell error:", stderr || err.message);
          reject(new Error("Excel転記エラー: " + (stderr || err.message)));
          return;
        }
        console.log("Excel COM:", stdout.trim());
        resolve();
      });
    });

    // 保存パスをJSON で返す
    res.json({ success: true, path: outputPath, filename: baseName + ".xlsx" });
  } catch (err) {
    console.error("Excel転記エラー:", err.message);
    res.status(500).json({ error: err.message });
  }
});

// 協議内容テキスト → Excel転記用データ配列
// ルール:
//   「受注者、了承。」→ 受注者了承（I列）
//   「受注者、〜」→ 受注者発言（J列）
//   「了承。」（単独）→ 発注者了承（G列）
//   空行 → 行スキップ（元の書式を再現）
//   それ以外 → 発注者側（B列、インデント付きはC列）
function parseKoujiForExcel(text) {
  const entries = [];
  const lines = text.split("\n");
  let row = 14; // Row 14から開始（仮番号、後でページ分割時に再計算）

  for (const line of lines) {
    const trimmed = line.trim();

    // 空行 → 行を進める（元のExcelの空行を再現）
    if (!trimmed) {
      row++;
      continue;
    }

    // Markdown見出し（AI生成の場合）
    if (trimmed.startsWith("## ") || trimmed.startsWith("# ")) {
      entries.push({ row, col: "B", value: trimmed.replace(/^#+\s*/, "") });
      row++;
      continue;
    }

    // 「受注者、了承。〜」パターン（了承＋追加コメント）
    const ryoshoWithComment = trimmed.match(/^受注者[、,]\s*了承[。.]?\s*(.+)$/);
    if (ryoshoWithComment) {
      entries.push({ row, col: "I", value: "了承。" });
      entries.push({ row, col: "J", value: ryoshoWithComment[1].trim() });
      row++;
      continue;
    }

    // 「受注者、了承。」パターン
    if (/^受注者[、,]\s*了承[。.]?\s*$/.test(trimmed)) {
      entries.push({ row, col: "I", value: "了承。" });
      row++;
      continue;
    }

    // 「受注者、〜」パターン → 受注者発言（J列）
    const juchuMatch = trimmed.match(/^受注者[、,]\s*(.+)$/);
    if (juchuMatch) {
      entries.push({ row, col: "J", value: juchuMatch[1].trim() });
      row++;
      continue;
    }

    // Markdown受注者マーカー（AI生成の場合）
    if (trimmed.startsWith("→") || trimmed.startsWith("->")) {
      const t = trimmed.replace(/^→\s*|^->\s*/, "").trim();
      if (/^了承[。.]?\s*$/.test(t)) {
        entries.push({ row, col: "I", value: "了承。" });
      } else {
        entries.push({ row, col: "J", value: t });
      }
      row++;
      continue;
    }

    // 「了承。」単独（受注者マーカーなし）→ 発注者側の了承（G列）
    if (/^了承[。.]?\s*$/.test(trimmed)) {
      entries.push({ row, col: "G", value: "了承。" });
      row++;
      continue;
    }

    // インデント付き（タブや複数スペース始まり）→ C列
    if (line.startsWith("\t\t") || line.startsWith("　　") || /^\s{4,}/.test(line)) {
      entries.push({ row, col: "C", value: trimmed });
      row++;
      continue;
    }

    // 発注者側（デフォルト）→ B列
    entries.push({ row, col: "B", value: trimmed });
    row++;
  }

  return entries;
}

// PowerShellスクリプト生成（複数ページ対応）
function buildExcelComScript(ki, entries, templatePath, outputPath) {
  const esc = (s) => (s || "").replace(/'/g, "''");
  const CONTENT_START = 14;
  const CONTENT_END = 72;
  const ROWS_PER_PAGE = CONTENT_END - CONTENT_START + 1; // 59行

  // エントリをページごとに分割
  const pages = [[]];
  let pageRow = CONTENT_START;

  for (const e of entries) {
    if (pageRow > CONTENT_END) {
      pages.push([]);
      pageRow = CONTENT_START;
    }
    // 元のrow番号をページ内rowに変換
    pages[pages.length - 1].push({ ...e, row: pageRow });
    pageRow++;
  }

  const totalPages = pages.length;

  let ps = `
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

try {
  $wb = $excel.Workbooks.Open('${esc(templatePath)}')
  $ws1 = $wb.Sheets.Item(1)
`;

  // 追加シートが必要な場合、先にコピーして作成
  if (totalPages > 1) {
    ps += `
  # ${totalPages - 1}ページ分のシートを追加
  for ($i = 2; $i -le ${totalPages}; $i++) {
    $ws1.Copy([System.Reflection.Missing]::Value, $wb.Sheets.Item($wb.Sheets.Count)) | Out-Null
    $newSheet = $wb.Sheets.Item($wb.Sheets.Count)
    $newSheet.Name = '${esc(ki.kaisu || "第1回")}(' + $i + ')'
    # 追加ページのヘッダーは「追番ページ」表記に
    $newSheet.Range('M2').Value2 = '追番ページ ' + $i
    # 内容エリアクリア
    $newSheet.Range('A14:O72').ClearContents | Out-Null
  }
`;
  }

  // 各ページにデータ転記
  for (let p = 0; p < totalPages; p++) {
    const sheetIdx = p + 1;
    ps += `
  # --- ページ ${sheetIdx} ---
  $ws = $wb.Sheets.Item(${sheetIdx})
`;

    // 1ページ目はヘッダー転記 + 内容クリア
    if (p === 0) {
      ps += `
  # ヘッダー転記
  $ws.Range('A2').Value2 = '${esc(ki.kaisu || "第　　回")}'
  $ws.Range('D7').Value2 = '${esc(ki.date || "")}'
  $ws.Range('L7').Value2 = '${esc(ki.place || "")}'
  $ws.Range('D8').Value2 = '${esc(ki.koujiName || "")}'
  $ws.Range('L8').Value2 = '${esc(ki.method || "会議")}'
  $ws.Range('D9').Value2 = '${esc(ki.hatchuName || "")}'
  $ws.Range('L9').Value2 = '${esc(ki.juchuName || "")}'
  $ws.Range('F11').Value2 = '${esc(ki.hatchuMembers || "")}'
  $ws.Range('L11').Value2 = '${esc(ki.juchuMembers || "")}'
  $ws.Range('A14:O72').ClearContents | Out-Null
`;
    } else {
      // 追加ページもヘッダー情報を同じに
      ps += `
  $ws.Range('D7').Value2 = '${esc(ki.date || "")}'
  $ws.Range('D8').Value2 = '${esc(ki.koujiName || "")}'
`;
    }

    // 内容転記
    for (const e of pages[p]) {
      ps += `  $ws.Range('${e.col}${e.row}').Value2 = '${esc(e.value)}'\n`;
    }
  }

  ps += `
  # 別名保存（xlsx形式: FileFormat=51）
  $wb.SaveAs('${esc(outputPath)}', 51) | Out-Null
  $wb.Close()
  Write-Host 'OK: ${totalPages} pages'
} catch {
  Write-Host ('ERROR: ' + $_.Exception.Message)
} finally {
  $excel.Quit()
  [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
}
`;

  return ps;
}

// 公共工事用: Markdownを発注者/受注者のペアにパース
function parseKoujiContent(md) {
  const rows = [];
  let currentHatchu = [];
  let currentJuchu = [];
  const lines = md.split("\n");

  for (const line of lines) {
    const trimmed = line.trim();
    if (!trimmed || trimmed.startsWith("#") || trimmed === "---") {
      // セクション区切り: 蓄積があればpush
      if (currentHatchu.length || currentJuchu.length) {
        rows.push({ hatchu: currentHatchu.length ? currentHatchu : [""], juchu: currentJuchu.length ? currentJuchu : [""] });
        currentHatchu = [];
        currentJuchu = [];
      }
      // 見出し行は発注者側に追加
      if (trimmed.startsWith("#")) {
        const heading = trimmed.replace(/^#+\s*/, "");
        currentHatchu.push(heading);
      }
      continue;
    }

    // 【受注者】マーカーで分離
    if (trimmed.startsWith("【受注者】") || trimmed.startsWith("[受注者]")) {
      currentJuchu.push(trimmed.replace(/^【受注者】|^\[受注者\]/, "").trim());
    } else if (trimmed.startsWith("【発注者】") || trimmed.startsWith("[発注者]")) {
      currentHatchu.push(trimmed.replace(/^【発注者】|^\[発注者\]/, "").trim());
    } else if (trimmed.startsWith("→")) {
      // 受注者の応答
      currentJuchu.push(trimmed.slice(1).trim());
    } else {
      // デフォルトは発注者側
      currentHatchu.push(trimmed.replace(/^[-*]\s*/, ""));
    }
  }

  // 残りをpush
  if (currentHatchu.length || currentJuchu.length) {
    rows.push({ hatchu: currentHatchu.length ? currentHatchu : [""], juchu: currentJuchu.length ? currentJuchu : [""] });
  }

  return rows.length ? rows : [{ hatchu: ["（内容なし）"], juchu: [""] }];
}

// Markdown → docx Paragraphs 変換（議事録フォーマット）
function markdownToDocxParagraphs(md) {
  // 連続空行を1つにまとめる
  const lines = md.split("\n").reduce((acc, line) => {
    if (line.trim() === "" && acc.length > 0 && acc[acc.length - 1].trim() === "") return acc;
    acc.push(line);
    return acc;
  }, []);
  const paragraphs = [];

  for (const line of lines) {
    // # タイトル → 中央揃え・16pt太字
    if (line.startsWith("# ")) {
      paragraphs.push(new Paragraph({
        heading: HeadingLevel.HEADING_1,
        children: [new TextRun({ text: line.slice(2), bold: true, size: 32, font: "Yu Gothic" })],
      }));

    // ## 見出し → 14pt太字
    } else if (line.startsWith("## ")) {
      paragraphs.push(new Paragraph({
        heading: HeadingLevel.HEADING_2,
        children: [new TextRun({ text: line.slice(3), bold: true, size: 28, font: "Yu Gothic" })],
      }));

    // ### 小見出し → 12pt太字
    } else if (line.startsWith("### ")) {
      paragraphs.push(new Paragraph({
        heading: HeadingLevel.HEADING_3,
        children: [new TextRun({ text: line.slice(4), bold: true, size: 24, font: "Yu Gothic" })],
      }));

    // --- 水平線 → スキップ（ページを節約）
    } else if (line.startsWith("---")) {
      // 何もしない

    // - [ ] / - [x] チェックボックス
    } else if (line.match(/^- \[[ x]\] /)) {
      const checked = line.startsWith("- [x]");
      const text = line.replace(/^- \[[ x]\] /, "");
      paragraphs.push(new Paragraph({
        numbering: { reference: "bullets", level: 0 },
        children: [
          new TextRun({ text: checked ? "\u2611 " : "\u2610 ", font: "Yu Gothic" }),
          ...parseInline(text),
        ],
      }));

    // - 箇条書き
    } else if (line.startsWith("- ")) {
      paragraphs.push(new Paragraph({
        numbering: { reference: "bullets", level: 0 },
        children: parseInline(line.slice(2)),
      }));

    // * 箇条書き（アスタリスク）
    } else if (line.startsWith("* ")) {
      paragraphs.push(new Paragraph({
        numbering: { reference: "bullets", level: 0 },
        children: parseInline(line.slice(2)),
      }));

    // 空行 → スキップ
    } else if (line.trim() === "") {
      // 何もしない

    // 通常段落
    } else {
      paragraphs.push(new Paragraph({
        children: parseInline(line),
      }));
    }
  }
  return paragraphs;
}

// インラインパース（太字・通常テキスト）
function parseInline(text) {
  const runs = [];
  // **太字** と 【ラベル】 を処理
  const regex = /\*\*(.+?)\*\*|【(.+?)】/g;
  let lastIndex = 0;
  let match;
  while ((match = regex.exec(text)) !== null) {
    if (match.index > lastIndex) {
      runs.push(new TextRun({ text: text.slice(lastIndex, match.index), font: "Yu Gothic", size: 22 }));
    }
    if (match[1]) {
      // **太字**
      runs.push(new TextRun({ text: match[1], bold: true, font: "Yu Gothic", size: 22 }));
    } else if (match[2]) {
      // 【ラベル】 → 太字で表示
      runs.push(new TextRun({ text: `【${match[2]}】`, bold: true, font: "Yu Gothic", size: 22 }));
    }
    lastIndex = regex.lastIndex;
  }
  if (lastIndex < text.length) {
    runs.push(new TextRun({ text: text.slice(lastIndex), font: "Yu Gothic", size: 22 }));
  }
  if (runs.length === 0) {
    runs.push(new TextRun({ text, font: "Yu Gothic", size: 22 }));
  }
  return runs;
}

// ─── プロンプト ───
const TRANSCRIBE_PROMPT = `会議は日本語音声です。この音声ファイルの文字起こしをしてください。

## ルール
- 複数人が話しているので、話す人ごとに改行してください
- 話者が判別できる場合は「話者A:」「話者B:」のように区別してください
- 「えー」や「あー」など話す人特有の癖は文字にしないでください
- 改行、句読点を付けてもっと詳細に書いてください
- 発言者ごとに改行してください
- 聞き取れない箇所は「[聞き取り不可]」と記載してください`;

function buildMinutesPrompt({ transcript, memo, meetingTitle, participants, date }) {
  return `あなたは優秀な議事録作成アシスタントです。以下の会議情報から、構造化された議事録をMarkdown形式で作成してください。

## 会議情報
- 会議名: ${meetingTitle || "（未設定）"}
- 日時: ${date || new Date().toLocaleDateString("ja-JP")}
- 参加者: ${participants || "（未設定）"}

## 文字起こしテキスト
${transcript || "（なし）"}

## 手書きメモ・補足
${memo || "（なし）"}

## 出力フォーマット
以下の形式でMarkdownの議事録を作成してください：

# 議事録: {会議名}

**日時:** {日時}
**参加者:** {参加者}
**作成者:** 議事録AI

---

## 1. 議題・アジェンダ
- 話し合われたトピックを箇条書き

## 2. 議論の要約
- 各トピックごとに要点をまとめる
- 誰が何を発言したか（判別可能な場合）

## 3. 決定事項
- 会議で決まったことを明確に列挙

## 4. アクションアイテム
- [ ] 担当者: タスク内容（期限があれば記載）

## 5. 次回予定
- 次回の会議予定や持ち越し事項

---

### ルール
- 文字起こしテキストは誤認識を含む可能性があるので、文脈から適切に補正してください
- 手書きメモがある場合は、文字起こしテキストと照合して正確性を高めてください
- 不明な点は「※要確認」と注記してください
- **省略せず、できる限り詳細に記述してください。発言内容・議論の経緯・各参加者の意見を漏らさず記録してください**
- ただし無駄な空行は入れず、コンパクトにまとめてください。箇条書きを活用して簡潔に整理してください
- 各トピックについて、誰がどのような発言をしたかを具体的に書いてください
- 議論の流れ（賛成・反対・質問・回答）を時系列で記録してください
- ビジネス文書として適切なトーンで記述してください
- **「承知いたしました」「以下に〜」などの前置き・挨拶文は一切不要です。いきなり「# 議事録:」から始めてください**`;
}

function buildKoujiPrompt({ transcript, memo, koujiInfo = {} }) {
  const ki = koujiInfo;
  return `あなたは公共工事の打合記録簿を作成するアシスタントです。以下の打合せ内容から、発注者側と受注者側の発言を分離した打合記録を作成してください。

## 打合せ情報
- 業務名称: ${ki.koujiName || "（未設定）"}
- 日時: ${ki.date || "（未設定）"}
- 場所: ${ki.place || "（未設定）"}
- 発注機関: ${ki.hatchuName || "（未設定）"}
- 受注者: ${ki.juchuName || "（未設定）"}
- 出席者（発注者側）: ${ki.hatchuMembers || "（未設定）"}
- 出席者（受注者側）: ${ki.juchuMembers || "（未設定）"}
- 回数: ${ki.kaisu || "（未設定）"}

## 文字起こしテキスト
${transcript || "（なし）"}

## 手書きメモ・補足
${memo || "（なし）"}

## 出力フォーマット（重要）

以下のルールに厳密に従ってください：

### 構成ルール
1. **発注者側の発言**はそのまま記述する（行頭にマーカーなし）
2. **受注者側の発言・回答**は行頭に「受注者、」を付ける
3. **受注者が了承する場合**は「受注者、了承。」と書く。追加コメントがあれば「受注者、了承。追加コメント」と同じ行に書く
4. **発注者が了承する場合**は「了承。」とだけ書く（「受注者、」は付けない）
5. 発注者の指示・質問に対して受注者が回答している場合は、対応がわかるように近い位置に配置する
6. 空行でセクションを区切る（見出し行の前には空行を入れる）
7. インデント（継続行）は元のテキストの構造を維持する

### 話者の判別方法
- メモに「話者Aは受注者」「話者Bは発注者」等の指定があれば、それに従って振り分ける
- 指定がない場合は文脈から判断する：指示・依頼・要望 → 発注者、回答・検討・了承 → 受注者
- 出席者情報も参考にする（発注者側出席者の名前が話者名に含まれていれば発注者）

### 出力例

${ki.hatchuName || "発注者"}にて打ち合わせを行う。


以下、協議事項
	造成図面案の説明がある。
	・先日の現地立会で決めた起点より30度ラインの内側で
		設計を進めていただきたい。
	・初回の資料では9m×17mでしたが、8m×15mに変更してください。
	・床面積120㎡は確保してほしい。
受注者、了承。


	建物位置の変更に伴い変更点の説明がある。
	・南面は片引きまたは両引き戸とし、シャッターは手動のシャッターに変更。
受注者、・シャッターは3m×3m＝9㎡までなら制作可能です。
	・内部に3箇所、電光掲示板用の充電用コンセントを設置していただきたい。
受注者、・差込み形状を後日教えてください。

受注者、・造成図面ができ敷地が決まりましたので、
受注者、近日中にスウェーデンサウンディング試験を行います。
	了承。
受注者、・特記仕様書の配布をお願いします。

---

### 注意事項
- 文字起こしの誤認識は文脈から適切に補正してください
- 手書きメモがある場合は照合して正確性を高めてください
- 建築・土木の専門用語は正確に使用してください
- 不明な点は「※要確認」と注記してください
- 省略せず、協議内容を漏らさず記録してください
- 空行の位置も元のテキストの構造を尊重してください
- **「承知いたしました」「以下に〜」などの前置き・挨拶文は一切不要です。いきなり最初の内容から始めてください**`;
}

// ─── サーバー起動 ───
const server = app.listen(PORT, () => {
  console.log(`\n🎙️  議事録AI サーバー起動`);
  console.log(`   http://localhost:${PORT}\n`);
});

server.timeout = 600_000;
server.keepAliveTimeout = 600_000;
