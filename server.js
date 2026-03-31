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
import { writeFileSync, readFileSync, unlinkSync, mkdirSync } from "fs";
import {
  Document, Packer, Paragraph, TextRun, HeadingLevel,
  AlignmentType, BorderStyle, LevelFormat,
} from "docx";

const __dirname = dirname(fileURLToPath(import.meta.url));
const TMP_DIR = resolve(__dirname, "tmp");
mkdirSync(TMP_DIR, { recursive: true });

const MODEL = "gemini-3-flash-preview";
const MODEL_FALLBACK = "gemini-2.5-flash";
const PORT = process.env.GIJIROKU_PORT || 3456;
const MAX_RETRIES = 3;
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

// Pro → Flash フォールバック
async function callGeminiSmart(apiKey, body, timeoutMs = 60_000) {
  try {
    return await callGeminiWithRetry(apiKey, MODEL, body, timeoutMs);
  } catch (err) {
    console.log(`  → ${MODEL} 失敗, ${MODEL_FALLBACK} にフォールバック...`);
    return await callGeminiWithRetry(apiKey, MODEL_FALLBACK, body, timeoutMs);
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
  }, 600_000); // 10分タイムアウト
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
    const { transcript, memo, meetingTitle, participants, date } = req.body;
    if (!transcript && !memo) {
      return res.status(400).json({ error: "文字起こしテキストまたはメモを入力してください" });
    }

    const prompt = buildMinutesPrompt({ transcript, memo, meetingTitle, participants, date });
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
    const filename = encodeURIComponent(meetingTitle || "議事録") + ".docx";
    res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.wordprocessingml.document");
    res.setHeader("Content-Disposition", `attachment; filename="${filename}"`);
    res.send(Buffer.from(buffer));
  } catch (err) {
    console.error("Word生成エラー:", err.message);
    res.status(500).json({ error: `Word生成エラー: ${err.message}` });
  }
});

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

// ─── サーバー起動 ───
const server = app.listen(PORT, () => {
  console.log(`\n🎙️  議事録AI サーバー起動`);
  console.log(`   http://localhost:${PORT}\n`);
});

server.timeout = 600_000;
server.keepAliveTimeout = 600_000;
