/**
 * 議事録AI - フロントエンド
 * 音声ファイルアップロード → Gemini文字起こし → 議事録生成
 */

let selectedFile = null;

// --- APIキー管理 ---
const API_KEY_STORAGE = "gijiroku_api_key";

function getApiKey() {
  return document.getElementById("apiKey").value.trim();
}

function requireApiKey() {
  const key = getApiKey();
  if (!key) {
    alert("Gemini APIキーを入力してください。");
    document.getElementById("apiKey").focus();
    return null;
  }
  return key;
}

function toggleApiKeyDialog() {
  const overlay = document.getElementById("apiKeyOverlay");
  overlay.classList.toggle("hidden");
}

function closeApiKeyDialog(e) {
  if (e.target === e.currentTarget) {
    document.getElementById("apiKeyOverlay").classList.add("hidden");
  }
}

function updateApiKeyLabel() {
  const key = getApiKey();
  const label = document.getElementById("apiKeyLabel");
  const btn = document.getElementById("btnApiKeyToggle");
  if (key) {
    label.textContent = "APIキー設定済み";
    btn.classList.add("configured");
  } else {
    label.textContent = "APIキー未設定";
    btn.classList.remove("configured");
  }
}

async function verifyApiKey() {
  const key = getApiKey();
  if (!key) {
    alert("APIキーを入力してください。");
    return;
  }

  const status = document.getElementById("apiKeyStatus");
  status.textContent = "確認中...";
  status.className = "apikey-status";

  try {
    const res = await fetch("/api/verify-key", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ apiKey: key }),
    });
    const data = await res.json();
    if (data.valid) {
      status.textContent = "有効なAPIキーです";
      status.className = "apikey-status valid";
      localStorage.setItem(API_KEY_STORAGE, key);
      updateApiKeyLabel();
      // 1秒後にダイアログを閉じる
      setTimeout(() => document.getElementById("apiKeyOverlay").classList.add("hidden"), 1000);
    } else {
      status.textContent = "無効なAPIキー: " + (data.error || "");
      status.className = "apikey-status invalid";
    }
  } catch (err) {
    status.textContent = "確認エラー: " + err.message;
    status.className = "apikey-status invalid";
  }
}

// --- ファイルアップロード ---
function initUpload() {
  const dropZone = document.getElementById("dropZone");
  const fileInput = document.getElementById("audioFile");

  dropZone.addEventListener("click", () => fileInput.click());

  dropZone.addEventListener("dragover", (e) => {
    e.preventDefault();
    dropZone.classList.add("drag-over");
  });

  dropZone.addEventListener("dragleave", () => {
    dropZone.classList.remove("drag-over");
  });

  dropZone.addEventListener("drop", (e) => {
    e.preventDefault();
    dropZone.classList.remove("drag-over");
    const file = e.dataTransfer.files[0];
    if (file && file.type.startsWith("audio/")) {
      handleFile(file);
    } else {
      alert("音声ファイルを選択してください。");
    }
  });

  fileInput.addEventListener("change", () => {
    if (fileInput.files[0]) {
      handleFile(fileInput.files[0]);
    }
  });
}

function handleFile(file) {
  selectedFile = file;
  document.getElementById("fileName").textContent = file.name;
  document.getElementById("fileSize").textContent = formatSize(file.size);
  document.getElementById("fileInfo").classList.remove("hidden");
  document.getElementById("dropZone").classList.add("has-file");
  document.getElementById("btnTranscribe").disabled = false;
}

function removeFile() {
  selectedFile = null;
  document.getElementById("audioFile").value = "";
  document.getElementById("fileInfo").classList.add("hidden");
  document.getElementById("dropZone").classList.remove("has-file");
  document.getElementById("btnTranscribe").disabled = true;
}

function formatSize(bytes) {
  if (bytes < 1024) return bytes + " B";
  if (bytes < 1024 * 1024) return (bytes / 1024).toFixed(1) + " KB";
  return (bytes / (1024 * 1024)).toFixed(1) + " MB";
}

// --- ステップ1: 文字起こし ---
async function transcribeAudio() {
  const apiKey = requireApiKey();
  if (!apiKey) return;

  if (!selectedFile) {
    alert("音声ファイルを選択してください。");
    return;
  }

  const overlay = showLoading("Geminiが音声を文字起こし中...<br><small>長い音声は数分かかります</small>");
  const btn = document.getElementById("btnTranscribe");
  btn.disabled = true;

  try {
    const formData = new FormData();
    formData.append("audio", selectedFile);

    const res = await fetch("/api/transcribe", {
      method: "POST",
      headers: { "X-Api-Key": apiKey },
      body: formData,
      signal: AbortSignal.timeout(600_000),
    });

    const data = await res.json();
    if (!res.ok) throw new Error(data.error);

    const section = document.getElementById("transcriptSection");
    section.classList.remove("hidden");
    document.getElementById("transcript").value = data.transcript;
    section.scrollIntoView({ behavior: "smooth" });
  } catch (err) {
    alert(`文字起こしエラー: ${err.message}`);
  } finally {
    overlay.remove();
    btn.disabled = false;
  }
}

// --- ステップ2: 議事録生成 ---
async function generateMinutes() {
  const apiKey = requireApiKey();
  if (!apiKey) return;

  const transcript = document.getElementById("transcript").value.trim();
  const memo = document.getElementById("memo").value.trim();
  const meetingTitle = document.getElementById("meetingTitle").value.trim();
  const participants = document.getElementById("participants").value.trim();
  const dateInput = document.getElementById("meetingDate").value;

  if (!transcript && !memo) {
    alert("文字起こしテキストまたはメモを入力してください。");
    return;
  }

  let date = "";
  if (dateInput) {
    const d = new Date(dateInput);
    date = d.toLocaleString("ja-JP", {
      year: "numeric", month: "long", day: "numeric",
      hour: "2-digit", minute: "2-digit",
    });
  }

  const overlay = showLoading("Geminiが議事録を生成中...");
  const btn = document.getElementById("btnGenerate");
  btn.disabled = true;

  try {
    const res = await fetch("/api/generate", {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        "X-Api-Key": apiKey,
      },
      body: JSON.stringify({ transcript, memo, meetingTitle, participants, date }),
    });

    const data = await res.json();
    if (!res.ok) throw new Error(data.error);

    showResult(data.minutes);
  } catch (err) {
    alert(`エラー: ${err.message}`);
  } finally {
    overlay.remove();
    btn.disabled = false;
  }
}

// --- ローディング ---
function showLoading(text) {
  const overlay = document.createElement("div");
  overlay.className = "loading-overlay";
  overlay.innerHTML = `<div class="spinner"></div><p class="loading-text">${text}</p>`;
  document.body.appendChild(overlay);
  return overlay;
}

// --- 結果表示 ---
let currentResult = "";

function showResult(markdown) {
  currentResult = markdown;
  const section = document.getElementById("resultSection");
  section.classList.remove("hidden");

  document.getElementById("resultPreview").innerHTML = renderMarkdown(markdown);
  document.getElementById("resultEdit").value = markdown;

  // プレビュータブに戻す
  switchTab("preview");
  section.scrollIntoView({ behavior: "smooth" });
}

function switchTab(mode) {
  const preview = document.getElementById("resultPreview");
  const edit = document.getElementById("resultEdit");
  const tabPreview = document.getElementById("tabPreview");
  const tabEdit = document.getElementById("tabEdit");

  if (mode === "edit") {
    preview.classList.add("hidden");
    edit.classList.remove("hidden");
    edit.value = currentResult;
    tabPreview.classList.remove("active");
    tabEdit.classList.add("active");
  } else {
    // 編集内容を反映
    if (!edit.classList.contains("hidden")) {
      currentResult = edit.value;
      preview.innerHTML = renderMarkdown(currentResult);
    }
    edit.classList.add("hidden");
    preview.classList.remove("hidden");
    tabEdit.classList.remove("active");
    tabPreview.classList.add("active");
  }
}

function renderMarkdown(md) {
  return md
    .replace(/```[\s\S]*?```/g, (m) => `<pre><code>${m.slice(3, -3).trim()}</code></pre>`)
    .replace(/^### (.+)$/gm, "<h3>$1</h3>")
    .replace(/^## (.+)$/gm, "<h2>$1</h2>")
    .replace(/^# (.+)$/gm, "<h1>$1</h1>")
    .replace(/^---$/gm, "<hr>")
    .replace(/\*\*(.+?)\*\*/g, "<strong>$1</strong>")
    .replace(/^- \[ \] (.+)$/gm, '<li style="list-style:none"><input type="checkbox" disabled> $1</li>')
    .replace(/^- \[x\] (.+)$/gm, '<li style="list-style:none"><input type="checkbox" checked disabled> $1</li>')
    .replace(/^- (.+)$/gm, "<li>$1</li>")
    .replace(/\n{2,}/g, "</p><p>")
    .replace(/\n/g, "<br>")
    .replace(/^/, "<p>")
    .replace(/$/, "</p>")
    .replace(/(<li>.*?<\/li>(?:<br>)?)+/g, (m) => `<ul>${m.replace(/<br>/g, "")}</ul>`);
}

function copyResult() {
  navigator.clipboard.writeText(currentResult).then(() => {
    const btn = event.target;
    const orig = btn.textContent;
    btn.textContent = "コピー済み";
    setTimeout(() => btn.textContent = orig, 1500);
  });
}

async function downloadTranscriptDocx() {
  const transcript = document.getElementById("transcript").value.trim();
  if (!transcript) {
    alert("文字起こし結果がありません。");
    return;
  }
  const overlay = showLoading("Wordファイルを生成中...");
  try {
    const meetingTitle = document.getElementById("meetingTitle").value.trim() || "文字起こし";
    const res = await fetch("/api/export/docx", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ markdown: transcript, meetingTitle: meetingTitle + "_文字起こし" }),
    });
    if (!res.ok) {
      const err = await res.json();
      throw new Error(err.error);
    }
    const blob = await res.blob();
    const dateStr = new Date().toISOString().slice(0, 10);
    triggerDownload(blob, `${meetingTitle}_文字起こし_${dateStr}.docx`);
  } catch (err) {
    alert(`Word生成エラー: ${err.message}`);
  } finally {
    overlay.remove();
  }
}

async function downloadDocx() {
  if (!currentResult) return;
  const overlay = showLoading("Wordファイルを生成中...");
  try {
    const meetingTitle = document.getElementById("meetingTitle").value.trim() || "議事録";
    const res = await fetch("/api/export/docx", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ markdown: currentResult, meetingTitle }),
    });
    if (!res.ok) {
      const err = await res.json();
      throw new Error(err.error);
    }
    const blob = await res.blob();
    const dateStr = new Date().toISOString().slice(0, 10);
    triggerDownload(blob, `${meetingTitle}_${dateStr}.docx`);
  } catch (err) {
    alert(`Word生成エラー: ${err.message}`);
  } finally {
    overlay.remove();
  }
}

function triggerDownload(blob, filename) {
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = filename;
  a.click();
  URL.revokeObjectURL(url);
}

// --- 初期化 ---
document.addEventListener("DOMContentLoaded", () => {
  initUpload();

  // 保存済みAPIキーを復元
  const savedKey = localStorage.getItem(API_KEY_STORAGE);
  if (savedKey) {
    document.getElementById("apiKey").value = savedKey;
  }
  updateApiKeyLabel();

  // 日時デフォルト
  const now = new Date();
  const offset = now.getTimezoneOffset();
  const local = new Date(now.getTime() - offset * 60000);
  document.getElementById("meetingDate").value = local.toISOString().slice(0, 16);
});
