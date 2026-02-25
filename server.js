require("dotenv").config();
const express = require("express");
const multer = require("multer");
const path = require("path");
const fs = require("fs");
const os = require("os");

const parsePDF = require("./lib/pdf-parser");
const scrapeWeb = require("./lib/web-scraper");
const processWithClaude = require("./lib/claude-processor");
const generateIM = require("./lib/pptx-generator");

const app = express();
const PORT = process.env.PORT || 3000;

// 一時ファイル保存先
const upload = multer({
  dest: os.tmpdir(),
  limits: { fileSize: 20 * 1024 * 1024 }, // 20MB
  fileFilter: (req, file, cb) => {
    if (file.mimetype === "application/pdf") cb(null, true);
    else cb(new Error("PDFファイルのみアップロード可能です"));
  },
});

app.use(express.json());
app.use(express.static(path.join(__dirname, "public")));

// IM生成エンドポイント
app.post("/api/generate", upload.array("pdfs", 3), async (req, res) => {
  const tmpFiles = (req.files || []).map((f) => f.path);

  try {
    const { url, companyName, repName, reason, price, scheme, managementIntent, empFull, empPart } = req.body;
    const shareholders = JSON.parse(req.body.shareholdersJson || "[]");

    if (!url) {
      return res.status(400).json({ error: "会社のURLは必須です" });
    }

    // 1. PDF テキスト抽出（各PDFを個別に配列へ）
    const pdfTexts = [];
    for (let i = 0; i < tmpFiles.length; i++) {
      const text = await parsePDF(tmpFiles[i]);
      console.log(`[PDF ${i + 1}] 抽出文字数: ${text.length}文字 / 先頭: ${text.slice(0, 100).replace(/\n/g, " ")}`);
      pdfTexts.push(text);
    }

    // 2. HP情報取得
    let webInfo = {};
    try {
      webInfo = await scrapeWeb(url);
    } catch (e) {
      console.warn("Webスクレイピング失敗（続行）:", e.message);
    }

    // 3. Claude APIで構造化
    const structured = await processWithClaude({
      pdfFilePaths: tmpFiles,  // Document API用（スキャンPDF対応）
      pdfTexts,                // メイン生成の会社情報抽出用
      webInfo,
      userInput: { url, companyName, repName, reason, price, scheme, managementIntent },
      shareholders,
      employeeBreakdown: { full: empFull, part: empPart },
    });

    // 4. PPTX生成（ユーザー入力の株主・役員情報を直接セット）
    if (shareholders.length) structured.shareholders = shareholders;
    structured.employeeBreakdown = { full: empFull || "", part: empPart || "" };
    const buffer = await generateIM(structured);

    // 5. ファイル返却
    const fileName = `企業概要書_${structured.company.name || "会社名"}.pptx`;
    res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.presentationml.presentation");
    res.setHeader("Content-Disposition", `attachment; filename*=UTF-8''${encodeURIComponent(fileName)}`);
    res.send(buffer);
  } catch (err) {
    console.error("生成エラー:", err);
    res.status(500).json({ error: err.message || "生成中にエラーが発生しました" });
  } finally {
    // 一時ファイル削除
    for (const f of tmpFiles) {
      fs.unlink(f, () => {});
    }
  }
});

app.listen(PORT, () => {
  console.log(`✅ IM Generator 起動中: http://localhost:${PORT}`);
});
