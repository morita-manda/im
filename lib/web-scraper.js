const axios = require("axios");
const cheerio = require("cheerio");

/**
 * RGBをHEX文字列（6桁大文字）に変換
 */
function rgbToHex(r, g, b) {
  return [r, g, b]
    .map((v) => Math.min(255, Math.max(0, v)).toString(16).padStart(2, "0"))
    .join("")
    .toUpperCase();
}

/**
 * 色が白・黒・グレー系かどうか判定（除外対象）
 */
function isNeutralColor(hex) {
  const r = parseInt(hex.slice(0, 2), 16);
  const g = parseInt(hex.slice(2, 4), 16);
  const b = parseInt(hex.slice(4, 6), 16);
  const brightness = (r * 299 + g * 587 + b * 114) / 1000;
  const saturation = Math.max(r, g, b) - Math.min(r, g, b);
  return brightness < 20 || brightness > 235 || saturation < 30;
}

/**
 * HPからブランドカラー候補を出現頻度順に抽出し、最多使用色も返す
 */
function extractBrandColors($, html) {
  const counts = {};

  function count(hex, weight = 1) {
    hex = hex.toUpperCase();
    if (isNeutralColor(hex)) return;
    counts[hex] = (counts[hex] || 0) + weight;
  }

  // 1. meta theme-color（重みを大きくして優先）
  const metaTheme =
    $('meta[name="theme-color"]').attr("content") ||
    $('meta[name="msapplication-TileColor"]').attr("content");
  if (metaTheme) {
    const m = metaTheme.match(/^#?([0-9a-fA-F]{6})$/);
    if (m) count(m[1], 20);
  }

  // 2. CSSから全HEXカラーを抽出してカウント
  const styleTexts = [];
  $("style").each((_, el) => styleTexts.push($(el).text()));
  const allCss = styleTexts.join("\n") + "\n" + html;

  // CSS変数（--primary / --brand / --accent 等）は重みを追加
  const varPattern = /--(primary|brand|accent|main|theme|key|base)[^:]*:\s*#([0-9a-fA-F]{6})/gi;
  let m;
  while ((m = varPattern.exec(allCss)) !== null) count(m[2], 5);

  // 全HEXカラーをカウント（inline styleも含む）
  const hexPattern = /#([0-9a-fA-F]{6})\b/g;
  while ((m = hexPattern.exec(allCss)) !== null) count(m[1], 1);

  // 頻度順にソート
  const sorted = Object.entries(counts)
    .sort((a, b) => b[1] - a[1])
    .map(([hex]) => hex);

  return sorted.slice(0, 15);
}

/**
 * 指定URLのテキストを取得する（サブページ用）
 */
async function fetchPageText(url) {
  try {
    const response = await axios.get(url, {
      timeout: 3000,
      headers: { "User-Agent": "Mozilla/5.0 (compatible; IM-Generator/1.0)" },
      maxRedirects: 3,
    });
    const $ = cheerio.load(response.data);
    $("script, style, nav, footer, header").remove();
    return $("body").text().replace(/\s+/g, " ").trim().slice(0, 3000);
  } catch {
    return "";
  }
}

/**
 * HPからテキスト情報とブランドカラーを取得する
 * @param {string} url
 * @returns {Promise<{title, description, bodyText, brandColors, url}>}
 */
async function scrapeWeb(url) {
  // URL正規化
  if (!url.startsWith("http")) url = "https://" + url;
  const baseUrl = url.replace(/\/$/, "");

  const response = await axios.get(url, {
    timeout: 15000,
    headers: {
      "User-Agent": "Mozilla/5.0 (compatible; IM-Generator/1.0)",
    },
    maxRedirects: 5,
  });

  const html = response.data;
  const $ = cheerio.load(html);

  // metaタグ・タイトル取得
  const title = $("title").text().trim();
  const description =
    $('meta[name="description"]').attr("content") ||
    $('meta[property="og:description"]').attr("content") ||
    "";

  // ブランドカラー抽出（script/style除去前に実行）
  const brandColors = extractBrandColors($, html);
  const dominantColor = brandColors[0] || null;

  // 本文テキスト抽出（script/style除去）
  $("script, style, nav, footer, header").remove();
  const bodyText = $("body").text().replace(/\s+/g, " ").trim().slice(0, 3000);

  // 会社概要ページを追加取得（代表者名・住所などを補完）
  const companyCandidates = [
    "/company", "/about", "/company/profile", "/corporate",
    "/company/", "/about/", "/about-us", "/aboutus",
    "/company/overview", "/company/about", "/corporate/profile",
    "/info", "/profile", "/outline",
  ];
  let companyPageText = "";
  for (const path of companyCandidates) {
    const text = await fetchPageText(baseUrl + path);
    if (text.length > 100) {
      companyPageText = text;
      break;
    }
  }

  return { title, description, bodyText, companyPageText, brandColors, dominantColor, url };
}

/**
 * 指定URLのトップページをスクリーンショット撮影し、一時ファイルパスを返す
 * @param {string} url
 * @param {string} outPath - 保存先ファイルパス
 * @returns {Promise<string|null>} 成功時はoutPath、失敗時はnull
 */
async function captureScreenshot(url, outPath) {
  let browser;
  try {
    const puppeteer = require("puppeteer");
    browser = await puppeteer.launch({
      headless: "new",
      executablePath: process.env.PUPPETEER_EXECUTABLE_PATH || undefined,
      args: ["--no-sandbox", "--disable-setuid-sandbox", "--disable-dev-shm-usage", "--disable-gpu"],
    });
    const page = await browser.newPage();
    await page.setViewport({ width: 1280, height: 900 });
    await page.goto(url, { waitUntil: "networkidle2", timeout: 20000 });
    await page.screenshot({ path: outPath, type: "png", clip: { x: 0, y: 0, width: 1280, height: 900 } });
    return outPath;
  } catch (e) {
    console.warn("[Screenshot] 撮影失敗:", e.message);
    return null;
  } finally {
    if (browser) await browser.close().catch(() => {});
  }
}

module.exports = scrapeWeb;
module.exports.captureScreenshot = captureScreenshot;
