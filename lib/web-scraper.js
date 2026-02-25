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
 * HPからブランドカラー候補を抽出する
 */
function extractBrandColors($, html) {
  const seen = new Set();
  const results = [];

  function add(hex) {
    hex = hex.toUpperCase();
    if (seen.has(hex)) return;
    // 白・黒・グレー系を除外
    const r = parseInt(hex.slice(0, 2), 16);
    const g = parseInt(hex.slice(2, 4), 16);
    const b = parseInt(hex.slice(4, 6), 16);
    const brightness = (r * 299 + g * 587 + b * 114) / 1000;
    const saturation = Math.max(r, g, b) - Math.min(r, g, b);
    if (brightness < 20 || brightness > 235 || saturation < 30) return;
    seen.add(hex);
    results.push(hex);
  }

  // 1. meta theme-color（最優先）
  const metaTheme =
    $('meta[name="theme-color"]').attr("content") ||
    $('meta[name="msapplication-TileColor"]').attr("content");
  if (metaTheme) {
    const m = metaTheme.match(/^#?([0-9a-fA-F]{6})$/);
    if (m) add(m[1]);
  }

  // 2. CSSからHEXカラーを抽出
  const styleTexts = [];
  $("style").each((_, el) => styleTexts.push($(el).text()));
  const allCss = styleTexts.join("\n") + "\n" + html;

  // CSS変数（--primary / --brand / --accent / --main / --color-primary 等）
  const varPattern = /--(primary|brand|accent|main|theme|key|base)[^:]*:\s*#([0-9a-fA-F]{6})/gi;
  let m;
  while ((m = varPattern.exec(allCss)) !== null) add(m[2]);

  // ヘッダー・ナビ・ボタン周辺のHEXカラー
  const keywordPattern =
    /(?:header|navbar|nav|\.btn|button|\.hero|\.top|\.cta|\.primary|background)[^{]*\{[^}]*?#([0-9a-fA-F]{6})/gi;
  while ((m = keywordPattern.exec(allCss)) !== null) add(m[1]);

  // 全HEXカラー（補足用、上限まで）
  const hexPattern = /#([0-9a-fA-F]{6})\b/g;
  while ((m = hexPattern.exec(allCss)) !== null && results.length < 20) add(m[1]);

  return results.slice(0, 15);
}

/**
 * HPからテキスト情報とブランドカラーを取得する
 * @param {string} url
 * @returns {Promise<{title, description, bodyText, brandColors, url}>}
 */
async function scrapeWeb(url) {
  // URL正規化
  if (!url.startsWith("http")) url = "https://" + url;

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

  // 本文テキスト抽出（script/style除去）
  $("script, style, nav, footer, header").remove();
  const bodyText = $("body").text().replace(/\s+/g, " ").trim().slice(0, 5000);

  return { title, description, bodyText, brandColors, url };
}

module.exports = scrapeWeb;
