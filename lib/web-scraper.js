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
 * 指定URLの会社概要ページを取得し、テキストと構造化データ({label,value}[])を返す
 */
async function fetchCompanyPage(url) {
  try {
    const response = await axios.get(url, {
      timeout: 5000,
      headers: { "User-Agent": "Mozilla/5.0 (compatible; IM-Generator/1.0)" },
      maxRedirects: 3,
    });
    const $ = cheerio.load(response.data);
    $("script, style, nav, footer, header").remove();
    const text = $("body").text().replace(/\s+/g, " ").trim().slice(0, 3000);

    const items = [];
    const seen = new Set();

    function addItem(rawLabel, rawValue) {
      const label = rawLabel.replace(/[\s　]+/g, "").trim();
      const value = rawValue.replace(/\s+/g, " ").trim();
      if (!label || !value || value.length > 300 || seen.has(label)) return;
      seen.add(label);
      items.push({ label, value });
    }

    // dl > dt + dd パターン（最も一般的な会社概要テーブル構造）
    $("dl").each((_, dl) => {
      $(dl).find("dt").each((_, dt) => {
        const dd = $(dt).nextAll("dd").first();
        if (dd.length) addItem($(dt).text(), dd.text());
      });
    });

    // table > tr > th + td パターン
    if (items.length < 3) {
      $("table tr").each((_, tr) => {
        const th = $(tr).find("th").first();
        const td = $(tr).find("td").first();
        if (th.length && td.length) {
          addItem(th.text(), td.text());
        } else {
          const tds = $(tr).find("td");
          if (tds.length >= 2) addItem($(tds[0]).text(), $(tds[1]).text());
        }
      });
    }

    // h2/h3/h4 見出し＋後続テキスト パターン（dl/tableがない日本語サイトに多い）
    if (items.length < 3) {
      $("h2, h3, h4, h5").each((_, heading) => {
        const label = $(heading).text();
        // 直接の兄弟から次の見出しまでのテキストを取得
        let cur = $(heading).next();
        const parts = [];
        while (cur.length && !cur.is("h2,h3,h4,h5") && parts.length < 4) {
          const t = cur.text().replace(/\s+/g, " ").trim();
          if (t) parts.push(t);
          cur = cur.next();
        }
        // 直接兄弟になければ親の次の兄弟も試す
        if (parts.length === 0) {
          let cur2 = $(heading).parent().next();
          while (cur2.length && !cur2.find("h2,h3,h4,h5").length && parts.length < 4) {
            const t = cur2.text().replace(/\s+/g, " ").trim();
            if (t) parts.push(t);
            cur2 = cur2.next();
          }
        }
        if (parts.length > 0) addItem(label, parts.join(" ").slice(0, 200));
      });
    }

    return { text, items };
  } catch {
    return { text: "", items: [] };
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
    "/company", "/company.html", "/about", "/about.html",
    "/company/profile", "/company/profile.html", "/corporate", "/corporate.html",
    "/company/", "/about/", "/about-us", "/aboutus",
    "/company/overview", "/company/about", "/corporate/profile",
    "/info", "/info.html", "/profile", "/profile.html", "/outline", "/outline.html",
  ];
  let companyPageText = "";
  let rawCompanyDetails = [];
  for (const pagePath of companyCandidates) {
    const result = await fetchCompanyPage(baseUrl + pagePath);
    if (result.text.length > 100) {
      companyPageText = result.text;
      rawCompanyDetails = result.items;
      console.log(`[Scraper] 会社概要ページ取得: ${pagePath}, items=${result.items.length}, text=${result.text.length}文字`);
      break;
    }
  }
  if (!companyPageText) {
    console.log("[Scraper] 会社概要ページ: 候補URL全て不一致、トップページ本文を使用");
  }

  // テキストから正規表現で住所・代表者を直接抽出（構造化データが取れない場合のフォールバック）
  const fullText = companyPageText || bodyText;
  let extractedAddress = "";
  let extractedRep = "";

  // 住所: 〒XXX-XXXX 形式 または 都道府県から始まる住所
  const zipMatch = fullText.match(/〒\s*\d{3}[-－]\d{4}[\s　]*([^\s　]{4,40})/);
  if (zipMatch) {
    extractedAddress = `〒${zipMatch[0].replace(/^〒\s*/, "")}`;
  } else {
    const prefMatch = fullText.match(/(東京都|大阪府|京都府|北海道|[^\s　]{2,3}[都道府県])[^\s　、。]{5,40}/);
    if (prefMatch) extractedAddress = prefMatch[0];
  }

  // 代表者: 役職＋氏名パターン（"代表取締役会長 良原一行" など）
  const repMatch = fullText.match(/(代表取締役(?:会長|社長|CEO)?|取締役(?:会長|社長)?|代表社員|理事長|会長|社長)[　\s]+([^\s　、。,\n]{2,10})/);
  if (repMatch) {
    extractedRep = `${repMatch[1]} ${repMatch[2].replace(/ほか.*$|など.*$/, "").trim()}`;
  }

  console.log(`[Scraper] 抽出: address="${extractedAddress}", rep="${extractedRep}", details=${rawCompanyDetails.length}件`);

  return { title, description, bodyText, companyPageText, rawCompanyDetails, extractedAddress, extractedRep, brandColors, dominantColor, url };
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
    await new Promise((r) => setTimeout(r, 1500)); // JS描画完了を待つ
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
