const PptxGenJS = require("pptxgenjs");

// ========== カラー定数 ==========
let _author = "MANDA"; // generateIM() で上書きされる
const DEFAULT_COLOR = "5E9E7E";
const WHITE = "FFFFFF";
const BLACK = "1A1A1A";
const LGRAY = "F5F5F5";
const DGRAY = "555555";
const DARK = "333333";

// メインカラーから薄い背景色・枠線色を生成
function hexToRgb(hex) {
  const r = parseInt(hex.slice(0, 2), 16);
  const g = parseInt(hex.slice(2, 4), 16);
  const b = parseInt(hex.slice(4, 6), 16);
  return { r, g, b };
}
function rgbToHex(r, g, b) {
  return [r, g, b].map((v) => Math.min(255, Math.max(0, Math.round(v))).toString(16).padStart(2, "0")).join("").toUpperCase();
}
function lightenColor(hex, factor) {
  const { r, g, b } = hexToRgb(hex);
  return rgbToHex(r + (255 - r) * factor, g + (255 - g) * factor, b + (255 - b) * factor);
}

const W = 13.33;
const H = 7.5;

// ========== ヘルパー関数 ==========
function getDate() {
  const d = new Date();
  return `${d.getFullYear()}/${String(d.getMonth() + 1).padStart(2, "0")}/${String(d.getDate()).padStart(2, "0")}`;
}

function footer(pptx, slide, pageNum, GREEN) {
  const DATE = getDate();
  slide.addShape(pptx.ShapeType.rect, {
    x: 0, y: H - 0.38, w: W, h: 0.38, fill: { color: GREEN }, line: { color: GREEN },
  });
  slide.addText(DATE, {
    x: 0.2, y: H - 0.35, w: 2.5, h: 0.3,
    fontSize: 11, color: WHITE, fontFace: "メイリオ",
  });
  slide.addText(`${_author}  |  Confidential`, {
    x: 0, y: H - 0.35, w: W, h: 0.3,
    fontSize: 11, color: WHITE, align: "center", fontFace: "メイリオ",
  });
  slide.addText(String(pageNum), {
    x: W - 0.8, y: H - 0.35, w: 0.6, h: 0.3,
    fontSize: 11, color: WHITE, align: "right", fontFace: "メイリオ",
  });
}

function sectionTitle(slide, title, GREEN, MGRAY) {
  slide.addShape("rect", {
    x: 0.4, y: 0.22, w: 0.07, h: 0.52, fill: { color: GREEN }, line: { color: GREEN },
  });
  slide.addText(title, {
    x: 0.55, y: 0.2, w: 11, h: 0.56,
    fontSize: 21, bold: true, color: BLACK, fontFace: "メイリオ",
  });
  slide.addShape("rect", {
    x: 0.4, y: 0.82, w: W - 0.8, h: 0.02, fill: { color: MGRAY }, line: { color: MGRAY },
  });
}

function divSlide(pptx, title, page, GREEN) {
  const slide = pptx.addSlide();
  slide.background = { color: WHITE };
  slide.addShape("rect", { x: 0, y: 0, w: W, h: 0.25, fill: { color: GREEN }, line: { color: GREEN } });
  slide.addShape("rect", { x: 0, y: H - 0.25, w: W, h: 0.25, fill: { color: GREEN }, line: { color: GREEN } });
  slide.addShape("rect", { x: 0, y: 0.25, w: 0.25, h: H - 0.5, fill: { color: GREEN }, line: { color: GREEN } });
  slide.addShape("rect", { x: W - 0.25, y: 0.25, w: 0.25, h: H - 0.5, fill: { color: GREEN }, line: { color: GREEN } });
  slide.addText(title, {
    x: 0.8, y: H / 2 - 0.5, w: W - 1.6, h: 1,
    fontSize: 39, bold: true, color: BLACK, align: "center", fontFace: "メイリオ",
  });
  footer(pptx, slide, page, GREEN);
  return slide;
}

// ========== メイン生成関数 ==========
/**
 * IMデータからPPTXバッファを生成する
 * @param {Object} data - claude-processor.jsで生成された構造化データ
 * @returns {Promise<Buffer>}
 */
async function generateIM(data) {
  _author = data.author || "MANDA";
  const pptx = new PptxGenJS();
  pptx.layout = "LAYOUT_WIDE";
  pptx.author = _author;

  // テーマカラー設定（# プレフィックス除去 + 6桁HEX検証 → 不正な場合はデフォルト）
  const rawColor = (data.themeColor || data._themeColor || DEFAULT_COLOR).replace(/^#/, "").toUpperCase();
  const GREEN = /^[0-9A-F]{6}$/.test(rawColor) ? rawColor : DEFAULT_COLOR;
  const MGRAY = lightenColor(GREEN, 0.5);   // 枠線：メインの50%薄め
  const GREEN_BG = lightenColor(GREEN, 0.8); // 背景：メインの80%薄め

  // 円 → 万円変換（KPIカード表示用）
  function toMan(val) {
    if (!val || val === "-") return val || "-";
    const str = String(val).replace(/,/g, "");
    if (/[%円]/.test(str)) return val; // %や円が含まれる場合はそのまま
    const num = parseFloat(str);
    if (isNaN(num) || Math.abs(num) < 10000) return val;
    const man = Math.round(num / 10000);
    return man.toString().replace(/\B(?=(\d{3})+(?!\d))/g, ",") + "万円";
  }

  const c = data.company || {};
  const f = data.financials || {};
  const pl = f.pl || {};
  const bs = f.bs || {};
  const sga = f.sga || {};
  const kpi = f.kpi || {};
  const transfer = data.transfer || {};
  const stores = data.stores || [];
  const menu = data.menu || [];
  const strengths = data.strengths || [];
  const plComments = data.plComments || [];

  const screenshotData = data.screenshotData || null;

  // 複数期間データ（損益計算書の多期比較用）
  const periods = (data.financialPeriods && data.financialPeriods.length > 0)
    ? data.financialPeriods
    : [{ period: f.period, pl, bs, sga, kpi }];

  // 最新期（期番号が最大の期）を特定 — 1期分のみ使うスライドはすべてこれを使う
  const latestPeriod = periods.reduce((best, p) => {
    const num = parseInt(((p.period || "").match(/第(\d+)期/) || [0, 0])[1]) || 0;
    const bestNum = parseInt(((best.period || "").match(/第(\d+)期/) || [0, 0])[1]) || 0;
    return num > bestNum ? p : best;
  }, periods[periods.length - 1]);
  const bsLargest = latestPeriod.bs || bs;
  const sgaLargest = latestPeriod.sga || sga;
  const plLargest = latestPeriod.pl || pl;

  // 主要財務指標スライド追加によるページオフセット
  const fsOffset = 1;

  // ========== Slide 1: 表紙 ==========
  {
    const s = pptx.addSlide();
    s.background = { color: WHITE };
    s.addShape("rect", { x: 0, y: 0, w: W, h: 0.28, fill: { color: GREEN }, line: { color: GREEN } });
    s.addShape("rect", { x: 0, y: H - 0.28, w: W, h: 0.28, fill: { color: GREEN }, line: { color: GREEN } });
    s.addShape("rect", { x: 0, y: 0.28, w: 0.22, h: H - 0.56, fill: { color: GREEN }, line: { color: GREEN } });
    s.addShape("rect", { x: W - 0.22, y: 0.28, w: 0.22, h: H - 0.56, fill: { color: GREEN }, line: { color: GREEN } });
    s.addText(c.name || "会社名", {
      x: 1, y: 2.5, w: W - 2, h: 1.2,
      fontSize: 47, bold: true, color: BLACK, align: "center", fontFace: "メイリオ",
    });
    s.addText("企業概要書", {
      x: 1, y: 3.85, w: W - 2, h: 1,
      fontSize: 41, bold: true, color: GREEN, align: "center", fontFace: "メイリオ",
    });
    s.addText(`作成：${_author}`, {
      x: 1, y: 5.0, w: W - 2, h: 0.45,
      fontSize: 15, color: DGRAY, align: "center", fontFace: "メイリオ",
    });
    s.addText(getDate(), {
      x: 0.35, y: H - 0.26, w: 3, h: 0.22,
      fontSize: 11, color: WHITE, fontFace: "メイリオ",
    });
    s.addText(_author, {
      x: W - 3.2, y: H - 0.26, w: 3.0, h: 0.22,
      fontSize: 11, color: WHITE, align: "right", fontFace: "メイリオ",
    });
  }

  // ========== Slide 2: 免責事項 ==========
  {
    const s = pptx.addSlide();
    s.background = { color: WHITE };
    sectionTitle(s, "免責事項", GREEN, MGRAY);
    footer(pptx, s, 2, GREEN);
    const items = [
      "本書は、貴社による買収検討資料としてのみご利用されることを目的として作成されたものであります。本書の複製、再製または第三者への提供については、必ず事前に書面による同意をお取りくださいますようお願い申し上げます。",
      "本書に含まれる情報の一切について、正確性、完全性及び信頼性についていかなる明示的又は黙示的な表明又は保証を行うものではありません。",
      "本書は、対象会社の将来の業績に関する記述、評価及び予想を含むことがあります。かかる記述、評価及び予想は主観的な判断も含まれ、今後変更される可能性があります。",
      "本書は、本件取引の検討、評価を行う上で必要となる情報が全て網羅されているわけではありません。",
      "本書は、別途締結された秘密保持契約書に関する取り決めに基づいて提供されるものであり、その存在及び内容は全て機密情報として取り扱いにご留意いただきますようお願い申し上げます。",
      "本書は、別段の記載がない限り、本書作成日現在の情報をもって作成されています。",
    ];
    let y = 1.05;
    items.forEach((txt) => {
      s.addText("• " + txt, {
        x: 0.5, y, w: W - 1, h: 0.5,
        fontSize: 12.5, color: BLACK, fontFace: "メイリオ", valign: "top",
      });
      y += 0.6;
    });
  }

  // ========== Slide 3: 目次 ==========
  {
    const s = pptx.addSlide();
    s.background = { color: WHITE };
    sectionTitle(s, "目次", GREEN, MGRAY);
    footer(pptx, s, 3, GREEN);

    const hasStores = stores.length > 0;
    const hasMenu = menu.length > 0;

    const hasOwners = (data.shareholders || []).length > 0;
    const hasOfficers = (data.officers || []).length > 0;
    const hasOwnerSlide = hasOwners || hasOfficers;

    const po = (hasOwnerSlide ? 1 : 0) + (hasStores ? 1 : 0) + (hasMenu ? 1 : 0);
    const tocItems = [
      { text: "I.   エグゼクティブサマリー", indent: false, page: "" },
      { text: "1.  会社概要・主要指標", indent: true, page: "P. 5" },
      { text: "2.  主要財務指標", indent: true, page: "P. 6" },
      { text: "II.  会社概要", indent: false, page: "" },
      { text: "1.  対象会社概要", indent: true, page: `P. ${7 + fsOffset}` },
      ...(hasOwnerSlide ? [{ text: "2.  株主構成", indent: true, page: `P. ${8 + fsOffset}` }] : []),
      ...(hasStores ? [{ text: hasOwnerSlide ? "3.  店舗展開" : "2.  店舗展開", indent: true, page: `P. ${8 + fsOffset + (hasOwnerSlide ? 1 : 0)}` }] : []),
      { text: "III. 事業概要", indent: false, page: "" },
      ...(hasMenu ? [{ text: "1.  事業メニュー・サービス", indent: true, page: `P. ${10 + fsOffset + (hasOwnerSlide ? 1 : 0)}` }] : []),
      { text: hasMenu ? "2.  実績・データ" : "1.  実績・データ", indent: true, page: `P. ${11 + fsOffset + (hasOwnerSlide ? 1 : 0)}` },
      { text: "IV.  財務実績", indent: false, page: "" },
      { text: "1.  損益計算書", indent: true, page: `P. ${13 + fsOffset + (hasOwnerSlide ? 1 : 0)}` },
      { text: "2.  貸借対照表", indent: true, page: `P. ${14 + fsOffset + (hasOwnerSlide ? 1 : 0)}` },
      { text: "3.  販売費及び一般管理費", indent: true, page: `P. ${15 + fsOffset + (hasOwnerSlide ? 1 : 0)}` },
      { text: "V.   譲渡情報", indent: false, page: "" },
      { text: "1.  希望譲渡条件", indent: true, page: `P. ${17 + fsOffset + po}` },
      { text: "2.  ノンネーム情報", indent: true, page: `P. ${18 + fsOffset + po}` },
    ];

    s.addShape("rect", {
      x: W / 2, y: 0.95, w: 0.02, h: H - 1.45, fill: { color: MGRAY }, line: { color: MGRAY },
    });
    s.addText("Section", { x: 0.5, y: 0.9, w: 5, h: 0.25, fontSize: 13, bold: true, color: DGRAY, fontFace: "メイリオ" });
    s.addText("Page", { x: 5.8, y: 0.9, w: 0.8, h: 0.25, fontSize: 13, bold: true, color: DGRAY, fontFace: "メイリオ" });
    s.addText("Section", { x: W / 2 + 0.3, y: 0.9, w: 5.5, h: 0.25, fontSize: 13, bold: true, color: DGRAY, fontFace: "メイリオ" });
    s.addText("Page", { x: W - 1.3, y: 0.9, w: 0.8, h: 0.25, fontSize: 13, bold: true, color: DGRAY, fontFace: "メイリオ" });

    const mid = Math.ceil(tocItems.length / 2);
    const leftItems = tocItems.slice(0, mid);
    const rightItems = tocItems.slice(mid);

    let yL = 1.2;
    leftItems.forEach((item) => {
      const xBase = item.indent ? 0.9 : 0.5;
      s.addText(item.text, {
        x: xBase, y: yL, w: 5, h: 0.32,
        fontSize: item.indent ? 10.5 : 11.5, bold: !item.indent,
        color: item.indent ? DGRAY : BLACK, fontFace: "メイリオ",
      });
      if (item.page) {
        s.addText(item.page, {
          x: 5.5, y: yL, w: 1, h: 0.32,
          fontSize: 13.5, color: DGRAY, align: "right", fontFace: "メイリオ",
        });
      }
      yL += item.indent ? 0.36 : 0.46;
    });

    let yR = 1.2;
    rightItems.forEach((item) => {
      const xBase = item.indent ? W / 2 + 0.7 : W / 2 + 0.35;
      s.addText(item.text, {
        x: xBase, y: yR, w: 5, h: 0.32,
        fontSize: item.indent ? 10.5 : 11.5, bold: !item.indent,
        color: item.indent ? DGRAY : BLACK, fontFace: "メイリオ",
      });
      if (item.page) {
        s.addText(item.page, {
          x: W - 1.3, y: yR, w: 1, h: 0.32,
          fontSize: 13.5, color: DGRAY, align: "right", fontFace: "メイリオ",
        });
      }
      yR += item.indent ? 0.36 : 0.46;
    });
  }

  // ========== Slide 4: Section I ==========
  divSlide(pptx, "I.　エグゼクティブサマリー", 4, GREEN);

  // ========== Slide 5: 会社概要サマリー ==========
  {
    const s = pptx.addSlide();
    s.background = { color: WHITE };
    sectionTitle(s, "会社概要", GREEN, MGRAY);
    footer(pptx, s, 5, GREEN);

    s.addText(data.summary || `${c.name}の企業概要`, {
      x: 0.5, y: 0.9, w: W - 1, h: 0.85, fontSize: 12.5, color: BLACK, fontFace: "メイリオ",
    });

    // 左：会社概要テーブル（社名・所在地・事業内容・代表者の4項目）
    s.addText(`■ ${c.name}：概要`, {
      x: 0.5, y: 1.8, w: 5.8, h: 0.3, fontSize: 14, bold: true, color: BLACK, fontFace: "メイリオ",
    });
    const companyRows = [
      ["社名",     c.name     || ""],
      ["所在地",   c.address  || ""],
      ["事業内容", c.business || ""],
      ["代表者",   c.rep      || ""],
    ];
    const companyRowH = [0.5, 0.65, 2.4, 0.5];
    let companyY = 2.15;
    companyRows.forEach((row, i) => {
      const rh = companyRowH[i];
      s.addShape("rect", { x: 0.5, y: companyY, w: 1.5, h: rh, fill: { color: LGRAY }, line: { color: MGRAY } });
      s.addShape("rect", { x: 2.0, y: companyY, w: 4.3, h: rh, fill: { color: WHITE }, line: { color: MGRAY } });
      s.addText(row[0], { x: 0.55, y: companyY + 0.05, w: 1.4, h: rh - 0.1, fontSize: 12.5, bold: true, color: DGRAY, fontFace: "メイリオ", valign: "middle" });
      s.addText(row[1], { x: 2.05, y: companyY + 0.05, w: 4.2, h: rh - 0.1, fontSize: 12.5, color: BLACK, fontFace: "メイリオ", valign: "top" });
      companyY += rh + 0.02;
    });

    // 右：ホームページスクリーンショット
    s.addText("■ ホームページ", {
      x: 6.9, y: 1.8, w: 6.0, h: 0.3, fontSize: 14, bold: true, color: BLACK, fontFace: "メイリオ",
    });
    if (screenshotData) {
      s.addImage({ data: screenshotData, x: 6.9, y: 2.15, w: 6.0, h: 4.1 });
    } else {
      s.addShape("rect", { x: 6.9, y: 2.15, w: 6.0, h: 4.1, fill: { color: LGRAY }, line: { color: MGRAY } });
      s.addText("HP スクリーンショット\n（取得中または取得失敗）", {
        x: 6.9, y: 3.8, w: 6.0, h: 0.8, fontSize: 12, color: DGRAY, align: "center", fontFace: "メイリオ",
      });
    }
  }

  // ========== Slide 6: 主要財務指標 ==========
  {
    const s = pptx.addSlide();
    s.background = { color: WHITE };
    sectionTitle(s, `主要財務指標（${plLargest.period || latestPeriod.period || "直近期"}）`, GREEN, MGRAY);
    footer(pptx, s, 6, GREEN);

    // 上段：4つの主要指標カード
    const topKpis = [
      { label: "売上高",    value: toMan(plLargest.revenue         || "-") },
      { label: "営業利益",  value: toMan(plLargest.operatingProfit || "-"), sub: plLargest.operatingMargin ? `利益率 ${plLargest.operatingMargin}` : "" },
      { label: "経常利益",  value: toMan(plLargest.ordinaryProfit  || "-"), sub: plLargest.ordinaryMargin  ? `利益率 ${plLargest.ordinaryMargin}`  : "" },
      { label: "当期純利益",value: toMan(plLargest.netProfit        || "-") },
    ];
    topKpis.forEach((item, i) => {
      const x = 0.5 + i * 3.1;
      s.addShape("rect", { x, y: 1.05, w: 2.8, h: 1.8, fill: { color: GREEN_BG }, line: { color: GREEN } });
      s.addText(item.label, { x, y: 1.18, w: 2.8, h: 0.3, fontSize: 13, bold: true, color: GREEN, align: "center", fontFace: "メイリオ" });
      s.addText(item.value, { x, y: 1.52, w: 2.8, h: 0.85, fontSize: 17, bold: true, color: BLACK, align: "center", fontFace: "メイリオ" });
      s.addText(item.sub || "", { x, y: 2.45, w: 2.8, h: 0.3, fontSize: 11, color: DGRAY, align: "center", fontFace: "メイリオ" });
    });

    // 下段：PLテーブル（全項目）
    s.addText("■ 損益サマリー", {
      x: 0.5, y: 3.1, w: W - 1, h: 0.3, fontSize: 14, bold: true, color: BLACK, fontFace: "メイリオ",
    });
    const sumRows = [
      { label: "売上高",              val: plLargest.revenue         || "-", bold: true  },
      { label: "　売上原価",           val: plLargest.cogs            || "-", bold: false },
      { label: "売上総利益",           val: plLargest.grossProfit     || "-", sub: plLargest.grossMargin    ? `粗利率 ${plLargest.grossMargin}`    : "", bold: true  },
      { label: "　販売費及び一般管理費",val: plLargest.sga             || "-", bold: false },
      { label: "営業利益",             val: plLargest.operatingProfit || "-", sub: plLargest.operatingMargin ? `利益率 ${plLargest.operatingMargin}` : "", bold: true  },
      { label: "経常利益",             val: plLargest.ordinaryProfit  || "-", sub: plLargest.ordinaryMargin  ? `利益率 ${plLargest.ordinaryMargin}`  : "", bold: true  },
      { label: "当期純利益",           val: plLargest.netProfit       || "-", bold: true  },
      { label: "EBITDA（概算）",       val: plLargest.ebitda          || "-", bold: false },
    ];
    const lblW2 = 4.0, valW2 = 2.2, subW2 = 2.5;
    const sumHY = 3.45;
    s.addShape("rect", { x: 0.5, y: sumHY, w: lblW2, h: 0.3, fill: { color: DARK }, line: { color: WHITE } });
    s.addShape("rect", { x: 0.5 + lblW2, y: sumHY, w: valW2, h: 0.3, fill: { color: DARK }, line: { color: WHITE } });
    s.addShape("rect", { x: 0.5 + lblW2 + valW2, y: sumHY, w: subW2, h: 0.3, fill: { color: DARK }, line: { color: WHITE } });
    s.addText("科目", { x: 0.55, y: sumHY + 0.04, w: lblW2 - 0.1, h: 0.22, fontSize: 12, bold: true, color: WHITE, fontFace: "メイリオ" });
    s.addText("金額", { x: 0.5 + lblW2 + 0.05, y: sumHY + 0.04, w: valW2 - 0.1, h: 0.22, fontSize: 12, bold: true, color: WHITE, align: "right", fontFace: "メイリオ" });
    s.addText("備考", { x: 0.5 + lblW2 + valW2 + 0.05, y: sumHY + 0.04, w: subW2 - 0.1, h: 0.22, fontSize: 12, bold: true, color: WHITE, fontFace: "メイリオ" });
    sumRows.forEach((row, i) => {
      const yy = sumHY + 0.34 + i * 0.34;
      const fill = row.bold ? (i === 2 ? LGRAY : (i === 6 ? "C8E6D4" : (i >= 4 ? GREEN_BG : WHITE))) : WHITE;
      s.addShape("rect", { x: 0.5, y: yy, w: lblW2, h: 0.32, fill: { color: fill }, line: { color: MGRAY } });
      s.addShape("rect", { x: 0.5 + lblW2, y: yy, w: valW2, h: 0.32, fill: { color: fill }, line: { color: MGRAY } });
      s.addShape("rect", { x: 0.5 + lblW2 + valW2, y: yy, w: subW2, h: 0.32, fill: { color: fill }, line: { color: MGRAY } });
      s.addText(row.label, { x: 0.55, y: yy + 0.04, w: lblW2 - 0.1, h: 0.24, fontSize: 12, bold: row.bold, color: BLACK, fontFace: "メイリオ" });
      s.addText(row.val,   { x: 0.5 + lblW2 + 0.05, y: yy + 0.04, w: valW2 - 0.1, h: 0.24, fontSize: 12, bold: row.bold, color: BLACK, align: "right", fontFace: "メイリオ" });
      s.addText(row.sub || "", { x: 0.5 + lblW2 + valW2 + 0.05, y: yy + 0.04, w: subW2 - 0.1, h: 0.24, fontSize: 11, color: DGRAY, fontFace: "メイリオ" });
    });

    // 右：BSサマリー
    s.addText("■ 貸借対照表サマリー", {
      x: 8.8, y: 3.1, w: 4.3, h: 0.3, fontSize: 14, bold: true, color: BLACK, fontFace: "メイリオ",
    });
    const bsSum = [
      { label: "資産合計",   val: bsLargest.totalAssets       || "-", bold: true },
      { label: "負債合計",   val: bsLargest.totalLiabilities  || "-", bold: false },
      { label: "　短期借入金", val: bsLargest.shortTermLoan   || "-", bold: false },
      { label: "　長期借入金", val: bsLargest.longTermLoan    || "-", bold: false },
      { label: "純資産",     val: bsLargest.equity             || "-", bold: true },
    ];
    const bsLX = 8.8, bsLW = 2.2, bsVW = 2.0;
    s.addShape("rect", { x: bsLX, y: sumHY, w: bsLW, h: 0.3, fill: { color: DARK }, line: { color: WHITE } });
    s.addShape("rect", { x: bsLX + bsLW, y: sumHY, w: bsVW, h: 0.3, fill: { color: DARK }, line: { color: WHITE } });
    s.addText("科目", { x: bsLX + 0.05, y: sumHY + 0.04, w: bsLW - 0.1, h: 0.22, fontSize: 12, bold: true, color: WHITE, fontFace: "メイリオ" });
    s.addText("金額", { x: bsLX + bsLW + 0.05, y: sumHY + 0.04, w: bsVW - 0.1, h: 0.22, fontSize: 12, bold: true, color: WHITE, align: "right", fontFace: "メイリオ" });
    bsSum.forEach((row, i) => {
      const yy = sumHY + 0.34 + i * 0.34;
      const fill = row.bold ? GREEN_BG : WHITE;
      s.addShape("rect", { x: bsLX, y: yy, w: bsLW, h: 0.32, fill: { color: fill }, line: { color: MGRAY } });
      s.addShape("rect", { x: bsLX + bsLW, y: yy, w: bsVW, h: 0.32, fill: { color: fill }, line: { color: MGRAY } });
      s.addText(row.label, { x: bsLX + 0.05, y: yy + 0.04, w: bsLW - 0.1, h: 0.24, fontSize: 12, bold: row.bold, color: BLACK, fontFace: "メイリオ" });
      s.addText(row.val,   { x: bsLX + bsLW + 0.05, y: yy + 0.04, w: bsVW - 0.1, h: 0.24, fontSize: 12, bold: row.bold, color: BLACK, align: "right", fontFace: "メイリオ" });
    });
  }

  // ========== Slide 7: Section II ==========
  divSlide(pptx, "II.　会社概要", 6 + fsOffset, GREEN);

  // ========== Slide 8: 対象会社概要 ==========
  {
    const s = pptx.addSlide();
    s.background = { color: WHITE };
    sectionTitle(s, "対象会社概要", GREEN, MGRAY);
    footer(pptx, s, 7 + fsOffset, GREEN);

    s.addText(`■ ${c.name}：概要`, {
      x: 0.5, y: 0.95, w: 5.5, h: 0.3, fontSize: 14, bold: true, color: BLACK, fontFace: "メイリオ",
    });
    // HPの会社概要テーブルがあればそれを使用、なければ標準フィールドにフォールバック
    const companyDetails = (data.companyDetails && data.companyDetails.length > 0)
      ? data.companyDetails.filter(d => d.label && d.value)
      : [
          { label: "社名",     value: c.name        || "" },
          { label: "設立",     value: c.established || "" },
          { label: "決算期",   value: c.fiscal      || "" },
          { label: "資本金",   value: c.capital     || "" },
          { label: "本社住所", value: c.address     || "" },
          { label: "代表者",   value: c.rep         || "" },
          { label: "電話番号", value: c.tel         || "" },
          { label: "URL",      value: c.url         || "" },
          { label: "事業内容", value: c.business    || "" },
          { label: "従業員数", value: c.employees   || "" },
        ].filter(d => d.value);

    function detailRowHeight(label) {
      if (label === "事業内容" || label === "事業目的" || label === "沿革") return 1.4;
      if (/住所|所在地|本社|設立|主要取引先|許認可|加盟団体|取引銀行/.test(label)) return 0.6;
      return 0.44;
    }
    let detailY = 1.3;
    companyDetails.forEach((row) => {
      const rh = detailRowHeight(row.label);
      if (detailY + rh > 6.8) return; // スライド下端を超えない
      s.addShape("rect", { x: 0.5, y: detailY, w: 1.6, h: rh, fill: { color: LGRAY }, line: { color: MGRAY } });
      s.addShape("rect", { x: 2.1, y: detailY, w: 4.2, h: rh, fill: { color: WHITE }, line: { color: MGRAY } });
      s.addText(row.label, { x: 0.55, y: detailY + 0.05, w: 1.5, h: rh - 0.1, fontSize: 12, bold: true, color: DGRAY, fontFace: "メイリオ", valign: "middle" });
      s.addText(row.value, { x: 2.15, y: detailY + 0.05, w: 4.1, h: rh - 0.1, fontSize: 12, color: BLACK, fontFace: "メイリオ", valign: "top", shrinkText: true });
      detailY += rh + 0.02;
    });

    // 右：特徴・強み
    s.addText(`■ ${c.name}：特徴・強み`, {
      x: 7.0, y: 0.95, w: 5.8, h: 0.3, fontSize: 14, bold: true, color: BLACK, fontFace: "メイリオ",
    });
    const displayStrengths = strengths.length > 0
      ? strengths
      : [{ label: "事業特性", desc: c.business || "" }];

    let yR = 1.3;
    displayStrengths.slice(0, 6).forEach((item) => {
      const itemH = 0.75;
      s.addShape("rect", { x: 7.0, y: yR, w: 1.2, h: itemH, fill: { color: GREEN_BG }, line: { color: MGRAY } });
      s.addShape("rect", { x: 8.2, y: yR, w: 4.8, h: itemH, fill: { color: WHITE }, line: { color: MGRAY } });
      s.addText(item.label || "", {
        x: 7.05, y: yR + 0.06, w: 1.1, h: itemH - 0.12, fontSize: 11, bold: true, color: GREEN, align: "center", fontFace: "メイリオ", valign: "middle",
      });
      s.addText(item.desc || "", {
        x: 8.25, y: yR + 0.05, w: 4.7, h: itemH - 0.1, fontSize: 12, color: BLACK, fontFace: "メイリオ", valign: "top", shrinkText: true,
      });
      yR += itemH + 0.02;
    });
  }

  // ========== Slide 9: 株主・役員構成（入力がある場合） ==========
  const shareholders = data.shareholders || [];
  const empBreak = data.employeeBreakdown || {};
  const hasOwnerSlide = shareholders.length > 0;
  let pageOffset = 0;

  if (hasOwnerSlide) {
    pageOffset = 1;
    const s = pptx.addSlide();
    s.background = { color: WHITE };
    sectionTitle(s, "株主構成", GREEN, MGRAY);
    footer(pptx, s, 8 + fsOffset, GREEN);

    // ヘッダー行（幅広レイアウト）
    const shNameW = 7.0;
    const shRatioW = 2.5;
    const shBarW  = 2.5;
    const shX = 0.4;
    s.addShape("rect", { x: shX,                   y: 1.3, w: shNameW,  h: 0.34, fill: { color: GREEN }, line: { color: WHITE } });
    s.addShape("rect", { x: shX + shNameW,          y: 1.3, w: shRatioW, h: 0.34, fill: { color: GREEN }, line: { color: WHITE } });
    s.addShape("rect", { x: shX + shNameW + shRatioW, y: 1.3, w: shBarW,  h: 0.34, fill: { color: GREEN }, line: { color: WHITE } });
    s.addText("株主名",   { x: shX + 0.1,                    y: 1.34, w: shNameW - 0.2,  h: 0.26, fontSize: 13, bold: true, color: WHITE, fontFace: "メイリオ" });
    s.addText("持株比率", { x: shX + shNameW + 0.1,           y: 1.34, w: shRatioW - 0.2, h: 0.26, fontSize: 13, bold: true, color: WHITE, align: "center", fontFace: "メイリオ" });
    s.addText("構成比",   { x: shX + shNameW + shRatioW + 0.1, y: 1.34, w: shBarW - 0.2,  h: 0.26, fontSize: 13, bold: true, color: WHITE, align: "center", fontFace: "メイリオ" });

    shareholders.slice(0, 10).forEach((sh, i) => {
      const yy = 1.68 + i * 0.44;
      const fill = i % 2 === 0 ? WHITE : LGRAY;
      s.addShape("rect", { x: shX,                    y: yy, w: shNameW,  h: 0.42, fill: { color: fill }, line: { color: MGRAY } });
      s.addShape("rect", { x: shX + shNameW,           y: yy, w: shRatioW, h: 0.42, fill: { color: fill }, line: { color: MGRAY } });
      s.addShape("rect", { x: shX + shNameW + shRatioW, y: yy, w: shBarW,  h: 0.42, fill: { color: fill }, line: { color: MGRAY } });
      s.addText(sh.name || "", { x: shX + 0.1, y: yy + 0.07, w: shNameW - 0.2, h: 0.28, fontSize: 13, color: BLACK, fontFace: "メイリオ" });
      s.addText(sh.ratio || "", { x: shX + shNameW + 0.1, y: yy + 0.07, w: shRatioW - 0.2, h: 0.28, fontSize: 13, bold: true, color: GREEN, align: "center", fontFace: "メイリオ" });
      // 構成比バー
      const ratio = parseFloat((sh.ratio || "0").replace(/[^0-9.]/g, "")) || 0;
      const barFill = Math.min(shBarW - 0.3, (shBarW - 0.3) * ratio / 100);
      if (barFill > 0) {
        s.addShape("rect", { x: shX + shNameW + shRatioW + 0.15, y: yy + 0.13, w: barFill, h: 0.16, fill: { color: GREEN }, line: { color: GREEN } });
      }
    });

    // ---- 従業員内訳（下部） ----
    if (empBreak.full || empBreak.part) {
      s.addShape("rect", { x: 0.4, y: 6.3, w: W - 0.8, h: 0.02, fill: { color: MGRAY }, line: { color: MGRAY } });
      s.addText("■ 従業員内訳", { x: 0.4, y: 6.38, w: 3, h: 0.26, fontSize: 13, bold: true, color: BLACK, fontFace: "メイリオ" });
      const empParts = [];
      if (empBreak.full) empParts.push(`正社員：${empBreak.full}`);
      if (empBreak.part) empParts.push(`パート・アルバイト：${empBreak.part}`);
      s.addText(empParts.join("　　"), { x: 3.5, y: 6.38, w: 9, h: 0.26, fontSize: 13, color: BLACK, fontFace: "メイリオ" });
    }
  }

  // ========== 店舗展開（店舗がある場合） ==========
  if (stores.length > 0) {
    const s = pptx.addSlide();
    s.background = { color: WHITE };
    sectionTitle(s, "店舗展開", GREEN, MGRAY);
    footer(pptx, s, 8 + fsOffset + pageOffset, GREEN);

    s.addText(`${stores.length}店舗を展開。`,
      { x: 0.5, y: 0.9, w: W - 1, h: 0.36, fontSize: 13.5, color: BLACK, fontFace: "メイリオ" });

    const headers = ["No.", "店舗名", "都道府県", "最寄り駅・アクセス", "電話番号"];
    const colW = [0.5, 2.0, 1.1, 4.5, 2.5];
    const colX = [0.4];
    colW.slice(0, -1).forEach((w, i) => colX.push(colX[i] + colW[i]));

    const headerY = 1.35;
    headers.forEach((h, i) => {
      s.addShape("rect", { x: colX[i], y: headerY, w: colW[i], h: 0.36, fill: { color: GREEN }, line: { color: WHITE } });
      s.addText(h, {
        x: colX[i] + 0.05, y: headerY + 0.04, w: colW[i] - 0.1, h: 0.28,
        fontSize: 13, bold: true, color: WHITE, align: "center", fontFace: "メイリオ",
      });
    });

    stores.slice(0, 8).forEach((store, row) => {
      const yy = 1.75 + row * 0.52;
      const fillC = row % 2 === 0 ? WHITE : LGRAY;
      const rowData = [String(row + 1), store.name || "", store.pref || "", store.access || "", store.tel || ""];
      rowData.forEach((val, col) => {
        s.addShape("rect", { x: colX[col], y: yy, w: colW[col], h: 0.5, fill: { color: fillC }, line: { color: MGRAY } });
        s.addText(val, {
          x: colX[col] + 0.05, y: yy + 0.1, w: colW[col] - 0.1, h: 0.3,
          fontSize: col === 1 ? 11 : 10.5,
          bold: col === 1,
          color: col === 1 ? GREEN : BLACK,
          align: col === 0 ? "center" : "left",
          fontFace: "メイリオ",
        });
      });
    });
  }

  // ========== Section III ==========
  divSlide(pptx, "III.　事業概要", 9 + fsOffset + pageOffset, GREEN);

  // ========== サービスメニュー（あれば） ==========
  if (menu.length > 0) {
    const s = pptx.addSlide();
    s.background = { color: WHITE };
    sectionTitle(s, "事業メニュー・サービス", GREEN, MGRAY);
    footer(pptx, s, 10 + fsOffset + pageOffset, GREEN);

    s.addText("主要サービス・メニュー一覧",
      { x: 0.5, y: 0.9, w: W - 1, h: 0.4, fontSize: 13.5, color: BLACK, fontFace: "メイリオ" });

    const hcols = ["カテゴリ", "メニュー名", "概要"];
    const hwidths = [1.5, 2.8, 8.7];
    let hx = 0.4;
    const headerY2 = 1.38;
    hcols.forEach((h, i) => {
      s.addShape("rect", { x: hx, y: headerY2, w: hwidths[i], h: 0.34, fill: { color: DARK }, line: { color: WHITE } });
      s.addText(h, {
        x: hx + 0.05, y: headerY2 + 0.04, w: hwidths[i] - 0.1, h: 0.26,
        fontSize: 13, bold: true, color: WHITE, align: "center", fontFace: "メイリオ",
      });
      hx += hwidths[i];
    });

    menu.slice(0, 9).forEach((item, i) => {
      const yy = 1.76 + i * 0.56;
      const fillC = i % 2 === 0 ? WHITE : LGRAY;
      s.addShape("rect", { x: 0.4, y: yy, w: 1.5, h: 0.54, fill: { color: fillC }, line: { color: MGRAY } });
      s.addText(item.cat || "", {
        x: 0.45, y: yy + 0.12, w: 1.4, h: 0.3, fontSize: 12.5, bold: true, color: GREEN, align: "center", fontFace: "メイリオ",
      });
      s.addShape("rect", { x: 1.9, y: yy, w: 2.8, h: 0.54, fill: { color: fillC }, line: { color: MGRAY } });
      s.addText(item.name || "", {
        x: 1.95, y: yy + 0.12, w: 2.7, h: 0.3, fontSize: 13.5, bold: true, color: BLACK, fontFace: "メイリオ",
      });
      s.addShape("rect", { x: 4.7, y: yy, w: 8.4, h: 0.54, fill: { color: fillC }, line: { color: MGRAY } });
      s.addText(item.desc || "", {
        x: 4.75, y: yy + 0.12, w: 8.3, h: 0.3, fontSize: 13, color: DGRAY, fontFace: "メイリオ",
      });
    });
  }

  // ========== 実績データ ==========
  {
    const s = pptx.addSlide();
    s.background = { color: WHITE };
    sectionTitle(s, "実績・データ", GREEN, MGRAY);
    footer(pptx, s, 11 + fsOffset + pageOffset, GREEN);

    s.addText("主要KPIと財務ハイライト",
      { x: 0.5, y: 0.9, w: W - 1, h: 0.36, fontSize: 13.5, color: BLACK, fontFace: "メイリオ" });

    // 上段：KPIデータ（なければ財務指標から上段専用のものを使用）
    const hasKpi = kpi.customers || kpi.procedures || kpi.reviewCount || kpi.reviewScore;
    const topKpiItems = hasKpi
      ? [
          { label: kpi.customers  ? "累計顧客数"    : "売上高",    value: kpi.customers  || toMan(plLargest.revenue)        || "-", note: "" },
          { label: kpi.procedures ? "累計施術件数"  : "営業利益",  value: kpi.procedures || toMan(plLargest.operatingProfit) || "-", note: "" },
          { label: kpi.reviewCount? "口コミ件数"    : "経常利益",  value: kpi.reviewCount|| toMan(plLargest.ordinaryProfit)  || "-", note: kpi.reviewCount ? "口コミプラットフォーム" : "" },
          { label: kpi.reviewScore? "平均評価"      : "当期純利益",value: kpi.reviewScore|| toMan(plLargest.netProfit)       || "-", note: kpi.reviewScore ? "（5点満点）" : "" },
        ]
      : [
          { label: "売上高",    value: toMan(plLargest.revenue)         || "-", note: "" },
          { label: "営業利益",  value: toMan(plLargest.operatingProfit) || "-", note: plLargest.operatingMargin ? `利益率 ${plLargest.operatingMargin}` : "" },
          { label: "経常利益",  value: toMan(plLargest.ordinaryProfit)  || "-", note: plLargest.ordinaryMargin  ? `利益率 ${plLargest.ordinaryMargin}`  : "" },
          { label: "当期純利益",value: toMan(plLargest.netProfit)        || "-", note: "" },
        ];

    topKpiItems.forEach((kpiItem, i) => {
      const x = 0.5 + i * 3.1;
      s.addShape("rect", { x, y: 1.4, w: 2.8, h: 1.8, fill: { color: GREEN_BG }, line: { color: GREEN } });
      s.addText(kpiItem.label, {
        x, y: 1.55, w: 2.8, h: 0.3, fontSize: 13, bold: true, color: GREEN, align: "center", fontFace: "メイリオ",
      });
      s.addText(kpiItem.value, {
        x, y: 1.95, w: 2.8, h: 0.85, fontSize: 17, bold: true, color: BLACK, align: "center", fontFace: "メイリオ",
      });
      s.addText(kpiItem.note, {
        x, y: 2.85, w: 2.8, h: 0.3, fontSize: 12, color: DGRAY, align: "center", fontFace: "メイリオ",
      });
    });

    // 下段：上段と重複しない財務指標
    s.addText("■ 財務ハイライト", {
      x: 0.5, y: 3.45, w: W - 1, h: 0.32, fontSize: 14, bold: true, color: BLACK, fontFace: "メイリオ",
    });
    // 上段で使用済みラベルを収集し、下段は重複しない指標のみ表示
    const usedLabels = new Set(topKpiItems.map(k => k.label));
    const finCandidates = [
      { label: "粗利率",       value: plLargest.grossMargin            || "-" },
      { label: "EBITDA（概算）",value: toMan(plLargest.ebitda)         || "-" },
      { label: "販管費",       value: toMan(plLargest.sga)             || "-" },
      { label: "純資産",       value: toMan(bsLargest.equity)          || "-" },
      { label: "売上高",       value: toMan(plLargest.revenue)         || "-" },
      { label: "当期純利益",   value: toMan(plLargest.netProfit)       || "-" },
      { label: "経常利益",     value: toMan(plLargest.ordinaryProfit)  || "-" },
    ];
    const finKpis = finCandidates.filter(f2 => !usedLabels.has(f2.label)).slice(0, 4);
    finKpis.forEach((item, i) => {
      const x = 0.5 + i * 3.1;
      s.addShape("rect", { x, y: 3.85, w: 2.8, h: 1.6, fill: { color: LGRAY }, line: { color: MGRAY } });
      s.addText(item.label, {
        x, y: 3.98, w: 2.8, h: 0.3, fontSize: 13, bold: true, color: DGRAY, align: "center", fontFace: "メイリオ",
      });
      s.addText(item.value, {
        x, y: 4.3, w: 2.8, h: 0.75, fontSize: 15, bold: true, color: BLACK, align: "center", fontFace: "メイリオ",
      });
    });
  }

  // ========== Section IV ==========
  divSlide(pptx, "IV.　財務実績", 12 + fsOffset + pageOffset, GREEN);

  // ========== 損益計算書 ==========
  {
    const s = pptx.addSlide();
    s.background = { color: WHITE };
    // 表示する期（最大3期）を期番号で昇順ソート（古い期が左）
    const plPeriods = [...periods.slice(-3)].sort((a, b) => {
      const na = parseInt(((a.period || "").match(/第(\d+)期/) || [0, 0])[1]) || 0;
      const nb = parseInt(((b.period || "").match(/第(\d+)期/) || [0, 0])[1]) || 0;
      return na - nb;
    });
    const nPer = plPeriods.length;
    const periodTitle = nPer > 1 ? `損益計算書（過去${nPer}期比較）` : `損益計算書（${plPeriods[0].period || "直近期"}）`;
    sectionTitle(s, periodTitle, GREEN, MGRAY);
    footer(pptx, s, 13 + fsOffset + pageOffset, GREEN);

    s.addText(data.summary || "財務概要",
      { x: 0.5, y: 0.88, w: W - 1, h: 0.85, fontSize: 12, color: BLACK, fontFace: "メイリオ" });

    // 列幅の計算（期数に応じて変動）
    const plHY = 1.83;
    const lblX = 0.4;
    const lblW = nPer === 1 ? 4.5 : nPer === 2 ? 3.8 : 3.2;
    const valW = nPer === 1 ? 2.0 : nPer === 2 ? 2.2 : 1.85;
    const noteW = nPer === 1 ? 2.5 : 0;
    const commentX = 9.4;
    const commentW = W - commentX - 0.2;
    const valFontSize = nPer === 1 ? 13 : nPer === 2 ? 12 : 11;

    // ヘッダー行
    s.addShape("rect", { x: lblX, y: plHY, w: lblW, h: 0.36, fill: { color: DARK }, line: { color: WHITE } });
    s.addText("科目", { x: lblX + 0.05, y: plHY + 0.05, w: lblW - 0.1, h: 0.27, fontSize: 14, bold: true, color: WHITE, fontFace: "メイリオ" });
    plPeriods.forEach((per, pi) => {
      const colX = lblX + lblW + pi * valW;
      s.addShape("rect", { x: colX, y: plHY, w: valW, h: 0.36, fill: { color: DARK }, line: { color: WHITE } });
      // 期ラベル：括弧内を省いて短縮（例 "第10期（令和6年...）" → "第10期"）
      const perLabel = per.period ? per.period.replace(/（[^）]*）/, "").trim() : `第${pi + 1}期`;
      s.addText(perLabel, { x: colX + 0.02, y: plHY + 0.05, w: valW - 0.04, h: 0.27, fontSize: nPer === 1 ? 14 : 11, bold: true, color: WHITE, align: "center", fontFace: "メイリオ" });
    });
    if (noteW > 0) {
      const noteX = lblX + lblW + valW;
      s.addShape("rect", { x: noteX, y: plHY, w: noteW, h: 0.36, fill: { color: DARK }, line: { color: WHITE } });
      s.addText("備考", { x: noteX + 0.05, y: plHY + 0.05, w: noteW - 0.1, h: 0.27, fontSize: 14, bold: true, color: WHITE, fontFace: "メイリオ" });
    }

    // 行定義（key でデータを参照）
    const latestPl2 = (plPeriods[plPeriods.length - 1].pl) || {};
    const plRowDefs = [
      { label: "売上高",            key: "revenue",         sub: "",                                                                    bold: true,  fill: WHITE },
      { label: "　売上原価",         key: "cogs",            sub: latestPl2.grossMargin    ? `粗利率 ${latestPl2.grossMargin}`    : "", bold: false, fill: WHITE },
      { label: "売上総利益",         key: "grossProfit",     sub: "",                                                                    bold: true,  fill: LGRAY },
      { label: "　販売費及び一般管理費", key: "sga",          sub: "",                                                                    bold: false, fill: WHITE },
      { label: "営業利益",           key: "operatingProfit", sub: latestPl2.operatingMargin ? `利益率 ${latestPl2.operatingMargin}` : "", bold: true,  fill: GREEN_BG },
      { label: "　営業外収益",       key: "nonOpIncome",     sub: "",                                                                    bold: false, fill: WHITE },
      { label: "　営業外費用",       key: "nonOpExpense",    sub: "",                                                                    bold: false, fill: WHITE },
      { label: "経常利益",           key: "ordinaryProfit",  sub: latestPl2.ordinaryMargin ? `利益率 ${latestPl2.ordinaryMargin}` : "",  bold: true,  fill: GREEN_BG },
      { label: "法人税等",           key: "tax",             sub: "",                                                                    bold: false, fill: WHITE },
      { label: "当期純利益",         key: "netProfit",       sub: "",                                                                    bold: true,  fill: "C8E6D4" },
    ];

    plRowDefs.forEach((rowDef, i) => {
      const yy = plHY + 0.40 + i * 0.40;
      // ラベルセル
      s.addShape("rect", { x: lblX, y: yy, w: lblW, h: 0.38, fill: { color: rowDef.fill }, line: { color: MGRAY } });
      s.addText(rowDef.label, { x: lblX + 0.05, y: yy + 0.05, w: lblW - 0.1, h: 0.28, fontSize: 13, bold: rowDef.bold, color: BLACK, fontFace: "メイリオ" });
      // 各期の値セル
      plPeriods.forEach((per, pi) => {
        const colX = lblX + lblW + pi * valW;
        const val = (per.pl || {})[rowDef.key] || "";
        s.addShape("rect", { x: colX, y: yy, w: valW, h: 0.38, fill: { color: rowDef.fill }, line: { color: MGRAY } });
        s.addText(val, { x: colX + 0.03, y: yy + 0.05, w: valW - 0.06, h: 0.28, fontSize: valFontSize, bold: rowDef.bold, color: BLACK, align: "right", fontFace: "メイリオ" });
      });
      // 備考セル（1期のみ）
      if (noteW > 0) {
        const noteX = lblX + lblW + valW;
        s.addShape("rect", { x: noteX, y: yy, w: noteW, h: 0.38, fill: { color: rowDef.fill }, line: { color: MGRAY } });
        s.addText(rowDef.sub || "", { x: noteX + 0.05, y: yy + 0.05, w: noteW - 0.1, h: 0.28, fontSize: 12, color: DGRAY, fontFace: "メイリオ" });
      }
    });

    // 右側コメント
    s.addText("■ 分析コメント", {
      x: commentX, y: 1.83, w: commentW, h: 0.32, fontSize: 14, bold: true, color: BLACK, fontFace: "メイリオ",
    });
    const comments = plComments.length > 0 ? plComments : ["財務データを確認してください"];
    const commentText = comments.slice(0, 3).map(c2 => "• " + c2).join("\n\n");
    s.addText(commentText, {
      x: commentX, y: 2.2, w: commentW, h: 4.3,
      fontSize: 11.5, color: BLACK, fontFace: "メイリオ", valign: "top",
    });
  }

  // ========== 貸借対照表（最大売上期） ==========
  {
    const s = pptx.addSlide();
    s.background = { color: WHITE };
    const bsPeriodLabel = latestPeriod.period ? `貸借対照表（${latestPeriod.period.replace(/（[^）]*）/, "").trim()}）` : "貸借対照表";
    sectionTitle(s, bsPeriodLabel, GREEN, MGRAY);
    footer(pptx, s, 14 + fsOffset + pageOffset, GREEN);

    s.addText(`資産合計 ${bsLargest.totalAssets || "-"}、負債合計 ${bsLargest.totalLiabilities || "-"}、純資産 ${bsLargest.equity || "-"}`,
      { x: 0.5, y: 0.88, w: W - 1, h: 0.5, fontSize: 14, color: BLACK, fontFace: "メイリオ" });

    const assetRows = [
      { label: "【流動資産】", val: bsLargest.currentAssets || "", bold: true, fill: LGRAY },
      { label: "　現金・預金", val: bsLargest.cash || "", bold: false, fill: WHITE },
      { label: "　売掛金", val: bsLargest.receivables || "", bold: false, fill: WHITE },
      { label: "【固定資産】", val: bsLargest.fixedAssets || "", bold: true, fill: LGRAY },
      { label: "資産合計", val: bsLargest.totalAssets || "", bold: true, fill: "D4EDE1" },
    ];
    const liabRows = [
      { label: "【流動負債】", val: bsLargest.currentLiabilities || "", bold: true, fill: LGRAY },
      { label: "　短期借入金", val: bsLargest.shortTermLoan || "", bold: false, fill: WHITE },
      { label: "【固定負債】", val: bsLargest.fixedLiabilities || "", bold: true, fill: LGRAY },
      { label: "　長期借入金", val: bsLargest.longTermLoan || "", bold: false, fill: WHITE },
      { label: "負債合計", val: bsLargest.totalLiabilities || "", bold: true, fill: LGRAY },
      { label: "純資産合計", val: bsLargest.equity || "", bold: true, fill: GREEN_BG },
      { label: "負債・純資産合計", val: bsLargest.totalAssets || "", bold: true, fill: "D4EDE1" },
    ];

    const bsHY = 1.4;
    s.addShape("rect", { x: 0.4, y: bsHY, w: 4.0, h: 0.32, fill: { color: DARK }, line: { color: WHITE } });
    s.addShape("rect", { x: 4.4, y: bsHY, w: 1.9, h: 0.32, fill: { color: DARK }, line: { color: WHITE } });
    s.addText("資産の部", { x: 0.45, y: bsHY + 0.04, w: 3.9, h: 0.24, fontSize: 13, bold: true, color: WHITE, fontFace: "メイリオ" });
    s.addText("金額（円）", { x: 4.45, y: bsHY + 0.04, w: 1.8, h: 0.24, fontSize: 13, bold: true, color: WHITE, align: "right", fontFace: "メイリオ" });

    assetRows.forEach((row, i) => {
      const yy = bsHY + 0.36 + i * 0.36;
      s.addShape("rect", { x: 0.4, y: yy, w: 4.0, h: 0.34, fill: { color: row.fill }, line: { color: MGRAY } });
      s.addShape("rect", { x: 4.4, y: yy, w: 1.9, h: 0.34, fill: { color: row.fill }, line: { color: MGRAY } });
      s.addText(row.label, { x: 0.45, y: yy + 0.05, w: 3.9, h: 0.24, fontSize: 12.5, bold: row.bold, color: BLACK, fontFace: "メイリオ" });
      s.addText(row.val, { x: 4.45, y: yy + 0.05, w: 1.8, h: 0.24, fontSize: 12.5, bold: row.bold, color: BLACK, align: "right", fontFace: "メイリオ" });
    });

    s.addShape("rect", { x: 6.8, y: bsHY, w: 4.0, h: 0.32, fill: { color: DARK }, line: { color: WHITE } });
    s.addShape("rect", { x: 10.8, y: bsHY, w: 1.9, h: 0.32, fill: { color: DARK }, line: { color: WHITE } });
    s.addText("負債・純資産の部", { x: 6.85, y: bsHY + 0.04, w: 3.9, h: 0.24, fontSize: 13, bold: true, color: WHITE, fontFace: "メイリオ" });
    s.addText("金額（円）", { x: 10.85, y: bsHY + 0.04, w: 1.8, h: 0.24, fontSize: 13, bold: true, color: WHITE, align: "right", fontFace: "メイリオ" });

    liabRows.forEach((row, i) => {
      const yy = bsHY + 0.36 + i * 0.36;
      s.addShape("rect", { x: 6.8, y: yy, w: 4.0, h: 0.34, fill: { color: row.fill }, line: { color: MGRAY } });
      s.addShape("rect", { x: 10.8, y: yy, w: 1.9, h: 0.34, fill: { color: row.fill }, line: { color: MGRAY } });
      s.addText(row.label, { x: 6.85, y: yy + 0.05, w: 3.9, h: 0.24, fontSize: 12.5, bold: row.bold, color: BLACK, fontFace: "メイリオ" });
      const isNeg = String(row.val).includes("△") || String(row.val).includes("-");
      s.addText(row.val, { x: 10.85, y: yy + 0.05, w: 1.8, h: 0.24, fontSize: 12.5, bold: row.bold, color: isNeg ? GREEN : BLACK, align: "right", fontFace: "メイリオ" });
    });
  }

  // ========== 販管費内訳（最大売上期） ==========
  {
    const s = pptx.addSlide();
    s.background = { color: WHITE };
    const sgaPeriodLabel = latestPeriod.period ? `販売費及び一般管理費（${latestPeriod.period.replace(/（[^）]*）/, "").trim()}）` : "販売費及び一般管理費";
    sectionTitle(s, sgaPeriodLabel, GREEN, MGRAY);
    footer(pptx, s, 15 + fsOffset + pageOffset, GREEN);

    s.addText(`販管費合計 ${sgaLargest.personnel ? plLargest.sga || "-" : pl.sga || "-"}。主要コスト内訳。`,
      { x: 0.5, y: 0.88, w: W - 1, h: 0.42, fontSize: 14, color: BLACK, fontFace: "メイリオ" });

    // 人件費
    s.addShape("rect", { x: 0.4, y: 1.3, w: 2.5, h: 0.3, fill: { color: GREEN }, line: { color: GREEN } });
    s.addText(`【人件費】${sgaLargest.personnel || "-"}`, {
      x: 0.5, y: 1.32, w: 2.4, h: 0.26, fontSize: 13, bold: true, color: WHITE, fontFace: "メイリオ",
    });
    const jinRows = [
      ["役員報酬", sgaLargest.executivePay || ""],
      ["給料手当", sgaLargest.salary || ""],
      ["法定福利費", sgaLargest.socialInsurance || ""],
    ];
    jinRows.forEach((row, i) => {
      const yy = 1.65 + i * 0.34;
      s.addShape("rect", { x: 0.4, y: yy, w: 1.8, h: 0.32, fill: { color: i % 2 === 0 ? WHITE : LGRAY }, line: { color: MGRAY } });
      s.addShape("rect", { x: 2.2, y: yy, w: 1.7, h: 0.32, fill: { color: i % 2 === 0 ? WHITE : LGRAY }, line: { color: MGRAY } });
      s.addText(row[0], { x: 0.45, y: yy + 0.04, w: 1.7, h: 0.24, fontSize: 12.5, color: BLACK, fontFace: "メイリオ" });
      s.addText(row[1], { x: 2.25, y: yy + 0.04, w: 1.6, h: 0.24, fontSize: 12.5, color: BLACK, align: "right", fontFace: "メイリオ" });
    });

    // 経費
    s.addShape("rect", { x: 4.2, y: 1.3, w: 2.5, h: 0.3, fill: { color: "6BAE89" }, line: { color: "6BAE89" } });
    s.addText(`【主要経費】`, {
      x: 4.3, y: 1.32, w: 2.4, h: 0.26, fontSize: 13, bold: true, color: WHITE, fontFace: "メイリオ",
    });
    const keiRows = [
      ["広告宣伝費", sgaLargest.advertising || ""],
      ["賃借料", sgaLargest.rent || ""],
      ["減価償却費", sgaLargest.depreciation || ""],
      ["その他", sgaLargest.other || ""],
    ];
    const keiRowH = [0.34, 0.34, 0.34, 0.85];
    let keiY = 1.65;
    keiRows.forEach((row, i) => {
      const rh = keiRowH[i] || 0.34;
      s.addShape("rect", { x: 4.2, y: keiY, w: 2.0, h: rh, fill: { color: i % 2 === 0 ? WHITE : LGRAY }, line: { color: MGRAY } });
      s.addShape("rect", { x: 6.2, y: keiY, w: 4.5, h: rh, fill: { color: i % 2 === 0 ? WHITE : LGRAY }, line: { color: MGRAY } });
      s.addText(row[0], { x: 4.25, y: keiY + 0.05, w: 1.9, h: rh - 0.1, fontSize: 12.5, color: BLACK, fontFace: "メイリオ", valign: "top" });
      s.addText(row[1], { x: 6.25, y: keiY + 0.05, w: 4.4, h: rh - 0.1, fontSize: 12.5, color: BLACK, fontFace: "メイリオ", valign: "top" });
      keiY += rh + 0.02;
    });

  }

  // ========== Section V: 譲渡情報 ==========
  const po2 = pageOffset + (stores.length > 0 ? 1 : 0) + (menu.length > 0 ? 1 : 0);
  divSlide(pptx, "V.　譲渡情報", 16 + fsOffset + po2, GREEN);

  // ========== 希望譲渡条件 ==========
  {
    const s = pptx.addSlide();
    s.background = { color: WHITE };
    sectionTitle(s, "希望譲渡条件", GREEN, MGRAY);
    footer(pptx, s, 17 + fsOffset + po2, GREEN);

    // 希望譲渡金額：大きく目立たせる
    s.addShape("rect", { x: 0.5, y: 1.0, w: W - 1, h: 1.3, fill: { color: GREEN_BG }, line: { color: GREEN } });
    s.addText("希望譲渡金額", {
      x: 0.5, y: 1.08, w: W - 1, h: 0.32,
      fontSize: 14, bold: true, color: GREEN, align: "center", fontFace: "メイリオ",
    });
    s.addText(transfer.price || "応相談", {
      x: 0.5, y: 1.42, w: W - 1, h: 0.7,
      fontSize: 36, bold: true, color: BLACK, align: "center", fontFace: "メイリオ",
    });

    // 詳細テーブル
    const tRows = [
      ["譲渡スキーム", transfer.scheme || ""],
      ["譲渡理由", transfer.reason || ""],
      ["売却後の経営参加の意向", transfer.managementIntent || ""],
    ];
    const tRowH = [0.52, 0.75, 0.65];
    let tY = 2.55;
    tRows.forEach((row, i) => {
      const rh = tRowH[i];
      s.addShape("rect", { x: 0.5, y: tY, w: 2.8, h: rh, fill: { color: GREEN }, line: { color: GREEN } });
      s.addShape("rect", { x: 3.3, y: tY, w: W - 3.8, h: rh, fill: { color: WHITE }, line: { color: MGRAY } });
      s.addText(row[0], {
        x: 0.6, y: tY + 0.05, w: 2.6, h: rh - 0.1,
        fontSize: 14, bold: true, color: WHITE, fontFace: "メイリオ", valign: "middle",
      });
      s.addText(row[1], {
        x: 3.4, y: tY + 0.05, w: W - 3.9, h: rh - 0.1,
        fontSize: 14, color: BLACK, fontFace: "メイリオ", valign: "top",
      });
      tY += rh + 0.06;
    });
  }

  // ========== ノンネーム情報 ==========
  {
    const s = pptx.addSlide();
    s.background = { color: WHITE };
    sectionTitle(s, "ノンネーム情報", GREEN, MGRAY);
    footer(pptx, s, 18 + fsOffset + po2, GREEN);

    const nn = data.nonname || {};
    const notes = Array.isArray(nn.notes) ? nn.notes.filter(Boolean) : [];

    // 破線ボーダーボックス
    s.addShape("rect", {
      x: 0.5, y: 1.0, w: W - 1, h: H - 1.7,
      fill: { color: LGRAY },
      line: { color: MGRAY, width: 1, dashType: "dash" },
    });

    // 案件名ヘッダー
    s.addText(`【案件名：${nn.title || ""}】`, {
      x: 0.75, y: 1.1, w: W - 1.5, h: 0.42,
      fontSize: 15, bold: true, color: BLACK, fontFace: "メイリオ",
    });
    s.addShape("rect", { x: 0.75, y: 1.55, w: W - 1.5, h: 0.02, fill: { color: MGRAY }, line: { color: MGRAY } });

    // 基本情報
    const nnBasic = [
      ["事業内容", nn.business || ""],
      ["従業員数", nn.employees || ""],
    ];
    let nnY = 1.65;
    nnBasic.forEach((row) => {
      s.addText(row[0] + "：", {
        x: 0.85, y: nnY, w: 1.6, h: 0.36,
        fontSize: 13, bold: true, color: DGRAY, fontFace: "メイリオ",
      });
      s.addText(row[1], {
        x: 2.45, y: nnY, w: W - 3.2, h: 0.36,
        fontSize: 13, color: BLACK, fontFace: "メイリオ",
      });
      nnY += 0.38;
    });

    // 財務数値
    nnY += 0.1;
    s.addText("【財務数値】", {
      x: 0.85, y: nnY, w: W - 1.7, h: 0.34,
      fontSize: 14, bold: true, color: GREEN, fontFace: "メイリオ",
    });
    nnY += 0.38;
    const nnFin = [
      ["年商", nn.revenue || ""],
      ...(nn.profit ? [["修正後営業利益", nn.profit]] : []),
    ];
    nnFin.forEach((row) => {
      s.addText(row[0] + "：", {
        x: 1.05, y: nnY, w: 2.0, h: 0.34,
        fontSize: 13, bold: true, color: DGRAY, fontFace: "メイリオ",
      });
      s.addText(row[1], {
        x: 3.05, y: nnY, w: W - 3.8, h: 0.34,
        fontSize: 13, color: BLACK, fontFace: "メイリオ",
      });
      nnY += 0.36;
    });

    // 譲渡スキーム
    nnY += 0.1;
    s.addText("【譲渡スキーム】", {
      x: 0.85, y: nnY, w: W - 1.7, h: 0.34,
      fontSize: 14, bold: true, color: GREEN, fontFace: "メイリオ",
    });
    nnY += 0.38;
    const nnScheme = [
      { label: "譲渡方法",     val: nn.scheme || transfer.scheme || "", h: 0.34 },
      { label: "希望譲渡価格", val: nn.price || transfer.price || "応相談", h: 0.34 },
      { label: "譲渡理由",     val: nn.reason || transfer.reason || "", h: 0.62 },
      { label: "譲渡時期",     val: nn.timing || "応相談", h: 0.34 },
    ];
    nnScheme.forEach((row) => {
      s.addText(row.label + "：", {
        x: 1.05, y: nnY, w: 2.0, h: row.h,
        fontSize: 12.5, bold: true, color: DGRAY, fontFace: "メイリオ", valign: "top",
      });
      s.addText(row.val, {
        x: 3.05, y: nnY, w: W - 3.8, h: row.h,
        fontSize: 12.5, color: BLACK, fontFace: "メイリオ", valign: "top",
      });
      nnY += row.h + 0.04;
    });

    // 特記事項
    const MAX_Y = 6.65;
    if (notes.length > 0 && nnY + 0.84 < MAX_Y) {
      nnY += 0.1;
      s.addText("【特記事項】", {
        x: 0.85, y: nnY, w: W - 1.7, h: 0.34,
        fontSize: 14, bold: true, color: GREEN, fontFace: "メイリオ",
      });
      nnY += 0.38;
      notes.forEach((note) => {
        if (nnY + 0.34 > MAX_Y) return;
        s.addText("・" + note, {
          x: 1.05, y: nnY, w: W - 1.8, h: 0.34,
          fontSize: 12.5, color: BLACK, fontFace: "メイリオ", valign: "top",
        });
        nnY += 0.38;
      });
    }
  }

  // バッファとして返す
  const buffer = await pptx.write({ outputType: "nodebuffer" });
  return buffer;
}

module.exports = generateIM;
