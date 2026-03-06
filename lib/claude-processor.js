const Anthropic = require("@anthropic-ai/sdk");
const fs = require("fs");

const client = new Anthropic({ apiKey: process.env.ANTHROPIC_API_KEY });

// JSONサニタイズ（共通）
function sanitizeJson(str) {
  str = str.replace(/,\s*([}\]])/g, "$1");
  str = str.replace(/[\x00-\x08\x0B\x0C\x0E-\x1F]/g, "");
  return str;
}

function parseJson(text) {
  let jsonStr = text.trim();
  const m1 = text.match(/```json[\r\n]+([\s\S]*?)[\r\n]+```/);
  const m2 = text.match(/```[\r\n]+([\s\S]*?)[\r\n]+```/);
  const m3 = text.match(/(\{[\s\S]*\})/);
  if (m1) jsonStr = m1[1];
  else if (m2) jsonStr = m2[1];
  else if (m3) jsonStr = m3[1];
  try {
    return JSON.parse(jsonStr);
  } catch {
    return JSON.parse(sanitizeJson(jsonStr));
  }
}

// 財務データの空テンプレート
function emptyPeriod() {
  return {
    period: "",
    pl: { revenue: "", cogs: "", grossProfit: "", grossMargin: "", sga: "", operatingProfit: "",
          operatingMargin: "", nonOpIncome: "", nonOpExpense: "", ordinaryProfit: "",
          ordinaryMargin: "", tax: "", netProfit: "", ebitda: "", depreciation: "" },
    bs: { currentAssets: "", cash: "", receivables: "", fixedAssets: "", totalAssets: "",
          currentLiabilities: "", shortTermLoan: "", fixedLiabilities: "", longTermLoan: "",
          totalLiabilities: "", equity: "", capital: "" },
    sga: { personnel: "", executivePay: "", salary: "", socialInsurance: "", rent: "",
           advertising: "", depreciation: "", entertainment: "", insurance: "", other: "" },
  };
}

/**
 * 1期分のPDFファイルから財務データを抽出する（Claude Document API使用）
 * スキャンPDF・テキストPDF両対応
 */
async function extractOnePeriod(filePath, index) {
  const pdfBase64 = fs.readFileSync(filePath).toString("base64");

  const instruction = `添付の決算書PDFから財務数値を読み取り、以下のJSON形式のみで返してください。説明文は不要です。

{
  "period": "第○期（令和○年○月期）",
  "pl": {
    "revenue": "売上高", "cogs": "売上原価", "grossProfit": "売上総利益", "grossMargin": "粗利率（例: 95.0%）",
    "sga": "販売費及び一般管理費", "operatingProfit": "営業利益", "operatingMargin": "営業利益率",
    "nonOpIncome": "営業外収益", "nonOpExpense": "営業外費用", "ordinaryProfit": "経常利益",
    "ordinaryMargin": "経常利益率", "tax": "法人税等", "netProfit": "当期純利益",
    "ebitda": "EBITDA（概算）", "depreciation": "減価償却費"
  },
  "bs": {
    "currentAssets": "流動資産", "cash": "現金・預金", "receivables": "売掛金",
    "fixedAssets": "固定資産", "totalAssets": "資産合計",
    "currentLiabilities": "流動負債", "shortTermLoan": "短期借入金",
    "fixedLiabilities": "固定負債", "longTermLoan": "長期借入金",
    "totalLiabilities": "負債合計", "equity": "純資産合計", "capital": "資本金"
  },
  "sga": {
    "personnel": "人件費", "executivePay": "役員報酬", "salary": "給料手当",
    "socialInsurance": "法定福利費", "rent": "賃借料", "advertising": "広告宣伝費",
    "depreciation": "減価償却費", "entertainment": "交際費", "insurance": "保険料", "other": "その他経費合計"
  }
}

数値は文字列（カンマ区切り）で。読み取れない項目は空文字。JSONのみを返してください。`;

  try {
    const message = await client.messages.create({
      model: "claude-sonnet-4-6",
      max_tokens: 4000,
      messages: [{
        role: "user",
        content: [
          {
            type: "document",
            source: { type: "base64", media_type: "application/pdf", data: pdfBase64 },
          },
          { type: "text", text: instruction },
        ],
      }],
    });

    const responseText = message.content[0].text.trim();
    console.log(`  [PDF ${index + 1}] レスポンス先頭:`, responseText.slice(0, 80));

    const result = parseJson(responseText);
    console.log(`  [PDF ${index + 1}] 抽出成功: period="${result.period}", revenue="${(result.pl || {}).revenue}"`);
    return result;
  } catch (e) {
    console.error(`  [PDF ${index + 1}] 財務抽出失敗（空データで続行）:`, e.message);
    return emptyPeriod();
  }
}

/**
 * PDF・HP情報をClaudeで構造化データに変換する
 * @param {{pdfTexts, webInfo, userInput, shareholders, employeeBreakdown}} input
 * @returns {Promise<Object>} 構造化されたIMデータ
 */
async function processWithClaude({ pdfFilePaths, pdfTexts, webInfo, userInput, shareholders, employeeBreakdown }) {
  const { tel } = userInput;
  const shareholdersText = (shareholders || []).map(s => `  - ${s.name}：${s.ratio}`).join("\n") || "  （未入力）";

  // フェーズ1: 各PDFファイルからClaudeのDocument APIで財務データを個別抽出
  const filePaths = pdfFilePaths && pdfFilePaths.length > 0 ? pdfFilePaths : [];
  let financialPeriods = [];
  if (filePaths.length > 0) {
    console.log(`[Claude] ${filePaths.length}期分のPDFをDocument APIで財務抽出中...`);
    financialPeriods = await Promise.all(filePaths.map((p, i) => {
      console.log(`  → PDF ${i + 1}/${filePaths.length} 抽出中`);
      return extractOnePeriod(p, i);
    }));
    console.log(`[Claude] 財務抽出完了: ${financialPeriods.length}期分`);
  }

  const latestPeriod = financialPeriods.length > 0 ? financialPeriods[financialPeriods.length - 1] : {};
  const latestPdfText = (pdfTexts && pdfTexts.length > 0) ? pdfTexts[pdfTexts.length - 1] : "";

  // フェーズ2: メインのIM生成（財務データは構造化済みのものを渡す）
  const financialPeriodsJson = JSON.stringify(financialPeriods, null, 2);

  const prompt = `
あなたはM&Aアドバイザーです。以下の情報から企業概要書（IM）用の構造化データをJSONで返してください。

## ホームページURL
${userInput.url}

## HP情報
タイトル: ${webInfo.title || "不明"}
説明: ${webInfo.description || "不明"}
本文: ${webInfo.bodyText || "不明"}
会社概要ページ本文: ${webInfo.companyPageText || "取得できず"}
ブランドカラー候補（HPのCSSから抽出）: ${webInfo.brandColors && webInfo.brandColors.length > 0 ? webInfo.brandColors.join(", ") : "なし"}

## HPからの情報抽出ルール
- **所在地（address）**: HP本文・会社概要ページから住所を必ず探して記載。「〒」「都」「道」「府」「県」「市」「区」「町」「村」などから特定。見つからない場合のみ空文字。
- **代表者（rep）**: HP本文・会社概要ページから役職付きで記載（例：「代表取締役 山田太郎」）。経営陣が複数いる場合は最初に記載されている人の役職と氏名のみ。見つからない場合のみ空文字。
- **電話番号（tel）**: ユーザー入力がない場合はHPから探す。
- **companyDetails（会社概要テーブル）**: 会社概要ページに記載されている項目をすべて {label, value} の配列で抽出する。商号・設立・資本金・所在地・代表者・事業目的・取引銀行・主要取引先・許認可・加盟団体・沿革など、HPに掲載されているものをそのまま列挙。company フィールドと重複していても構わない。HPに会社概要テーブルが見つからない場合は空配列。

## 最新期の決算書テキスト（PDF抽出）
${latestPdfText || "決算書なし"}

## ユーザー入力情報
会社名: ${userInput.companyName || "HPから取得"}
代表者名: ${userInput.repName || "HPまたはPDFから取得してください"}
電話番号: ${userInput.tel || "HPまたはPDFから取得してください"}
譲渡理由: ${userInput.reason || "未記入"}
希望譲渡金額: ${userInput.price || "未入力（下記の指示に従い試算してください）"}
譲渡スキーム: ${userInput.scheme || "未選択"}
売却後の経営参加の意向: ${userInput.managementIntent || "未選択"}

## 株主構成（ユーザー入力）
${shareholdersText}

## 従業員内訳
正社員: ${(employeeBreakdown && employeeBreakdown.full) || "未入力"}
パート・アルバイト: ${(employeeBreakdown && employeeBreakdown.part) || "未入力"}

## 【抽出済み】全期の財務データ
以下はすでに各期のPDFから抽出済みの財務データです。financialPeriodsにはこの値をそのまま使用してください。
${financialPeriodsJson}

${!userInput.price ? `## 譲渡希望金額の試算指示
希望譲渡金額が未入力です。抽出済み財務データをもとに、以下の手順で試算し transfer.price に記載してください。
- まず修正利益を算出：経常利益（または営業利益）＋ 交際費 ＋ 保険料（節税目的の費用を利益に戻す）
- 年買法：修正純資産 ＋ 修正利益 × 3〜5年分
- EBITDAマルチプル：（修正利益 ＋ 減価償却費）× 3〜5倍
- 上記2手法の結果をもとにレンジで記載
結果は「約○○百万円〜○○百万円（試算値）」の形式のみで記載。計算過程は含めないでください。
財務データが不足している場合は「応相談」としてください。

` : ""}## 重要：ノンネーム情報の注意点
nonnameフィールドは秘匿性が最重要です。以下の情報は絶対に含めないでください：
- 会社名・店舗名・ブランド名・サービス固有名
- 証券コード・法人番号・登記番号
- 具体的な住所・番地（都道府県や地域名のみ可）
- ホームページURL・メールアドレス・電話番号
- 対象企業を特定できる固有の数値や表現

## 出力形式（JSON）
以下のJSON構造で返してください。数値は文字列（カンマ区切り）で。不明な項目は空文字または推定値を記載。

{
  "company": {
    "name": "会社名",
    "address": "所在地",
    "established": "設立年月",
    "capital": "資本金",
    "rep": "役職 氏名（例：代表取締役 山田太郎）",
    "tel": "電話番号",
    "url": "URL",
    "business": "事業内容（2〜3文）",
    "employees": "従業員数",
    "fiscal": "決算期"
  },
  "stores": [
    {"name": "店舗名", "access": "アクセス", "tel": "電話", "pref": "都道府県"}
  ],
  "menu": [
    {"cat": "カテゴリ", "name": "メニュー名", "desc": "説明"}
  ],
  "financials": {抽出済み財務データの最新期をここにそのまま展開},
  "transfer": {
    "reason": "譲渡理由",
    "price": "希望譲渡金額",
    "scheme": "譲渡スキーム",
    "managementIntent": "売却後の経営参加の意向"
  },
  "strengths": [
    {"label": "強みラベル", "desc": "説明"}
  ],
  "market": "市場・競合状況（3〜5文）",
  "summary": "会社概要文（2〜3文、IM冒頭に使用）",
  "plComments": [
    "分析コメント1",
    "分析コメント2"
  ],
  "qa": [
    {"q": "質問", "a": "回答"}
  ],
  "nonname": {
    "title": "案件名（社名・証券コード・URL等の特定情報は絶対不可。業種と地域のみ。例：関東・まつげ専門サロン運営法人の譲渡）",
    "business": "事業内容（社名・固有ブランド名・商品名は不可。業種とサービス内容のみ。例：まつげエクステ施術・スクール運営）",
    "employees": "従業員数（概算のみ。例：代表＋正社員○名）",
    "revenue": "年商（概算のみ。例：~1億円）",
    "profit": "修正後営業利益または経常利益（概算。記載がなければ空文字）",
    "scheme": "譲渡方法（例：100%株式譲渡）",
    "price": "希望譲渡価格",
    "reason": "譲渡理由（一般的な表現のみ。社名・固有事情は不可。例：代表者の事業集中のため）",
    "timing": "応相談",
    "notes": ["特記事項（社名・固有名称は絶対不可。例：大手企業との業務委託実績有）"]
  },
  "companyDetails": [
    {"label": "項目名", "value": "値"}
  ],
  "financialPeriods": [抽出済み財務データをそのまま全期分配列に入れる],
  "themeColor": "ブランドカラー候補からこの会社に最も合う色を1色選び、6桁HEXコード（#なし）で。候補がなければ事業内容に合う色を選択。"
}

JSONのみを返してください。説明文は不要です。
`.trim();

  const message = await client.messages.create({
    model: "claude-sonnet-4-6",
    max_tokens: 8000,
    messages: [{ role: "user", content: prompt }],
  });

  const text = message.content[0].text.trim();
  let data;
  try {
    data = parseJson(text);
  } catch (e) {
    console.error("JSON解析失敗。stop_reason:", message.stop_reason);
    console.error("レスポンス末尾:", text.slice(-200));
    throw new Error("Claude APIのレスポンスをJSONとして解析できませんでした: " + e.message);
  }

  // 抽出済み財務データで確実に上書き（メインコールが改変しても安全）
  if (financialPeriods.length > 0) {
    data.financialPeriods = financialPeriods;
    data.financials = { ...latestPeriod };
  }

  // テーマカラーをHPの最多使用色で確定（Claude任せにしない）
  if (webInfo.dominantColor) {
    data.themeColor = webInfo.dominantColor;
  }

  // 意味のない値をクリア
  const unusableValues = ["不明", "不明確", "N/A", "-", "未記入", "不詳", "未取得", "HPまたはPDFから取得してください", "取得できず"];
  if (data.company) {
    if (unusableValues.some(v => (data.company.rep || "").trim() === v)) data.company.rep = "";
    if (unusableValues.some(v => (data.company.address || "").trim() === v)) data.company.address = "";
    if (unusableValues.some(v => (data.company.tel || "").trim() === v)) data.company.tel = "";
  }

  // null safety
  data.company = data.company || {};
  data.transfer = data.transfer || {};

  // userInputの値でオーバーライド（明示的に入力された場合）
  if (userInput.companyName) data.company.name = userInput.companyName;
  if (userInput.repName) data.company.rep = userInput.repName;
  if (userInput.tel) data.company.tel = userInput.tel;
  if (userInput.reason) data.transfer.reason = userInput.reason;
  if (userInput.price) data.transfer.price = userInput.price;
  // 価格未入力かつClaudeが試算値を返さなかった場合のフォールバック
  if (!userInput.price && !data.transfer.price) {
    data.transfer.price = "応相談";
  }
  if (userInput.scheme) data.transfer.scheme = userInput.scheme;
  if (userInput.managementIntent) data.transfer.managementIntent = userInput.managementIntent;
  data.company.url = userInput.url;

  return data;
}

module.exports = processWithClaude;
