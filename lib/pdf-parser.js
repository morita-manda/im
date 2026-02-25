const fs = require("fs");
const pdf = require("pdf-parse");

/**
 * PDFファイルからテキストを抽出する
 * @param {string} filePath - PDFファイルのパス
 * @returns {Promise<string>} 抽出されたテキスト
 */
async function parsePDF(filePath) {
  const dataBuffer = fs.readFileSync(filePath);
  const data = await pdf(dataBuffer);
  return data.text;
}

module.exports = parsePDF;
