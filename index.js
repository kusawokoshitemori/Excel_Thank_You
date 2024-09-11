const express = require("express");
const XLSX = require("xlsx");
const path = require("path");

const app = express();
const port = 3000;

// Excelファイルのパス
const excelFilePath = path.join(__dirname, "Thank_You_demo2.xlsx");

// Excelファイルを読み取る
const workbook = XLSX.readFile(excelFilePath);
const sheet_name_list = workbook.SheetNames;
const sheet = workbook.Sheets[sheet_name_list[0]];

// データを分かりやすく
const sheetData = XLSX.utils.sheet_to_json(sheet, { header: 1 });

// データの空白を無くして操作しやすいデータにする
function processData(data) {
  return data
    .map((row) => {
      return row.filter(
        (cell) => cell !== undefined && cell !== null && cell !== ""
      ); // 空白やundefined、nullを取り除く
    })
    .filter((row) => row.length > 0); // 空の行を取り除く
}

const processedData = processData(sheetData);
console.log("Processed Data:", processedData); // 処理後のデータをコンソールに表示

// keyword の次の文字を抽出する関数
function getValuesAfterKeywords(data, keywords) {
  let results = {};
  keywords.forEach((keyword) => {
    for (let row of data) {
      for (let cell of row) {
        if (cell && cell.startsWith(keyword)) {
          // キーワードが見つかった場合、その次のセルの値を取得
          const index = row.indexOf(cell);
          if (index !== -1 && index + 1 < row.length) {
            results[keyword] = row[index + 1];
            break; // 一度見つかったら次のキーワードに進む
          }
        }
      }
      if (results[keyword]) break; // キーワードが見つかったら次のキーワードに進む
    }
    if (!results[keyword]) {
      results[keyword] = null; // キーワードが見つからない場合
    }
  });
  return results;
}

// 複数のキーワードをリストとしてまとめる
const keywords = ["ShippedPer", "From", "To", "On or About", "Date", "No"];

// 各キーワードの次の文字を取得
const keywordValues = getValuesAfterKeywords(processedData, keywords);
console.log("Values after keywords:", keywordValues);

// 静的ファイルの提供
app.use(express.static(path.join(__dirname, "public")));

// データを取得するエンドポイント
app.get("/data", (req, res) => {
  res.json(processedData);
});

// HTMLページを表示するルート
app.get("/", (req, res) => {
  res.sendFile(path.join(__dirname, "public", "index.html"));
});

app.listen(port, () => {
  console.log(`Server is running at http://localhost:${port}`);
});
