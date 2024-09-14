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

// Consigneeが含まれる行の次にある3行の情報を取得する関数
function getThreeRowsAfterConsignee(data) {
  let consigneeRowIndex = -1;

  // Consigneeが含まれる行を見つける
  for (let i = 0; i < data.length; i++) {
    if (data[i].some((cell) => cell && cell.includes("Consignee"))) {
      consigneeRowIndex = i;
      break;
    }
  }

  // Consigneeの次にある3行を取得
  if (consigneeRowIndex !== -1 && consigneeRowIndex + 1 < data.length) {
    return data.slice(consigneeRowIndex + 1, consigneeRowIndex + 4);
  } else {
    return []; // Consigneeが見つからない、または次に行がない場合
  }
}

// エクセルのデータから trade_words の中にある単語を探し、見つかったら表示する関数
function findAndLogTradeWords(data, tradeWords) {
  data.forEach((row) => {
    row.forEach((cell) => {
      tradeWords.forEach((word) => {
        if (typeof cell === "string" && cell.includes(word)) {
          console.log(`Found trade word: ${word}`);
        }
      });
    });
  });
}

// 使用する貿易用語のリスト
const trade_words = ["FOB", "CIF", "EXY", "FCA", "DAP", "DDP"];

// "JPY"の次の数値を抽出する関数
function getJPYValueFromCell(cell) {
  const jpyPattern = /JPY\s*([0-9,]+)/; // "JPY"に続く数値を正規表現で抽出
  const match = cell.match(jpyPattern);
  if (match && match[1]) {
    // マッチした場合、その値を返す
    return match[1].replace(/,/g, ""); // カンマを取り除いて数値を返す
  }
  return null; // マッチしない場合はnullを返す
}

// エクセルのデータを走査し、"JPY"の次の数値を見つける関数
function findJPYValues(data) {
  data.forEach((row) => {
    row.forEach((cell) => {
      if (typeof cell === "string" && cell.includes("JPY")) {
        const jpyValue = getJPYValueFromCell(cell);
        if (jpyValue) {
          console.log(`Found JPY value: ${jpyValue}`);
        }
      }
    });
  });
}

// 複数のキーワードをリストとしてまとめる
const keywords = ["ShippedPer", "From", "To", "On or About", "Date", "No"];

// 各キーワードの次の文字を取得
const keywordValues = getValuesAfterKeywords(processedData, keywords);
console.log("Values after keywords:", keywordValues);

const consigneeRelatedRows = getThreeRowsAfterConsignee(processedData);
console.log("Rows after Consignee:", consigneeRelatedRows);

// trade_words を探して表示する
findAndLogTradeWords(processedData, trade_words);

// エクセルシートデータからJPYの次の値を探す
findJPYValues(processedData);

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
