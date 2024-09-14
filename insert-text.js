const puppeteer = require("puppeteer");

// コマンドライン引数からデータを取得
const args = process.argv.slice(2);
const jsonData = JSON.parse(args[0]); // 引数は JSON 形式の文字列として受け取る

(async () => {
  const browser = await puppeteer.launch({ headless: false });
  const page = await browser.newPage();

  // URLを確認
  await page.goto("http://www.drpartners.jp/tools/browser-memocho.htm", {
    waitUntil: "networkidle2",
  }); // ページが完全に読み込まれるまで待つ

  // 最初の <textarea> 要素を選択
  const selector = "textarea:first-of-type";

  // データを組み合わせてテキストを作成
  const textToInsert = `
    Keyword Values: ${JSON.stringify(jsonData.keywordValues)}
    Consignee Related Rows: ${JSON.stringify(jsonData.consigneeRelatedRows)}
    Found Trade Words: ${JSON.stringify(jsonData.foundTradeWords)}
    JPY Values: ${JSON.stringify(jsonData.jpyValues)}
  `;

  await page.waitForSelector(selector); // 入力フィールドが読み込まれるまで待機
  await page.click(selector); // 入力フィールドをクリックしてフォーカスを当てる
  await page.keyboard.type(textToInsert); // テキストを入力する

  console.log("操作が完了しました。ブラウザを手動で閉じるまで待機中...");

  // ブラウザを閉じるまで手動で待つ
})();
