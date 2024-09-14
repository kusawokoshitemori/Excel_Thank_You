const puppeteer = require("puppeteer");

(async () => {
  const browser = await puppeteer.launch({ headless: false });
  const page = await browser.newPage();

  // URLを確認
  await page.goto("http://www.drpartners.jp/tools/browser-memocho.htm", {
    waitUntil: "networkidle2",
  }); // ページが完全に読み込まれるまで待つ

  // 最初の <textarea> 要素を選択
  const selector = "textarea:first-of-type";

  const textToInsert = "お肉冷蔵庫に入れるの忘れてた"; // 挿入するテキスト

  await page.waitForSelector(selector); // 入力フィールドが読み込まれるまで待機
  await page.click(selector); // 入力フィールドをクリックしてフォーカスを当てる
  await page.keyboard.type(textToInsert); // テキストを入力する

  console.log("操作が完了しました。ブラウザを手動で閉じるまで待機中...");

  // ブラウザを閉じるまで手動で待つ
})();
