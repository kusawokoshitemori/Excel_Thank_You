const express = require("express");
const XLSX = require("xlsx");
const path = require("path");

const app = express();
const port = 3000;

// Excelファイルを読み取る
const workbook = XLSX.readFile("Thank_You_demo.xlsx"); // 読み取りたいExcelファイルのパスを指定
const sheet_name_list = workbook.SheetNames;
const sheetData = XLSX.utils.sheet_to_json(workbook.Sheets[sheet_name_list[0]]);

// HTMLページを表示するルート
app.get("/", (req, res) => {
  let tableContent = sheetData
    .map((row) => {
      return `<tr>${Object.values(row)
        .map((cell) => `<td>${cell}</td>`)
        .join("")}</tr>`;
    })
    .join("");

  let html = `
    <!DOCTYPE html>
    <html lang="en">
    <head>
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <title>Excel Data</title>
    </head>
    <body>
        <table border="1">
            ${tableContent}
        </table>
    </body>
    </html>
    `;

  res.send(html);
});

app.listen(port, () => {
  console.log(`Server is running at http://localhost:${port}`);
});
