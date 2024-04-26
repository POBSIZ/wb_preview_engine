import express from "express";
import bodyParser from "body-parser";
import open from "open";
import readXlsxFile, { readSheetNames } from "read-excel-file/node";

import { readFile } from "./utils.js";

const app = express();
const port = 3000;

app.set("view engine", "ejs");
app.use(bodyParser.urlencoded({ extended: true }));

app.use("/static", express.static("public"));

app.get("/", (req, res) => {
  res.render("home");
});

app.get("/A", async (req, res) => {
  // Excel 파일의 경로

  const bookName = await readFile();
  const filePath = `./public/books/${bookName}/A/workbook.xlsx`;

  // 파일을 불러와서 workbook 객체 생성
  const workbook = readSheetNames(filePath).then((sheetNames) => {
    const slength = sheetNames.length;
    sheetNames.forEach((sheetName, i) => {
      readXlsxFile(filePath, { sheet: sheetName }).then((data) => {
        const columnName = data[0];
        if (req.query.type == "a1" && i === 0)
          return res.render("A/a1", { data: data.slice(1), bookName });
        if (req.query.type == "a2" && i === 1)
          return res.render("A/a2", { data: data.slice(1), bookName });
        if (req.query.type == "a3" && i === 2)
          return res.render("A/a3", { data: data.slice(1), bookName });
        if (req.query.type == "a4" && i === 3)
          return res.render("A/a4", { data: data.slice(1), bookName });
        if (req.query.type == "a5" && i === 4)
          return res.render("A/a5", { data: data.slice(1), bookName });
      });
    });
  });
});

app.get("/B", async (req, res) => {
  // Excel 파일의 경로

  const bookName = await readFile();
  const filePath = `./books/${bookName}/B/workbook.xlsx`;

  // 파일을 불러와서 workbook 객체 생성
  const workbook = readSheetNames(filePath).then((sheetNames) => {
    const slength = sheetNames.length;
    sheetNames.forEach((sheetName, i) => {
      readXlsxFile(filePath, { sheet: sheetName }).then((data) => {
        const columnName = data[0];
        if (i === 0) res.render("A", { data });
        // console.log(data[1]);
      });
    });
  });
});

app.post("/calculate", (req, res) => {
  const num1 = Number(req.body.num1);
  const num2 = Number(req.body.num2);
  const result = num1 + num2;

  res.render("home", { result: `결과는 ${result} 입니다.` });
});

app.listen(port, () => {
  console.log(`서버가 http://localhost:${port} 에서 실행중입니다.`);
  open(`http://localhost:${port}`);
});
