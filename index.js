const csv = require("csv-parse");
const iconv = require("iconv-lite");
const jschardet = require("jschardet");
const xlsx = require("xlsx");
const fs = require("fs");
const path = require("path");

const inputDir = "input";
const outputDir = "output";

fs.readdir(inputDir, (err, filelist) => {
  if (err) {
    console.log("error", err);
  }
  filelist.forEach((file) => {
    const name = file.replace(".csv", "");
    const input = path.resolve(inputDir, `${name}.csv`);
    const output = path.resolve(outputDir, `${name}.xlsx`);
    run(input, output);
  });
});

const run = (inputFile, outputFile) => {
  const content = fs.readFileSync(inputFile, "binary");
  const charset = jschardet.detect(content);
  if (charset.encoding === "EUC-KR") {
    const utf8Text = iconv.decode(content, "euc-kr");
    fs.writeFileSync(inputFile, utf8Text, "utf-8");
  }

  const dataArr = [];
  fs.createReadStream(inputFile)
    .pipe(csv.parse())
    .on("data", (data) => {
      console.log("data: ", data);
      if (data[2] === "종목명") {
        return;
      }

      let number = data[7];
      number = number.replaceAll(",", "");
      const result = {
        header: `${data[4]} ${data[2]} (${data[6]}%) (${Math.round(
          parseInt(number) / 1000
        )}K)`,
        거래량: `${data[7]}`,
      };
      dataArr.push(result);
    })
    .on("end", () => {
      console.log(dataArr);
      exportExcel(dataArr, outputFile);
    });
};

const exportExcel = (dataset, outputFile) => {
  const book = xlsx.utils.book_new();
  const data = xlsx.utils.json_to_sheet(dataset);
  xlsx.utils.book_append_sheet(book, data);
  xlsx.writeFile(book, outputFile);
};
