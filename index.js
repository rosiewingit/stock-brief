const csv = require("csv-parse");
const iconv = require("iconv-lite");
const jschardet = require("jschardet");
const fs = require("fs");
const path = require("path");
const excel = require("excel4node");

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
      if (data.length == 10) {
        if (data[1] === "종목명") {
          return;
        }

        let number = data[6];
        number = number.replaceAll(",", "");
        const result = {
          header: `${data[3]} ${data[1]} (${data[5]}) (${Math.round(
            parseInt(number) / 1000
          )}K)`,
          거래량: `${data[6]}`,
        };
        dataArr.push(result);
      } else {
        // 장중상승
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
      }
    })
    .on("end", () => {
      console.log("Start to create an excel.");
      console.log(outputFile);
      console.log(dataArr);
      console.log("Completed to create an excel.\n\n");
      exportExcel(dataArr, outputFile);
    });
};

const exportExcel = (dataset, outputFile) => {
  const book = new excel.Workbook({});
  const sheet = book.addWorksheet("stock");

  const headerStyle = book.createStyle({
    font: {
      bold: true,
    },
  });

  const underlineStyle = book.createStyle({
    font: {
      underline: true,
    },
  });

  const highlightStyle = book.createStyle({
    font: {
      color: "red",
    },
  });

  sheet.cell(1, 1).string("header").style(headerStyle);
  sheet.cell(1, 2).string("거래량").style(headerStyle);
  let i = 0;
  for (i = 0; i < dataset.length; i++) {
    let header = dataset[i].header;
    let number = dataset[i]["거래량"];

    if (checkCell(number)) {
      sheet
        .cell(i + 2, 1)
        .string(header)
        .style(underlineStyle);
      sheet
        .cell(i + 2, 2)
        .string(number)
        .style(highlightStyle);
    } else {
      sheet.cell(i + 2, 1).string(header);
      sheet.cell(i + 2, 2).string(number);
    }
  }

  book.write(outputFile);
};

const checkCell = (value) => {
  let modified = value.replaceAll(",", "");
  let intValue = parseInt(modified);
  if (intValue >= 10000000) {
    return true;
  }
  return false;
};
