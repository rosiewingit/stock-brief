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
      console.log(dataArr);
      exportExcel(dataArr, outputFile);
    });
};

const exportExcel = (dataset, outputFile) => {
  const book = xlsx.utils.book_new();
  const data = xlsx.utils.json_to_sheet(dataset);
  data["!cols"] = [{ wpx: 200 }, { wpx: 70 }];

  // for (let i in data) {
  //   data[`B${i}`].s = {
  //     fill: {
  //       patternType: "solid",
  //       bgColor: { rgb: "FFFFAA00" },
  //     },
  //   };
  // }

  // for (let i = 2; i < data.length; i++) {
  //   // data[`B${i}`].s = {
  //   // fill: {
  //   //   patternType: "solid",
  //   //   bgColor: { rgb: "FFFFAA00" },
  //   // },
  //   // };
  //   let tmpValue = data[`B${i}`].v;
  //   console.log("tmpValue: ", tmpValue);
  //   let modified = tmpValue.replaceAll(",", "");
  //   console.log("modified: ", modified);
  //   let intValue = parseInt(modified);
  //   if (intValue >= 10000000) {
  //     data[`B${i}`].v = `▲ ${tmpValue}`;
  //     console.log("data[`B${i}`].v: ", data[`B${i}`].v);
  //   }
  // }

  xlsx.utils.book_append_sheet(book, data, "stock");
  xlsx.writeFile(book, outputFile);
};
