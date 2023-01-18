const csv = require("csv-parse");
const fs = require("fs");
const path = require("path");

const inputFile = path.resolve("input/2_input_utf8.csv");

// const data = fs.readFileSync(inputFile, { encoding: "utf8" });
// // console.log(data.toString());
// const records = csv.parse(data.toString());
// console.log(records);

const dataArr = [];
fs.createReadStream(inputFile)
  .pipe(csv.parse())
  .on("data", (data) => {
    let number = data[7];
    number = number.replaceAll(",", "");
    const result = {
      header: `${data[4]} ${data[2]} (${data[6]}%) (${Math.round(
        parseInt(number) / 1000
      )}K)}`,
      거래량: `${data[7]}`,
    };
    dataArr.push(result);
  })
  .on("end", () => {
    console.log(dataArr);
  });
