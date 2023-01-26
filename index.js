const csv = require("csv-parse");
const iconv = require("iconv-lite");
const jschardet = require("jschardet");
const fs = require("fs");
const path = require("path");
const excel = require("excel4node");
const fetch = require("sync-fetch");
const figlet = require("figlet");

// open dart
const ACCESS_TOKEN = "7fbfd9d3f1ac4b90767562d3de65bf56f621a36b";
const hostUrl = "https://opendart.fss.or.kr/api";

// input / output
const inputDir = "input";
const outputDir = "output";

console.log(
  figlet.textSync("MUZINSTOCK", {
    font: "Standard",
    horizontalLayout: "default",
    verticalLayout: "default",
    width: 80,
    whitespaceBreak: true,
  })
);

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
          //data[0] 종목코드
          return;
        }

        let number = data[6];
        number = number.replaceAll(",", "");
        console.log(getCompanyDart(data[0]));
        const result = {
          header: `${data[3]} ${data[1]} (${data[5]}) (${Math.round(
            parseInt(number) / 1000
          )}K)`,
          거래량: `${data[6]}`,
          사업보고서: `${getCompanyDart(data[0])}`,
        };
        dataArr.push(result);
      } else {
        // 장중상승
        if (data[2] === "종목명") {
          //data[1] 종목코드
          return;
        }

        let number = data[7];
        number = number.replaceAll(",", "");
        console.log(getCompanyDart(data[1]));
        const result = {
          header: `${data[4]} ${data[2]} (${data[6]}%) (${Math.round(
            parseInt(number) / 1000
          )}K)`,
          거래량: `${data[7]}`,
          사업보고서: `${getCompanyDart(data[1])}`,
        };
        dataArr.push(result);
      }
    })
    .on("end", () => {
      console.log("Start to create an excel.");
      !fs.existsSync(outputDir) && fs.mkdirSync(outputDir);
      console.log(outputFile);
      console.log(dataArr);
      exportExcel(dataArr, outputFile);
      console.log("Completed to create an excel.\n\n");
    });
};

const exportExcel = (dataset, outputFile) => {
  console.log(dataset);
  const book = new excel.Workbook({});
  const sheet = book.addWorksheet("stock", {
    pageSetup: {
      fitToWidth: 3,
    },
  });

  const headerStyle = book.createStyle({
    font: {
      bold: true,
    },
  });

  const underlineStyle = book.createStyle({
    font: {
      bold: true,
      underline: true,
    },
  });

  const highlightStyle = book.createStyle({
    font: {
      color: "red",
    },
  });

  sheet.column(1).setWidth(40);
  sheet.column(3).setWidth(56);
  sheet.cell(1, 1).string("header").style(headerStyle);
  sheet.cell(1, 2).string("거래량").style(headerStyle);
  sheet.cell(1, 3).string("사업보고서").style(headerStyle);
  let i = 0;
  for (i = 0; i < dataset.length; i++) {
    let header = dataset[i].header;
    let number = dataset[i]["거래량"];
    let doc = dataset[i]["사업보고서"];

    if (checkCell(number)) {
      sheet
        .cell(i + 2, 1)
        .string(header)
        .style(underlineStyle);
      sheet
        .cell(i + 2, 2)
        .string(number)
        .style(highlightStyle);
      sheet.cell(i + 2, 4).string(doc);
      sheet.cell(i + 2, 3).formula(`HYPERLINK(D${i + 2},D${i + 2})`);
    } else {
      sheet.cell(i + 2, 1).string(header);
      sheet.cell(i + 2, 2).string(number);
      sheet.cell(i + 2, 4).string(doc);
      sheet.cell(i + 2, 3).formula(`HYPERLINK(D${i + 2},D${i + 2})`);
    }
  }
  sheet.column(4).hide();
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

const getCompanyDart = (corpCode) => {
  let code = `${corpCode.replace("'", "").trim()}`;

  let searchUrl = `${hostUrl}/list.json?crtfc_key=${ACCESS_TOKEN}&corp_code=${code}&bgn_de=20220101&last_reprt_at=Y&pblntf_ty=A&pblntf_detail_ty=A001`;
  const body = fetch(searchUrl).json();

  if (!body) {
    return;
  }
  const data = body;
  if (data.status == "000") {
    const list = data.list;
    let result;
    if (list.length > 0) {
      list.forEach((item) => {
        if (item.report_nm.includes("사업보고서")) {
          result = item;
          return;
        }
      });
      if (result) {
        const link = `https://dart.fss.or.kr/dsaf001/main.do?rcpNo=${result.rcept_no}`;
        return link;
      }
    }
  }
};
