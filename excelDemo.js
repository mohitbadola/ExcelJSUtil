const ExcelJs = require("exceljs");

const workbook = new ExcelJs.Workbook();
workbook.xlsx.readFile("downloads/exceldownloadTest.xlsx").then(function () {
  const worksheet = workbook.getWorksheet("Sheet1");

  worksheet.eachRow((row, rowNumber) => {
    row.eachCell((cell, colNumber) => {
      console.log(cell.value);
    });
  });
});
