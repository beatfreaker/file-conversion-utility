"use strict";
const Excel = require("exceljs");
const fs = require("fs");

async function process() {
  const json = {};
  const workbook = new Excel.Workbook();
  await workbook.xlsx.readFile("resource_hindi.xlsx");
  const worksheet = workbook.getWorksheet(1);
  worksheet.eachRow(function (row, rowNumber) {
    console.log(row.values[2]);
    json[row.values[1]] = row.values[2];
  });
  const data = JSON.stringify(json);
  fs.writeFileSync("resource_hindi.json", data);
}

process();
