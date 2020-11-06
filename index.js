#!/usr/bin/env node
"use strict";

const meow = require('meow');
const Excel = require("exceljs");
const fs = require("fs");

const cli = meow(`
	Usage
	  $ file-conversion <input>

	Options
    --input, -i  path to input file (cant be .json or .xlsx file)
    --output, -o output file name (can be .json or .xlsx file)

	Examples
	  $ foo unicorns --input=./input.json --output=out.xlsx
    This will convert given json file into excel file
    
	  $ foo unicorns --input=./input.xlsx --output=out.json
	  This will convert given excel file into json file
`, {
	flags: {
		input: {
			type: 'string',
      alias: 'i',
      isRequired: true
    },
    output: {
      type: 'string',
      alias: 'o',
      isRequired: true
    }
	}
});

function generateExcelFromJson(input, output) {
  console.log(`Generated Excel :: ${output}`);
  let workbook = new Excel.Workbook();
  let worksheet = workbook.addWorksheet("q&a");

  const jsonFile = input;
  const excelFile = output;

  let rawdata = fs.readFileSync(jsonFile);
  let json = JSON.parse(rawdata);
  Object.keys(json).forEach((k, i) => {
    const row = worksheet.getRow(i);
    row.getCell(1).value = k;
    row.getCell(2).value = json[k];
    row.commit();
  });

  workbook.xlsx.writeFile(excelFile);
  console.log(`Generated Excel :: ${output}`);
}

async function generateJsonFromExcel(input, output) {
  console.log(`Generating Json :: ${output}`);
  const json = {};
  const workbook = new Excel.Workbook();
  await workbook.xlsx.readFile(input);
  const worksheet = workbook.getWorksheet(1);
  worksheet.eachRow(function (row, rowNumber) {
    json[row.values[1]] = row.values[2];
  });
  const data = JSON.stringify(json);
  fs.writeFileSync(output, data);
  console.log(`Generated Json :: ${output}`);
}

const {input, output} = cli.flags;

if(input.includes('.json')) {
  generateExcelFromJson(input, output);
} else {
  generateJsonFromExcel(input, output);
}
