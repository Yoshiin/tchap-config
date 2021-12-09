#!/usr/bin/env node
const yaml = require('js-yaml');
const fs = require('fs');
const ExcelJS = require('exceljs');

try {
	const doc = yaml.load(fs.readFileSync('./info-agent.yaml', 'utf8'));
	let list = doc.medium.email.patterns;
	let mObj = [];

	const workbook = new ExcelJS.Workbook();
	workbook.creator = 'Yoshin';
	const worksheet = workbook.addWorksheet('Email to Domain');

	let cellA = worksheet.getCell('A1');
	let cellB = worksheet.getCell('B1');
	cellA.value = "Email";
	cellB.value = "Domain";
	cellA.font = {bold: true};
	cellB.font = {bold: true};

	for (let i = 0; i < list.length; i++) {
		let tmpObj = {};
		let tmpKey = Object.keys(list[i]);
		let tmpKeyStr = tmpKey.toString();

		tmpKeyStr = tmpKeyStr.split('\\').join('');
		tmpKeyStr = tmpKeyStr.split('(.*.|)').join('*.');
		tmpKeyStr = tmpKeyStr.substring(1);

		tmpObj[tmpKeyStr] = list[i][Object.keys(list[i])].hs;
		mObj.push(tmpObj);

		let cellA = worksheet.getCell(`A${i + 2}`);
		let cellB = worksheet.getCell(`B${i + 2}`);
		cellA.value = tmpKeyStr;
		cellB.value = list[i][Object.keys(list[i])].hs;
	}

	workbook.xlsx.writeFile("info-agent.xlsx");
} catch (e) {
	console.log(e);
}
