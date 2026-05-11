const XLSX = require('xlsx');
const file = "C:\\Users\\Cassyano\\OneDrive - Brasil do Parafusos\\Comprasbrasil - Compras\\Análise Rupturas\\SIGER_XLS_1123274_070526_152305.xlsx";
const workbook = XLSX.readFile(file);
const sheet = workbook.Sheets[workbook.SheetNames[0]];
const headers = XLSX.utils.sheet_to_json(sheet, { header: 1 })[0];
console.log(headers);
