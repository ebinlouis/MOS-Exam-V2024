const XLSX = require('xlsx');

// Load the Excel file
const workbook = XLSX.readFile('path/to/your/file.xlsx');

// Get the first sheet (you can specify a sheet name or index)
const sheetName = workbook.SheetNames[0];
const worksheet = workbook.Sheets[sheetName];

// Get the value of cell A1
const cellA1 = worksheet['A1'];

if (cellA1) {
    console.log(`Value of A1: ${cellA1.v}`);
} else {
    console.log('Cell A1 is empty or does not exist.');
}
