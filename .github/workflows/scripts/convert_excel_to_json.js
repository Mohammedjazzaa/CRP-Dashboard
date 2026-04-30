const XLSX = require('xlsx');
const fs = require('fs');

// Read Excel file
const wb = XLSX.readFile('Data base for Dashboard .xlsx');
const data = {};

// Convert all sheets to JSON
wb.SheetNames.forEach(sheetName => {
  data[sheetName] = XLSX.utils.sheet_to_json(wb.Sheets[sheetName]);
});

// Save to data.json
fs.writeFileSync('data.json', JSON.stringify(data, null, 2));
console.log('✅ Conversion complete! File saved to data.json');
