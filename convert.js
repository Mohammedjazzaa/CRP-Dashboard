const XLSX = require('xlsx');
const fs = require('fs');

const wb = XLSX.readFile('Data base for Dashboard .xlsx');
const result = {};

wb.SheetNames.forEach(sheet => {
  result[sheet] = XLSX.utils.sheet_to_json(wb.Sheets[sheet]);
});

fs.writeFileSync('data.json', JSON.stringify(result, null, 2));
console.log('✅ Created data.json');
