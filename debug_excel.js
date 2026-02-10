const XLSX = require('xlsx-js-style');
const fs = require('fs');

const filePath = 'C:\\Users\\ramsa\\Downloads\\JAN Metric.xlsx';
const output = [];

try {
  if (!fs.existsSync(filePath)) {
    console.error(`File not found at ${filePath}`);
    process.exit(1);
  }

  const workbook = XLSX.readFile(filePath);
  const sheetNames = workbook.SheetNames;

  output.push(`Workbook Sheets: ${sheetNames.join(', ')}\n`);

  sheetNames.forEach(sheetName => {
    const worksheet = workbook.Sheets[sheetName];
    // limit rows to 5 for structure analysis
    const data = XLSX.utils.sheet_to_json(worksheet, { header: 1, range: 0, defval: "" }); 
    
    output.push(`\n--- Sheet: ${sheetName} ---`);
    if (data.length > 0) {
      // Headers
      output.push(`Headers (Row 1): ${JSON.stringify(data[0])}`);
      // First few data rows
      for(let i=1; i<Math.min(data.length, 4); i++) {
        output.push(`Row ${i+1}: ${JSON.stringify(data[i])}`);
      }
    } else {
      output.push('Empty Sheet');
    }
  });

  fs.writeFileSync('excel_structure.txt', output.join('\n'));
  console.log('Structure written to excel_structure.txt');

} catch (err) {
  console.error('Error:', err);
}
