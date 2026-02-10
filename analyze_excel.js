const XLSX = require('xlsx-js-style');
const fs = require('fs');

const filePath = 'C:\\Users\\ramsa\\Downloads\\JAN Metric.xlsx';

try {
  if (!fs.existsSync(filePath)) {
    console.error(`File not found at ${filePath}`);
    process.exit(1);
  }

  const workbook = XLSX.readFile(filePath);
  const sheetNames = workbook.SheetNames;

  console.log(`Workbook Sheets: ${sheetNames.join(', ')}`);

  sheetNames.forEach(sheetName => {
    const worksheet = workbook.Sheets[sheetName];
    const data = XLSX.utils.sheet_to_json(worksheet, { header: 1 }); // Array of arrays
    
    console.log(`\n--- Sheet: ${sheetName} ---`);
    if (data.length > 0) {
      console.log('Headers:', JSON.stringify(data[0]));
      // Print first 2 rows of data if available
      if (data.length > 1) {
        console.log('Row 1:', JSON.stringify(data[1]));
      }
      if (data.length > 2) {
        console.log('Row 2:', JSON.stringify(data[2]));
      }
    } else {
      console.log('Empty Sheet');
    }
  });

} catch (err) {
  console.error('Error reading file:', err);
}
