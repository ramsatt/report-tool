const fs = require('fs');
const path = 'src/app/monthly-report/monthly-report.component.ts';
let content = fs.readFileSync(path, 'utf8');
// Normalize newlines to \n
content = content.replace(/\r\n/g, '\n').replace(/\r/g, '\n');
fs.writeFileSync(path, content, 'utf8');
console.log('Normalized newlines');
