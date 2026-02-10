const fs = require('fs');
const path = 'src/app/monthly-report/monthly-report.component.ts';
const content = fs.readFileSync(path, 'utf8');
fs.writeFileSync(path, content, 'utf8');
console.log('Scrubbed file');
