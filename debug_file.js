const fs = require('fs');
const path = 'src/app/monthly-report/monthly-report.component.ts';
const content = fs.readFileSync(path); // Buffer
if (content.includes(0)) {
    console.log('NULL BYTES FOUND!');
    // Print stats
    console.log('First null at:', content.indexOf(0));
} else {
    console.log('No null bytes found.');
}
// check for high chars
let high = 0;
for(const b of content) { if(b > 127) high++; }
console.log('Non-ASCII chars:', high);

console.log('Line 1892:', lines[1891]);
