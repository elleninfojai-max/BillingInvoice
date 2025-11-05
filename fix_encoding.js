const fs = require('fs');
const path = require('path');

const filePath = path.join(__dirname, 'src', 'utils', 'exportUtilsDebit.js');
let content = fs.readFileSync(filePath, 'utf8');

// Fix encoding issues
content = content.replace(/Amount \(,1\)/g, 'Amount (₹)');
content = content.replace(/,1 /g, '₹ ');
content = content.replace(/Chennai.*?600094/g, 'Chennai – 600094');

fs.writeFileSync(filePath, content, 'utf8');
console.log('Fixed encoding issues in exportUtilsDebit.js');
