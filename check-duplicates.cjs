const fs = require('fs');
const path = require('path');

const dataDir = './data';
const files = fs.readdirSync(dataDir).filter(f => f.startsWith('customers_') && f.endsWith('.json')).sort().reverse();
const latestFile = path.join(dataDir, files[0]);
const data = JSON.parse(fs.readFileSync(latestFile, 'utf-8'));

const plates = {};
data.forEach(c => {
  const plate = (c.licensePlate || '').trim().toLowerCase();
  if (plate) {
    if (!plates[plate]) plates[plate] = [];
    plates[plate].push(c);
  }
});

const dups = Object.keys(plates).filter(p => plates[p].length > 1).map(p => ({
  plate: p,
  count: plates[p].length,
  records: plates[p].map(c => ({name: c.customerName, row: c.rowNumber, phone: c.phone, policy: c.policyNumber}))
}));

let report = '=== DUPLICATE LICENSE PLATES ===\nTotal: ' + data.length + '\nUnique: ' + Object.keys(plates).length + '\nDuplicates: ' + dups.length + '\n\n';

dups.forEach((d, i) => {
  report += '--- Duplicate #' + (i+1) + ' ---\n';
  report += 'License: ' + d.plate + '\n';
  report += 'Count: ' + d.count + '\n\n';
  d.records.forEach((r, j) => {
    report += '  ' + (j+1) + '. ' + r.name + '\n';
    report += '     Row: ' + r.row + '\n';
    report += '     Phone: ' + r.phone + '\n';
    report += '     Policy: ' + (r.policy || 'N/A') + '\n\n';
  });
});

fs.writeFileSync('./duplicate-report.txt', report, 'utf-8');
console.log('Report saved to: duplicate-report.txt');
console.log('Found ' + dups.length + ' duplicate(s)');
