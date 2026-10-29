// Reads the dispatch workbook (Detailed Dispatch Sheet + Total Order Summary) with the dashboard's own reader
// and writes public/data/dispatch.json and public/data/dispatch_material.json.
// Usage: node scripts/export_dispatch.js <workbook.xlsx> [outDir]
const fs = require('fs');
const path = require('path');
require('../dashboard/reports.js');
require('../dashboard/departments.js');
const R = globalThis.itoReports, Dp = globalThis.itoDepts;
const [file, outDir = path.join(__dirname, '..', 'public', 'data')] = process.argv.slice(2);
(async () => {
  const buf = fs.readFileSync(file);
  const sheets = (await R.readXlsx(buf.buffer.slice(buf.byteOffset, buf.byteOffset + buf.length))).map((s) => ({ name: s.name, ...R.toTable(s.aoa) }));
  const book = Dp.dispatchBook({ name: path.basename(file), sheets });
  if (!book) throw new Error('No dispatch sheet found in ' + file);
  fs.writeFileSync(path.join(outDir, 'dispatch.json'), JSON.stringify(book.rows));
  fs.writeFileSync(path.join(outDir, 'dispatch_material.json'), JSON.stringify(book.material));
  console.log(book.rows.length, 'dispatch rows from', JSON.stringify(book.sheet) + ',', book.material.length, 'material rows');
})().catch((e) => { console.error(e); process.exit(1); });
