const ExcelJS = require('exceljs');
const path = require('path');

async function main() {
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.readFile(path.join(__dirname, '..', 'xlsx', 'Barbe.xlsx'));
    const ws = wb.worksheets[0];
    console.log('Sheet Name:', ws.name);
    console.log('Columns:');
    const r3 = ws.getRow(3);
    const r4 = ws.getRow(4);
    const r5 = ws.getRow(5);
    for (let c = 1; c <= ws.columnCount; c++) {
        const cat = r3.getCell(c).value;
        const q = r4.getCell(c).value;
        const t = r5.getCell(c).value;
        if (cat || q || t) {
            console.log(`Col ${c}: Cat [${cat}], Q [${q}], Type [${t}]`);
        }
    }
    
    console.log('\nData Rows (first 3):');
    ws.eachRow((row, rIdx) => {
        if (rIdx >= 6 && rIdx <= 8) {
            const vals = [];
            for (let c = 1; c <= ws.columnCount; c++) {
                vals.push(row.getCell(c).value);
            }
            console.log(`Row ${rIdx}:`, vals.slice(0, 10));
        }
    });
}
main().catch(console.error);
