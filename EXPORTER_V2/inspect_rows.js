const ExcelJS = require('exceljs');
const path = require('path');

async function inspect() {
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.readFile('Rekap_Ngawas_Lab.xlsx');
    const ws = wb.worksheets[0];
    
    console.log('--- Row Inspection ---');
    for (let r = 3; r <= 9; r++) {
        let values = [];
        for (let c = 1; c <= 8; c++) {
            const cell = ws.getCell(r, c);
            const val = cell.value ? (typeof cell.value === 'object' ? JSON.stringify(cell.value) : cell.value) : '';
            values.push(`C${c}: [${val}]`);
        }
        console.log(`Row ${r}: ${values.join(' ')}`);
    }
}

inspect().catch(console.error);
