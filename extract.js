#!/usr/bin/env node
// extract.js
// Usage: node extract.js input.xlsx output.xlsx

const xlsx = require('xlsx');
const fs = require('fs');

function normalizeCellValue(cell) {
    if (cell === undefined || cell === null) return null;
    // If the cell is an object (from sheet), extract .v
    if (typeof cell === 'object' && 'v' in cell) return cell.v;
    return cell;
}

function cellStr(cell) {
    const v = normalizeCellValue(cell);
    if (v === null || v === undefined) return '';
    return String(v).trim();
}

function looksLikeHeaderText(text) {
    if (!text) return false;
    const s = String(text).toLowerCase();
    return (
        s.includes('n° immatricul') ||
        s.includes('nom et pr') ||
        s.includes('nombre de jours') ||
        s.includes('situation') ||
        s.includes('ﺭﻗﻢ') ||
        s.includes('رقم')
    );
}

function extractTablesFromSheet(sheet) {
    const outRows = [];
    const range = sheet['!ref'];
    if (!range) return outRows;
    const decoded = xlsx.utils.decode_range(range);

    const startR = decoded.s.r;
    const endR = decoded.e.r;
    const startC = decoded.s.c;
    const endC = decoded.e.c;

    for (let r = startR; r <= endR; r++) {
        for (let c = startC; c <= endC; c++) {
            const addr = xlsx.utils.encode_cell({ r, c });
            const cell = sheet[addr];
            const text = cellStr(cell);
            if (looksLikeHeaderText(text)) {
                // Found a header cell; treat this column as table start
                const headerRow = r;
                const tableStartCol = c;

                // offsets: a..p => 0..15
                const offsets = Array.from({ length: 16 }, (_, i) => i);

                // Collect rows until we hit an all-empty row for these columns
                let rr = headerRow + 1;
                while (rr <= endR) {
                    // check if all target cells are empty for the row
                    let allEmpty = true;
                    const extracted = [];
                    for (const off of offsets) {
                        const cc = tableStartCol + off;
                        const a = xlsx.utils.encode_cell({ r: rr, c: cc });
                        const v = sheet[a];
                        const s = cellStr(v);
                        if (s !== '') allEmpty = false;
                        // keep raw value (v ? v.v : null)
                        extracted.push(normalizeCellValue(v && v.v !== undefined ? v.v : (v ? v : null)));
                    }

                    if (allEmpty) break; // end of this table

                    // append extracted row
                    outRows.push(extracted);
                    rr++;
                }

                // move column pointer past this table (avoid double-detecting overlapping headers in same table)
                c = tableStartCol + 15;
            }
        }
    }

    return outRows;
}

function main() {
    const argv = process.argv.slice(2);
    if (argv.length < 2) {
        console.error('Usage: node extract.js input.xlsx output.xlsx');
        process.exit(2);
    }
    const [inputPath, outputPath] = argv;

    if (!fs.existsSync(inputPath)) {
        console.error('Input file not found:', inputPath);
        process.exit(2);
    }

    const wb = xlsx.readFile(inputPath, { cellNF: false, cellDates: true });
    const combined = [];

    // Add header row to combined output (optional) — we'll add simple column names
    const header = [];
    for (let i = 0; i < 16; i++) {
        // label groups: 0-2 -> immatricule, 3-6 -> nom, 7-11 -> nombre, 12-15 -> situation
        if (i <= 2) header.push(`immatricule_${i + 1}`);
        else if (i <= 6) header.push(`nom_${i - 2}`);
        else if (i <= 11) header.push(`nombre_${i - 6}`);
        else header.push(`situation_${i - 11}`);
    }
    combined.push(header);

    const sheetNames = wb.SheetNames;
    for (const name of sheetNames) {
        const sheet = wb.Sheets[name];
        const rows = extractTablesFromSheet(sheet);
        for (const r of rows) combined.push(r.map(v => (v === undefined ? null : v)));
    }

    // write output
    const outWb = xlsx.utils.book_new();
    const ws = xlsx.utils.aoa_to_sheet(combined);
    xlsx.utils.book_append_sheet(outWb, ws, 'Combined');
    xlsx.writeFile(outWb, outputPath);

    console.log(`Wrote ${combined.length - 1} data rows to ${outputPath}`);
}

if (require.main === module) main();
