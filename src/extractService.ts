import * as xlsx from 'xlsx';

type Sheet = xlsx.WorkSheet;

function normalizeCellValue(cell: any): any {
    if (cell === undefined || cell === null) return null;
    if (typeof cell === 'object' && 'v' in cell) return cell.v;
    return cell;
}

function cellStr(cell: any): string {
    const v = normalizeCellValue(cell);
    if (v === null || v === undefined) return '';
    return String(v).trim();
}

function groupsForHeader(text?: string): Array<'immatricule' | 'nom' | 'nombre' | 'situation'> {
    const groups: Array<'immatricule' | 'nom' | 'nombre' | 'situation'> = [];
    if (!text) return groups;
    const s = text.toLowerCase();
    if (s.includes('n° immatricul') || s.includes('immatricul') || s.includes('ﺭﻗﻢ') || s.includes('رقم')) groups.push('immatricule');
    if (s.includes('nom') || s.includes('prénom') || s.includes('prenom')) groups.push('nom');
    if (s.includes('nombre') || s.includes('jours') || s.includes('jour')) groups.push('nombre');
    if (s.includes('situation')) groups.push('situation');
    return groups;
}

type TableExtraction = {
    rows: any[][];
    counts: { imm: number; nom: number; nb: number; sit: number };
};

function extractTablesFromSheetByHeaderNames(sheet: Sheet): TableExtraction[] {
    const out: TableExtraction[] = [];
    const range = sheet['!ref'];
    if (!range) return out;
    const decoded = xlsx.utils.decode_range(range);

    const startR = decoded.s.r;
    const endR = decoded.e.r;
    const startC = decoded.s.c;
    const endC = decoded.e.c;

    for (let r = startR; r <= endR; r++) {
        let rowHasHeader = false;
        for (let c = startC; c <= endC; c++) {
            const addr = xlsx.utils.encode_cell({ r, c });
            const text = cellStr(sheet[addr]);
            if (groupsForHeader(text).length > 0) {
                rowHasHeader = true;
                break;
            }
        }
        if (!rowHasHeader) continue;

        const immCols: number[] = [];
        const nomCols: number[] = [];
        const nbCols: number[] = [];
        const sitCols: number[] = [];

        for (let c = startC; c <= endC; c++) {
            const addr = xlsx.utils.encode_cell({ r, c });
            const text = cellStr(sheet[addr]);
            const gr = groupsForHeader(text);
            if (gr.includes('immatricule')) immCols.push(c);
            if (gr.includes('nom')) nomCols.push(c);
            if (gr.includes('nombre')) nbCols.push(c);
            if (gr.includes('situation')) sitCols.push(c);
        }

        if (immCols.length === 0 && nomCols.length === 0 && nbCols.length === 0 && sitCols.length === 0) continue;

        const rows: any[][] = [];
        let rr = r + 1;
        while (rr <= endR) {
            let allEmpty = true;
            const parts: any[] = [];

            const readCols = (cols: number[]) => {
                for (const cc of cols) {
                    const a = xlsx.utils.encode_cell({ r: rr, c: cc });
                    const v = sheet[a];
                    const s = cellStr(v);
                    if (s !== '') allEmpty = false;
                    parts.push(normalizeCellValue(v && v.v !== undefined ? v.v : (v ? v : null)));
                }
            };

            readCols(immCols);
            readCols(nomCols);
            readCols(nbCols);
            readCols(sitCols);

            if (allEmpty) break;
            rows.push(parts);
            rr++;
        }

        out.push({ rows, counts: { imm: immCols.length, nom: nomCols.length, nb: nbCols.length, sit: sitCols.length } });
    }

    return out;
}

export function extractTablesFromWorkbook(inputBuffer: Buffer): Buffer {
    const wb = xlsx.read(inputBuffer, { type: 'buffer', cellNF: false, cellDates: true });

    const allExtractions: TableExtraction[] = [];
    for (const name of wb.SheetNames) {
        const sheet = wb.Sheets[name];
        const tbls = extractTablesFromSheetByHeaderNames(sheet);
        for (const t of tbls) allExtractions.push(t);
    }

    if (allExtractions.length === 0) {
        throw new Error('No tables found by header names');
    }

    let maxImm = 0;
    let maxNom = 0;
    let maxNb = 0;
    let maxSit = 0;
    for (const ex of allExtractions) {
        if (ex.counts.imm > maxImm) maxImm = ex.counts.imm;
        if (ex.counts.nom > maxNom) maxNom = ex.counts.nom;
        if (ex.counts.nb > maxNb) maxNb = ex.counts.nb;
        if (ex.counts.sit > maxSit) maxSit = ex.counts.sit;
    }

    const combined: any[][] = [];
    const header: string[] = [];
    for (let i = 1; i <= maxImm; i++) header.push(`immatricule_${i}`);
    for (let i = 1; i <= maxNom; i++) header.push(`nom_${i}`);
    for (let i = 1; i <= maxNb; i++) header.push(`nombre_${i}`);
    for (let i = 1; i <= maxSit; i++) header.push(`situation_${i}`);
    combined.push(header);

    for (const ex of allExtractions) {
        for (const parts of ex.rows) {
            let idx = 0;
            const immPart = parts.slice(idx, idx + ex.counts.imm); idx += ex.counts.imm;
            const nomPart = parts.slice(idx, idx + ex.counts.nom); idx += ex.counts.nom;
            const nbPart = parts.slice(idx, idx + ex.counts.nb); idx += ex.counts.nb;
            const sitPart = parts.slice(idx, idx + ex.counts.sit); idx += ex.counts.sit;

            const pad = (arr: any[], size: number) => {
                const outArr = arr.slice();
                while (outArr.length < size) outArr.push(null);
                return outArr;
            };

            const row = [...pad(immPart, maxImm), ...pad(nomPart, maxNom), ...pad(nbPart, maxNb), ...pad(sitPart, maxSit)];
            combined.push(row.map((v) => (v === undefined ? null : v)));
        }
    }

    const outWb = xlsx.utils.book_new();
    const ws = xlsx.utils.aoa_to_sheet(combined);
    xlsx.utils.book_append_sheet(outWb, ws, 'Combined');

    return xlsx.write(outWb, { type: 'buffer', bookType: 'xlsx' }) as Buffer;
}
