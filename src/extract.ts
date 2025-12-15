#!/usr/bin/env node
import * as fs from 'fs';
import { extractTablesFromWorkbook } from './extractService';

function main(): void {
    const argv = process.argv.slice(2);
    if (argv.length < 2) {
        console.error('Usage: npm run build && node dist/extract.js input.xlsx output.xlsx');
        process.exit(2);
    }
    const [inputPath, outputPath] = argv;

    if (!fs.existsSync(inputPath)) {
        console.error('Input file not found:', inputPath);
        process.exit(2);
    }

    const inputBuffer = fs.readFileSync(inputPath);
    const outputBuffer = extractTablesFromWorkbook(inputBuffer);
    fs.writeFileSync(outputPath, outputBuffer);
    console.log(`Wrote extracted tables to ${outputPath}`);
}

if (require.main === module) main();
