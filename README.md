# Excel table extractor API

Express TypeScript API that extracts specific column groups from multiple tables inside an Excel file and combines them into a single sheet.

What it does
- Detects table headers by searching for header-like text (for example: "N° immatriculé", "Nom et prénom", "Nombre de jours", "Situation" — it also looks for Arabic forms like "ﺭﻗﻢ").
- Detects columns by header names (not fixed alphabet positions). Finds the header row, collects all columns whose header cell contains the target keywords (for each group: immatriculé, nom, nombre de jours, situation), and extracts those columns.
- Appends all extracted rows from all tables and all sheets into a single sheet named `Combined` in the output Excel.

Usage

## API Server

1. Install dependencies:

```bash
npm install
```

2. Build (TypeScript -> JavaScript):

```bash
npm run build
```

3. Start the server:

```bash
npm start
```

Server runs on `http://localhost:3000` by default.

### API Endpoints

**POST /extract**
- Upload an Excel file (multipart/form-data with key `file`)
- Returns the combined Excel file with extracted tables

Example using curl:
```bash
curl -X POST http://localhost:3000/extract \
  -F "file=@input.xlsx" \
  --output combined_output.xlsx
```

**GET /health**
- Health check endpoint
- Returns `{"status": "ok"}`

## CLI Usage (optional)

You can still use the CLI version:

```bash
npm run cli -- input.xlsx output.xlsx
```

Output
- `output.xlsx` will contain one sheet `Combined`. The first row is a generated header with columns immatricule_1..3, nom_1..4, nombre_1..5, situation_1..4.

Assumptions and notes
- The script expects each table's header to contain recognizable header text. It detects the header cell and treats it as the table's first column.
- Column mapping per table (relative to header start):
  - N° immatriculé => columns a, b, c (offsets 0,1,2)
  - Nom et prénom => columns d, e, f, g (offsets 3,4,5,6)
  - Nombre de jours => columns h, i, j, k, l (offsets 7..11)
  - Situation => columns m, n, o, p (offsets 12..15)
- If your tables start at different columns, the header detection will still pick up the start column provided the header cell contains one of the expected strings.
- If any table uses a different layout, or headers are formatted in an unusual way, adapt `looksLikeHeaderText` in `extract.js` to improve detection.

If you want, I can:
- Add more robust header detection (fuzzy match or a configurable list of header keywords).
- Support column letter configuration per table via a small JSON config.
# bds-maroc-spliter
