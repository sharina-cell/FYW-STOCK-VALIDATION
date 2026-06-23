# FYW Stock Mismatch Checker

Compares the TC Product Master against marketplace stock files (Shopee, Lazada,
TikTok, Zalora) for Melissa, Ipanema, and CSpace, and produces a formatted Excel
report listing every SKU whose stock doesn't match.

## Setup (one-time)

1. **Install Python 3.8+** if you don't have it: https://www.python.org/downloads/

2. **Install required packages:**
   ```bash
   pip install pandas openpyxl
   ```

3. **Install LibreOffice** (needed only for Shopee files — their xlsx export
   format isn't readable directly by openpyxl):
   - Windows/Mac: https://www.libreoffice.org/download/download/
   - Linux: `sudo apt install libreoffice`

   After installing, make sure the command `soffice` (or `libreoffice`) works
   from a terminal/command prompt. On Windows you may need to add LibreOffice's
   `program` folder to your PATH.

## How to use

1. Put `fyw_stock_mismatch.py` in a folder together with:
   - Your TC Product Master file — named `TC_PRODUCT_MASTER.csv` or `ALL.csv`
   - Any marketplace files you want to check, named like:
     - `MELISSA_LAZADA.xlsx`, `IPANEMA_LAZADA.xlsx`, `CSPACE_LAZADA.xlsx`
     - `MELISSA_SHOPEE.xlsx`, `IPANEMA_SHOPEE.xlsx`, `CSPACE_SHOPEE.xlsx`
     - `MELISSA_TIKTOK.xlsx`, `IPANEMA_TIKTOK.xlsx`, `CSPACE_TIKTOK.xlsx`
     - `MELISSA_ZALORA.xlsx`
   - You don't need all of them — the script will just use whichever it finds.

2. Open a terminal in that folder and run:
   ```bash
   python fyw_stock_mismatch.py
   ```

3. It will print progress and create `Stock_Mismatch_Report.xlsx` in the same
   folder. Open it in Excel.

### Optional: run from a different folder / custom output name
```bash
python fyw_stock_mismatch.py --input-dir "C:\Users\you\Downloads\StockFiles" --output "MismatchReport.xlsx"
```

## What's in the report

- **ALL_MISMATCHES** sheet — one row per unique SKU with a mismatch anywhere,
  showing which marketplace(s) it's mismatched on.
- **One sheet per marketplace file** — every mismatched row for that file,
  color-coded:
  - 🔴 **Red** = marketplace stock is HIGHER than TC Master
  - 🟢 **Green** = marketplace stock is LOWER than TC Master

## Notes / known quirks the script already handles

- Filename typos like `CPSACE_LAZADA.xlsx` (instead of `CSPACE_LAZADA.xlsx`)
  are auto-corrected to CSpace.
- A `MELISSA_SHOPEE_COMBINED.xlsx` file (if present) is preferred over a plain
  `MELISSA_SHOPEE.xlsx` for the same brand/platform, since it's usually the
  merged/cleaned version.
- Trailing underscores in filenames (e.g. `MELISSA_ZALORA_.xlsx`) are fine.
- If a marketplace file fails to read, the script prints a warning and
  continues with the rest rather than crashing.

## Troubleshooting

**"LibreOffice not found on PATH"**
Shopee files specifically need LibreOffice installed for the conversion step.
Make sure you can run `soffice --version` (or `libreoffice --version`) from
your terminal. If not, LibreOffice isn't installed correctly or isn't on PATH.

**"Could not find TC Product Master file"**
Make sure your master file is named exactly `TC_PRODUCT_MASTER.csv` or `ALL.csv`
and is in the folder you're running the script from (or pointed to via `--input-dir`).

**A marketplace file didn't get picked up**
Check the filename starts with the brand (MELISSA/IPANEMA/CSPACE) and contains
the platform name (LAZADA/SHOPEE/TIKTOK/ZALORA) somewhere in it.
