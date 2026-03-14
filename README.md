# Customs PDF Extractor

Extracts highlighted fields from Indian Customs EDI System shipping bills into Excel.

## For end users

1. Download `CustomsExtractor.exe` from the [latest release](../../releases/latest)
2. Create a `key.txt` file in the same folder as the exe
3. Paste your Anthropic API key into `key.txt` and save
4. Double-click `CustomsExtractor.exe`
5. Select your PDF files, select output folder — done

## For developers

### Setup
```bash
pip install -r requirements.txt
python src/extract_customs.py
```

### Building the exe locally (Windows only)
```bash
pyinstaller --onefile --windowed --name "CustomsExtractor" src/extract_customs.py
# Output: dist/CustomsExtractor.exe
```

### Building via GitHub Actions
Push to `main` — the workflow automatically:
1. Builds the exe on a Windows runner
2. Creates a GitHub Release with the exe attached

You can also trigger it manually:
`Actions → Build Windows EXE → Run workflow`

## What gets extracted

**Sheet 1 — Part I Summary**
Port Code, SB No, SB Date, FOB Value, Freight, Insurance, Discount, COM,
Deductions, P/C, Duty, Cess, DBK Claim, IGST AMT, Cess AMT, IGST Value,
RODTEP AMT, ROSCTL AMT, INV No, INV AMT, Currency

**Sheet 2 — Part II Val Dtls**
Invoice Value, FOB Value, Freight, Insurance, Discount, Commission,
Deduct, P/C, Exchange Rate

**Sheet 3 — Part II Items**
Item No, HS Code, Description, Quantity, Rate (USD), Value F/C (USD)
