# Barcodes Suite (Offline)

Two features, fully offline:

1) **Masterlist Checking**
   - Load/replace Masterlist (.xlsx) → stored in local SQLite.
   - Only digit runs **8–18 digits** are treated as barcodes.
   - Checks another Excel file against Masterlist and **highlights matches yellow**.
   - Reports unique counts and total occurrence counts.

2) **Existing Barcodes Checker**
   - Upload **Existing list** (.xlsx) and **Second file** (.xlsx with text+barcodes).
   - Removes only the barcode substrings from the Second file that **exist in the Existing list**.
   - Leaves other text in the cell intact; clears the cell if it becomes empty.
   - Optionally removes **cell fills** (highlights) from the Second file.
   - Reports occurrences removed vs. kept.

**Leading zeros are ignored for matching** (e.g., `0123456789012` equals `123456789012`).

---

## Build in the cloud (no Python on your Mac)

Use **GitHub Actions** to produce a standalone `.app` bundle you can download.

1. Create a repo and upload these files.
2. Go to the **Actions** tab → run the **Build Barcodes Suite for macOS** workflow.
3. Download the artifact zip for your architecture:
   - `Barcodes_Suite_macOS_AppleSilicon.zip` for M1/M2/M3
   - `Barcodes_Suite_macOS_Intel.zip` for Intel
4. Unzip → run `Barcodes Suite.app` (right‑click → Open on first run).

---

## Local run (if you have Python)

```bash
pip install -r requirements.txt
python barcodes_suite.py
```

---

## Packaging locally (optional)

```bash
pip install pyinstaller
pyinstaller --noconfirm --windowed --name "Barcodes Suite" barcodes_suite.py
```

Outputs: `dist/Barcodes Suite.app`
