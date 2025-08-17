# Barcodes Suite (Offline) — v8 (GitHub Actions starter)

This repo builds a macOS app for Barcodes Suite v8 using GitHub Actions.

## Features (v8)
- Separators `,` `;` and ` - ` are read as **individual** barcodes.
- Hyphens inside one code (like `978-0-…`) are treated as formatting and **joined**.
- Reset buttons, per‑column top summaries, strong clearing of fills & conditional formats.
- Matching after normalization to **5–14 digits** (leading zeros ignored).
- Reads **.xlsx / .xlsm / .xls / .csv**; scans **all sheets**.

## How to use
1. Create a GitHub repo (public or private).
2. Upload these files at repo root:
   - `barcodes_suite.py`
   - `requirements.txt`
   - `.github/workflows/build-macos.yml`
3. Go to **Actions → Build Barcodes Suite for macOS → Run workflow**.
4. Download the artifact (Intel or AppleSilicon) and open the app on your Mac (right‑click → Open if Gatekeeper warns).
