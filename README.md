# BORG — Bulk Organizer for Retail Galleries
### Files Made Simple
Built with Python | World of Powersports | Summer 2026

## Summary

BORG was built for World of Powersports to automate the process of organizing and uploading product photos to ChannelAdvisor, reducing the manual processing time from hours to minutes.

At its heart, this program is a product shoot organizer. It takes photos from the product_input folder and organizes them into SKU labeled folders automatically.

For barcode suppliers (Kawasaki, Yamaha, Suzuki, Honda, Polaris, KTM, BRP, CF Moto), BORG uses the Cloudmersive Barcode API to automatically detect barcodes in photos and identify the product SKU — no divider card needed.

For text-based suppliers (Arctic Cat), BORG uses the Google Cloud Vision API to read the SKU printed on the product packaging.

For unlisted suppliers, BORG falls back to the original screenshot method — write the full SKU on a piece of paper or document, photograph it, and BORG reads it automatically using Google Cloud Vision with Tesseract OCR as a backup.

This process is repeated for each product. BORG groups the following photos under that SKU, renames them, uploads them to ChannelAdvisor via SFTP, and marks them as complete in the Excel tracking sheet.

## System Requirements

- Windows OS
- Python 3.x
- Tesseract OCR
- Google Cloud Vision API key
- Cloudmersive API key

## Setup

1. Install Python from python.org (check "Add to PATH")
2. Install Tesseract from github.com/UB-Mannheim/tesseract/wiki
3. Run setup.bat
4. Add inventory_images_record.xlsx to the BORG_Shared folder
5. Run BORG.exe
6. Enter FTP credentials when prompted
7. Enter Google Vision API key when prompted
8. Enter Cloudmersive API key when prompted

## Supported Suppliers

| Supplier | Method | Notes |
|---|---|---|
| Kawasaki | Barcode | Auto parsed |
| Yamaha | Barcode | Auto parsed |
| Suzuki | Barcode | Auto parsed |
| Honda | Barcode | Exact match |
| Polaris | Barcode | Exact match |
| KTM | Barcode | Exact match |
| BRP | Barcode | Auto parsed |
| CF Moto | Barcode | Auto parsed |
| Arctic Cat | Text OCR | Reads packaging text |
| Other | Screenshot | Write SKU on paper/doc |

## How To Use

- Press Enter — process photos
- Type 'update' — update Excel log without reprocessing

## Notes

- Only process ONE supplier brand per session
- BORG will warn you if a SKU group has more than 6 photos (possible missed detection)
- New SKUs not in the Excel sheet are added automatically
- If Excel file is open, close it before running BORG

## Built With

- Pillow — image processing and HEIC conversion
- pytesseract — OCR fallback
- Google Cloud Vision API — text detection
- Cloudmersive Barcode API — barcode detection
- paramiko — SFTP file uploads
- openpyxl — Excel automation
- pywin32 — Windows shortcut creation
- pillow-heif — HEIC file support