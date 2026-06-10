# BORG — Bulk Organizer for Retail Galleries
### Files made simple
Built with Python | World of Powersports | Summer 2026

## Summary

BORG was built for World of Powersports to automate the process of organizing and uploading product photos to ChannelAdvisory, reducing the manual processing time from hours to minutes.
At its heart, this program is a product shoot organizer. It takes photos that you enter into a file that comes pre-installed with BORG. 
These photos must be taken in this order: a piece of paper or a black screen with the word 'SKU' at the top and the SKU number below it.
This process is repeated for each product, and as long as you take the photos in that order, Borg will do the rest. Borg will take your files and loop through them, checking for the word 'SKU'. 
When it detects the word, it will immediately stop and ask if the text it found on the image is the SKU number.
If the program did not read it correctly, then you have the opportunity to correct it, and the program will continue to add photos to a folder under the corrected SKU number.
The program will then add these photos to groups, which will be turned into folders in the output folder that is also created with your first launch of the program.
Along with this, it also uploads these files to an FTP server, where it then adds the pictures to the Channel Advisor, where they are distributed to product listings.
An Excel sheet is also required, called Inventory_Images_Record, in order to track the products that have been captured and put on the listings.

## System Requirements

- Python 3.x
- Tesseract OCR
- Windows OS

## Setup

1. Install Python from python.org
2. Install Tesseract from github.com/UB-Mannheim/tesseract/wiki
3. Run setup.bat
4. Add the excel file inventory_images_record to a share folder called BORG_Shared
5. Run BORG.exe

## How To Use

- Press Enter to process photos
- Type clear to clear the input folder
- Type update to update the Excel log

## Built With

- Pillow: Image Processing
- pytesseract: OCR
- paramiko: SFTP
- openpyxl: Excel automation
- pywin32: Windows shortcuts
- pillow-heif: HEIC support
