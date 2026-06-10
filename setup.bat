@echo off
echo =====================================
echo BORG Setup - Files Made Simple
echo =====================================
echo.
echo Installing required Python libraries...
pip install pillow pytesseract pillow-heif paramiko openpyxl pywin32
echo.
echo Adding Tesseract to PATH...
setx PATH "%PATH%;C:\Program Files\Tesseract-OCR"
echo.
echo =====================================
echo Setup complete!
echo =====================================
echo.
echo IMPORTANT - Manual steps still required:
echo 1. Install Python from python.org if not already installed
echo 2. Install Tesseract from github.com/UB-Mannheim/tesseract/wiki
echo 3. Add inventory_images_record.xlsx to the BORG folder
echo 4. Double click BORG.exe to run
echo.
pause