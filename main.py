import os
import shutil
from PIL import Image
import pytesseract
from pillow_heif import register_heif_opener
import paramiko
from win32com.client import Dispatch
import sys



# -------SETTINGS-------

if getattr(sys, 'frozen', False):
    BASE_DIR = os.path.dirname(sys.executable)
else:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))  # Get the directory of the current script
EXCEL_PATH = os.path.join(os.environ['USERPROFILE'], 'OneDrive', 'Documents', 'BORG', 'BORG_Shared', 'inventory_images_record.xlsx')  # Path to save the Excel log file
SOURCE_FOLDER = os.path.join(BASE_DIR, "product_input")  # Default source folder
OUTPUT_FOLDER = os.path.join(BASE_DIR, "product_output")  # Default output folder

# Path to the logo image.... may need this later if we decide to add it
LOGO_PATH = os.path.join(BASE_DIR, "worldofpowersports_logo.png")  


def create_shortcut():
    try:
        desktop = os.path.join(os.environ['USERPROFILE'], 'OneDrive', 'Desktop')
        shortcut_path = os.path.join(desktop, 'BORG.lnk')
        print("finding shortcut...")

        if not os.path.exists(shortcut_path):
            shell = Dispatch('WScript.Shell')
            shortcut = shell.CreateShortCut(shortcut_path)
            shortcut.Targetpath = BASE_DIR
            shortcut.IconLocation = os.path.join(BASE_DIR, 'borg_icon.ico')
            shortcut.save()
            print("BORG shortcut created on Desktop! :)")
        else:
            print("BORG shortcut already exists on Desktop! :)")
    except Exception as e:
        print(f"Could not create shortcut: {e}")

def setup_ftp_credentials():
        print("FTP credentials not found -- first time setup required.")
        print("Please ask your supervisor for the FTP credentials.")
        print("")
        host = input("Enter FTP host: ").strip()
        user = input("Enter your FTP username: ").strip()
        password = input("Enter your FTP password: ")
        path = input("Enter the FTP path: ")

        os.system(f'setx FTP_HOST "{host}"')
        os.system(f'setx FTP_USER "{user}"')
        os.system(f'setx FTP_PASSWORD "{password}"')
        os.system(f'setx FTP_PATH "{path}"')

        print("FTP credentials saved! Please restart application for changes to take effect.")
        input("Press enter to exit...")
        raise SystemExit(0)

def first_time_setup():
    os.makedirs(SOURCE_FOLDER, exist_ok=True)
    os.makedirs(OUTPUT_FOLDER, exist_ok=True)
    create_shortcut()

    if not os.path.exists(EXCEL_PATH):
        print("ERROR: the file 'inventory_images_record.xlsx' does not exist in the BORG folder." \
        "You are not allowed to run this tool until you add this file.")
        input("Please seek assistance from your supervisor to obtain the 'inventory_images_record.xlsx' file and place it in the BORG folder, or await further instructions. Press enter to exit the application once you have read this message. ")
        raise SystemExit(0)
    
    if not os.getenv("FTP_HOST"):
        setup_ftp_credentials()

    print("product_input folder is ready")
    print("product_output folder is ready")
    print("Excel log file is ready")
    print("FTP credentials have been entered")
    setup_confirmation = input("Please place the photos if you have not done so already and restart the program, type 'exit' to quit and restart the program once the photos are placed. Otherwise, please press Enter to continue. ")
    if setup_confirmation == "":
        return
    elif setup_confirmation.lower() == "exit":
        raise SystemExit(0)


def get_jpgs(folder):
    valid_extentions = ['.jpg', '.jpeg']  # Add more valid extensions if needed
    jpgs= []
    for file in os.listdir(folder):
        ext = os.path.splitext(file)[1].lower()
        if ext in valid_extentions:
            jpgs.append(os.path.join(folder, file))
    jpgs.sort()
    return jpgs

def get_all_images(folder):
    for file in os.listdir(folder):
        print(f"file: {file}")
    valid_extentions = ['.jpg', '.jpeg', '.png', '.heic', '.webp', '.aae']  # Add more valid extensions if needed
    images = []
    for file in os.listdir(folder):
        ext = os.path.splitext(file)[1].lower()
        if ext in valid_extentions:
            images.append(os.path.join(folder, file))
    images.sort()
    return images
    

def convert_to_jpg(image_path):
    if image_path.lower().endswith('.jpg') or image_path.lower().endswith('.jpeg'):
        return image_path  # No conversion needed
    
    jpg_path = os.path.splitext(image_path)[0] + '.jpg'
    img = Image.open(image_path)
    img.convert('RGB').save(jpg_path, 'JPEG')
    print(f"Converted {image_path} to {jpg_path}")
    return jpg_path


def is_divider(image_path):
    img = Image.open(image_path)

    custom_config = r'--psm 6 -c min_characters_to_try=1'

    img = img.resize((800, 800))

    img = img.convert('L')  # Convert to grayscale

    text = pytesseract.image_to_string(img, config=custom_config).strip().upper()
    lines = text.splitlines()
    lines = [line.strip() for line in lines if line.strip()]

    if any('SKU' in line for line in lines):
        sku_index = next(i for i, line in enumerate(lines) if 'SKU' in line)
        if sku_index + 1 < len(lines):
            sku = lines[sku_index + 1].replace(' ', '').strip()  # Remove any unwanted characters
            sku = ''.join(c for c in sku if c.isalnum() or c in '_-')
            
            print(f"Detected SKU: {sku} in file {os.path.basename(image_path)}")
            confirm = input(f"Is this the correct SKU? (Detected: {sku}) [Press Enter to confirm or type the correct SKU]: ")
            if confirm != "":
                sku = confirm.strip().upper()
            
            return sku


def process_photos(photos):
    current_sku= None
    groups = {}

    for photo in photos: 
        sku = is_divider(photo) 
        if sku:
            current_sku = sku
            groups[sku] = []
        elif current_sku:
            groups[current_sku].append(photo)
    return groups


def export_folders(groups):
    # Create output folder if it doesn't exist
    os.makedirs(OUTPUT_FOLDER, exist_ok=True)

    for sku, photos in groups.items():
        sku_folder = os.path.join(OUTPUT_FOLDER, sku)
        os.makedirs(sku_folder, exist_ok=True)

        for i, photo in enumerate(photos, start=1):
            new_filename = f"{sku}__{i}.jpg"
            destination = os.path.join(sku_folder, new_filename)
            shutil.copy(photo, destination)

        if os.path.exists(LOGO_PATH):
            new_filename = f"{sku}__{len(photos) + 1}.jpg"
            logo_destination = os.path.join(sku_folder, new_filename)
            shutil.copy(LOGO_PATH, logo_destination)

    print(f"\nDONE! {len(groups)} SKU folders created in {OUTPUT_FOLDER}")


def clear_input_folder():
    for file in os.listdir(SOURCE_FOLDER):
        file_path = os.path.join(SOURCE_FOLDER, file)
        if os.path.isfile(file_path):
            os.remove(file_path)
    print(f"Cleared all files in {SOURCE_FOLDER}")


def upload_to_ftp(groups):
    FTP_HOST = os.getenv("FTP_HOST")
    FTP_USER = os.getenv("FTP_USER")
    FTP_PASS = os.getenv("FTP_PASSWORD")
    FTP_PATH = os.getenv("FTP_PATH")

    print(f"Uploading to FTP server: {FTP_HOST}")

    try:
        client = paramiko.SSHClient()
        client.set_missing_host_key_policy(paramiko.AutoAddPolicy())
        client.connect(hostname=FTP_HOST, port=22, username=FTP_USER, password=FTP_PASS)
        print(f"Connected to SSH server: {FTP_HOST}")
        sftp = client.open_sftp()
        sftp.chdir(FTP_PATH)
        print(f"Changed directory to: {FTP_PATH}")

        for sku in groups.keys():
            local_folder = os.path.join(OUTPUT_FOLDER, sku)
            if os.path.isdir(local_folder):
                for file in os.listdir(local_folder):
                    if file.endswith('.jpg'):
                        local_path = os.path.join(local_folder, file)
                        remote_path = f"{FTP_PATH}/{file}"
                        sftp.put(local_path, remote_path)
                        print(f"Uploaded: {file}")
        sftp.close()
        client.close()
        print("All files uploaded successfully!")
        return True
    except Exception as e:
        print(f"Failed to upload to FTP server: {e}")
        return False


def update_excel_log(groups, imported=False):
    import openpyxl

    try:
        wb = openpyxl.load_workbook(EXCEL_PATH)
        ws = wb.active
    except PermissionError:
        print("Please close out of excel and restart the program. Once restarted, type 'update' when prompted to do so. :)")
        return
    
    for row in ws.iter_rows(min_row=2):
        sku_cell = row[0]
        photos_taken_cell = row[4]
        photos_imported_cell = row[5]

        if sku_cell.value in groups:
            photos_taken_cell.value = "x"
            if imported == True:
                photos_imported_cell.value = "x"
                print(f"Marked SKU {sku_cell.value} as photographed and imported")
            else:
                print(f"Marked SKU {sku_cell.value} as photographed but not imported")


    wb.save(EXCEL_PATH)
    print(f"Excel log updated at: {EXCEL_PATH}")

#step 1: set up folders and check for Excel log file then start the process
register_heif_opener()
first_time_setup()
start = input("Press Enter to start processing photos, type 'clear' to clear the 'product_input' folder, or type 'update' to update the Excel log: ")


# Step 2 — Convert all photos to JPG
if start == "":
    print("BORG is running!!!")
    print(f"Looking for photos in: {SOURCE_FOLDER}")
    
    print("Converting all photos to JPG...")
    for photo in get_all_images(SOURCE_FOLDER):
        convert_to_jpg(photo)

    # Step 3 — Get only JPGs
    photos = get_jpgs(SOURCE_FOLDER)
    print(f"Found {len(photos)} JPG images")

    # Step 4 — Process them
    groups = process_photos(photos)
    print(f"\nFound {len(groups)} SKUs")
    for sku, photos in groups.items():
        print(f"{sku}: {len(photos)} photos")

    # Step 5 — create the folders with the images in the output folder
    export_folders(groups)

    # Step 6 - Upload to FTP
    upload_success = upload_to_ftp(groups)
    
    
    # Step 7 - Update Excel log
    update_excel_log(groups, imported=upload_success)
    
    
    # Step 8 — Clear source folder
    confirm_clear = input(f"Do you want to clear the source folder ({SOURCE_FOLDER})? This will delete all original photos. Type 'yes' to confirm: ")
    if confirm_clear.lower() == 'yes':
        clear_input_folder()

elif start.lower() == "clear":
    confirm_clear = input(f"Are you sure you want to clear the source folder ({SOURCE_FOLDER})? This will delete all original photos. Type 'yes' to confirm: ")
    if confirm_clear.lower() == 'yes':
        clear_input_folder()

elif start.lower() == "update":
    groups = {folder: [] for folder in os.listdir(OUTPUT_FOLDER) 
              if os.path.isdir(os.path.join(OUTPUT_FOLDER, folder))}
    did_upload = input("Did you already upload the photos to the FTP server? If yes, type 'yes' and press Enter to update the Excel log with imported status. If not, just press Enter to update without imported status: ")
    update_excel_log(groups, imported = (did_upload.lower() == 'yes'))


