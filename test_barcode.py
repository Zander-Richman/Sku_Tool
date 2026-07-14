import os
import cloudmersive_barcode_api_client
from cloudmersive_barcode_api_client.rest import ApiException
from PIL import Image as PILImage
from pillow_heif import register_heif_opener
import tempfile

register_heif_opener()

configuration = cloudmersive_barcode_api_client.Configuration()
configuration.api_key['Apikey'] = os.getenv("CLOUDMERSIVE_KEY")

api_instance = cloudmersive_barcode_api_client.BarcodeScanApi(
    cloudmersive_barcode_api_client.ApiClient(configuration)
)

def convert_to_jpg(image_path):
    with PILImage.open(image_path) as img:
        img = img.convert('RGB')
        temp = tempfile.NamedTemporaryFile(suffix='.jpg', delete=False)
        img.save(temp.name,'JPEG')
        return temp.name

def test_cloudmersive(image_path):
    try:
        jpg_path = convert_to_jpg(image_path)
        response = api_instance.barcode_scan_image(jpg_path)
        os.remove(jpg_path)
        if response.successful:
            return response.raw_text
        return "NOT DETECTED"
    except ApiException as e:
        return f"ERROR: {e}"
    except Exception as e:
        return f"TIMEOUT/CONNECTION ERROR: {e}"
    
test_folder = "test_images"
for filename in os.listdir(test_folder):
    path = os.path.join(test_folder, filename)
    result = test_cloudmersive(path)
    print(f"{filename}: {result}")