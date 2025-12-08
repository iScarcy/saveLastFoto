from pathlib import Path
from PIL import Image
from PIL.ExifTags import TAGS

p = Path("/mnt/nfs-peppe/Camera/da_sistemare/")

for img_path in p.glob("*.jpg"):
    img = Image.open(img_path)
    exif = img._getexif()
    print(img_path, exif)
