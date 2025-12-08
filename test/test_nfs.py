from pathlib import Path
import exifread

# cartella da leggere
p = Path("/mnt/nfs-peppe/Camera/da_sistemare/")

# scegli i tag che ti interessano (esempi comuni)
TAGS_OF_INTEREST = (
    "Image Make",
    "Image Model",
    "EXIF DateTimeOriginal",
    "EXIF DateTimeDigitized",
    "EXIF LensModel",
    "EXIF FNumber",
    "EXIF ExposureTime",
    "EXIF ISOSpeedRatings",
)

for img_path in sorted(p.glob("*.jpg")):
    print(f"\nFile: {img_path}")
    try:
        with img_path.open("rb") as fh:
            tags = exifread.process_file(fh, details=False, stop_tag="UNDEF")
        if not tags:
            print("  Nessun EXIF trovato")
            continue

        # stampa solo i tag di interesse, se presenti
        for tag in TAGS_OF_INTEREST:
            if tag in tags:
                print(f"  {tag}: {tags[tag]}")
        # se vuoi vedere tutti i tag decommenta qui sotto:
        # for k, v in tags.items():
        #     print(f"  {k}: {v}")

    except Exception as e:
        print("  Errore leggendo EXIF:", e)