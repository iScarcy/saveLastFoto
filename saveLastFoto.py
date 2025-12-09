import os
from datetime import datetime
import configparser
import traceback
from pathlib import Path
import shutil
import exifread
from openpyxl import Workbook

import smtplib
from email.mime.text import MIMEText

# -----------------------
# 1. LETTURA CONFIGURAZIONE
# -----------------------
config = configparser.ConfigParser()
config.read("config.ini")

share_input = config["Share"]["share_input"].strip()
share_output = config["Share"]["share_output"].strip()

move_file = config["Elaborazione"].getboolean("moveFile", fallback=True)
send_mail = config["Mail"].getboolean("sendMail", fallback=False)
allowed_exts = [e.strip().lower() for e in config["Elaborazione"]["allowedExtensions"].split(",")]

smtp_server = config["Mail"]["smtpServer"]
smtp_port = int(config["Mail"]["smtpPort"])
smtp_user = config["Mail"]["smtpUser"]
smtp_password = config["Mail"]["smtpPassword"]
mail_from = config["Mail"]["mailFrom"]
mail_to = config["Mail"]["mailTo"]

# Converti a Path objects
input_path = Path(share_input)
output_path = Path(share_output)

print(f"DEBUG: Lettura da {input_path} -> Scrivi a {output_path}")

# Verifica che le cartelle esistono
if not input_path.exists():
    print(f"ERRORE: Cartella input non trovata: {input_path}")
    exit(1)
if not output_path.exists():
    print(f"ERRORE: Cartella output non trovata: {output_path}")
    exit(1)

# -----------------------
# 2. Mesi in italiano
# -----------------------
mesi_italiano = {
    "01":"Gennaio","02":"Febbraio","03":"Marzo","04":"Aprile",
    "05":"Maggio","06":"Giugno","07":"Luglio","08":"Agosto",
    "09":"Settembre","10":"Ottobre","11":"Novembre","12":"Dicembre"
}

# -----------------------
# 3. Funzione estrazione EXIF con exifread
# -----------------------
def estrai_metadati(file_path):
    """
    Estrae i metadati EXIF da un file immagine usando exifread.
    Ritorna (metadati_dict, errore_msg) dove metadati_dict è None se errore.
    """
    try:
        with file_path.open("rb") as fh:
            tags = exifread.process_file(fh, details=False, stop_tag="UNDEF")
        
        if not tags:
            return None, "Nessun EXIF presente"
        
        metadati = {}
        # Cerca il tag della data (comune: DateTimeOriginal)
        data_creazione = None
        for tag in ("EXIF DateTimeOriginal", "EXIF DateTimeDigitized", "Image DateTime"):
            if tag in tags:
                data_creazione = str(tags[tag])
                break
        
        metadati["data_creazione"] = data_creazione
        # Salva anche altri tag utili se presenti
        for tag in ("Image Make", "Image Model", "EXIF LensModel", "EXIF FNumber"):
            if tag in tags:
                metadati[tag] = str(tags[tag])
        
        return metadati, None
    except Exception as e:
        return None, str(e)

# -----------------------
# 4. Preparazione log e Excel
# -----------------------
oggi = datetime.now().strftime("%Y-%m-%d")
excel_file = f"{oggi}.xlsx"
log_file = f"{oggi}_log.txt"
log_lines = []

def log_write(txt):
    log_lines.append(txt)

log_write("=== CONFIGURAZIONE ===")
for section in config.sections():
    log_write(f"[{section}]")
    for k,v in config[section].items():
        log_write(f"{k} = {v}")
    log_write("")

tempo_esecuzione = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
log_write(f"Esecuzione: {tempo_esecuzione}")
log_write("")

# -----------------------
# 5. Lettura file da NFS
# -----------------------
all_files = sorted([f.name for f in input_path.iterdir() if f.is_file()])
file_list=[]
ignored_files=[]

for f in all_files:
    ext=os.path.splitext(f)[1].lower()
    if ext in allowed_exts:
        file_list.append(f)
    else:
        ignored_files.append(f)

log_write(f"File ignorati (estensione non ammessa): {len(ignored_files)}")
for ig in ignored_files:
    log_write(f"- {ig}")
log_write("")

# -----------------------
# 6. Elaborazione file
# -----------------------
dati_excel=[]
count_ok=0
count_err=0

try:
    for file_name in file_list:
        file_path = input_path / file_name

        metadati, err_exif = estrai_metadati(file_path)

        if metadati and metadati.get("data_creazione"):
            try:
                dt=datetime.strptime(metadati["data_creazione"],"%Y:%m:%d %H:%M:%S")
                anno=dt.strftime("%Y")
                mese_num=dt.strftime("%m")
                mese_nome=mesi_italiano[mese_num]
            except:
                metadati=None
                err_exif="Formato data EXIF non valido"

        if metadati is None:
            count_err+=1
            dati_excel.append({
                "data_esecuzione": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                "file": file_name,
                "anno":"unknown",
                "mese":"unknown",
                "destinazione":"",
                "errore":err_exif if err_exif else "",
                "esito_spostamento":"non eseguito"
            })
            continue

        # Percorso destinazione foto
        month_folder=f"{mese_num} {mese_nome}"
        dest_dir = output_path / anno / month_folder
        dest_path = dest_dir / file_name

        # Crea cartelle se non esistono
        dest_dir.mkdir(parents=True, exist_ok=True)

        if move_file:
            try:
                shutil.copy2(file_path, dest_path)
                file_path.unlink()  # elimina il file originale
                esito="ok"
                count_ok+=1
            except Exception as e:
                esito=f"errore spostamento: {e}"
                count_err+=1
        else:
            esito="test_mode_nessuno_spostamento"

        dati_excel.append({
            "data_esecuzione": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "file": file_name,
            "anno": anno,
            "mese": month_folder,
            "destinazione": str(dest_path.relative_to(output_path)),
            "errore": err_exif if err_exif else "",
            "esito_spostamento": esito
        })

except Exception as exc:
    log_write("=== ERRORE NON GESTITO ===")
    log_write(str(exc))
    log_write(traceback.format_exc())
    with open(log_file,"w",encoding="utf-8") as f:
        f.write("\n".join(log_lines))
    print("Errore grave. Dettagli nel Log.")
    exit(1)

# -----------------------
# 7. Salvataggio Excel con openpyxl
# -----------------------
anno_corrente=datetime.now().strftime("%Y")
excel_dir = output_path / anno_corrente
excel_dir.mkdir(parents=True, exist_ok=True)
excel_path = excel_dir / excel_file

wb = Workbook()
ws = wb.active
ws.title = "Elaborazione"

# Intestazione
headers = ["data_esecuzione", "file", "anno", "mese", "destinazione", "errore", "esito_spostamento"]
ws.append(headers)

# Dati
for row_data in dati_excel:
    ws.append([row_data.get(h, "") for h in headers])

# Salva il file
wb.save(str(excel_path))

# -----------------------
# 8. Log finale
# -----------------------
log_write(f"Totale file analizzati: {len(file_list)}")
log_write(f"File spostati correttamente: {count_ok}")
log_write(f"File con errore (non spostati): {count_err}")
with open(log_file,"w",encoding="utf-8") as f:
    f.write("\n".join(log_lines))

# -----------------------
# 9. Invio mail
# -----------------------
if send_mail:
    try:
        body = (
            f"Esecuzione script: {tempo_esecuzione}\n"
            f"Totale file analizzati: {len(file_list)}\n"
            f"File spostati correttamente: {count_ok}\n"
            f"File con errore (non spostati): {count_err}\n"
            f"File ignorati (estensione non permessa): {len(ignored_files)}"
        )
        msg = MIMEText(body)
        msg["Subject"]=f"Report SaveLastFoto - {tempo_esecuzione}"
        msg["From"]=mail_from
        msg["To"]=mail_to

        smtp=smtplib.SMTP(smtp_server,smtp_port)
        smtp.starttls()
        smtp.login(smtp_user,smtp_password)
        smtp.sendmail(mail_from,[mail_to],msg.as_string())
        smtp.quit()
    except Exception as e:
        log_write("\nErrore invio mail:")
        log_write(str(e))
        with open(log_file,"w",encoding="utf-8") as f:
            f.write("\n".join(log_lines))

print("✔ Script completato.")
