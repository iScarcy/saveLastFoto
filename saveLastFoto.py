import os
import pandas as pd
from PIL import Image, ExifTags
from datetime import datetime
import configparser
import traceback
from io import BytesIO
import uuid
import sys

from smbprotocol.connection import Connection
from smbprotocol.session import Session
from smbprotocol.tree import TreeConnect
from smbprotocol.open import Open
from smbprotocol.create_contexts import CreateDisposition, CreateOptions

import smtplib
from email.mime.text import MIMEText

# -----------------------
# 1. LETTURA CONFIGURAZIONE
# -----------------------
config = configparser.ConfigParser()
config.read("config.ini")

server = config["Samba"]["server"]
share_input = config["Samba"]["share_input"]
share_output = config["Samba"]["share_output"]
username = config["Samba"]["username"]
password = config["Samba"]["password"]

# Legge client_guid da config (opzionale). Se presente, lo converte in uuid.UUID.
client_guid_cfg = config["Samba"].get("client_guid", "").strip()
if client_guid_cfg:
    try:
        client_guid = uuid.UUID(client_guid_cfg)
    except Exception:
        client_guid = uuid.uuid4()
else:
    client_guid = uuid.uuid4()

# Normalizza share: se config contiene un percorso locale (es. /mnt/...), usa il basename
def normalize_share(s):
    s = s.strip()
    if not s:
        return s
    if s.startswith("/") or s.startswith("\\") or ":" in s:
        return os.path.basename(s.rstrip("/\\"))
    return s

share_input = normalize_share(share_input)
share_output = normalize_share(share_output)

move_file = config["Elaborazione"].getboolean("moveFile", fallback=True)
send_mail = config["Mail"].getboolean("sendMail", fallback=False)
allowed_exts = [e.strip().lower() for e in config["Elaborazione"]["allowedExtensions"].split(",")]

smtp_server = config["Mail"]["smtpServer"]
smtp_port = int(config["Mail"]["smtpPort"])
smtp_user = config["Mail"]["smtpUser"]
smtp_password = config["Mail"]["smtpPassword"]
mail_from = config["Mail"]["mailFrom"]
mail_to = config["Mail"]["mailTo"]

# -----------------------
# 2. CONNESSIONE SMB
# -----------------------
# Usa il client GUID (uuid.UUID). Se non presente in config, usa uuid4().
print(f"DEBUG: Connecting to server={server} guid={client_guid} share_in={share_input} share_out={share_output}")
sys.stdout.flush()
conn = Connection(guid=client_guid, server_name=server, port=445)
conn.connect()

session = Session(conn, username, password)
session.connect()

tree_in = TreeConnect(session, fr"\\{server}\{share_input}")
tree_in.connect()

tree_out = TreeConnect(session, fr"\\{server}\{share_output}")
tree_out.connect()

# -----------------------
# Funzioni utility SMB
# -----------------------
def smb_listdir(tree):
    f = Open(tree, "", access=0x01)
    f.create(disposition=CreateDisposition.FILE_OPEN, options=CreateOptions.FILE_DIRECTORY_FILE)
    files = [entry["file_name"].get_value() for entry in f.query_directory("*") if entry["file_name"].get_value() not in [".",".."]]
    f.close()
    return files

def smb_open_file(tree, filepath):
    f = Open(tree, filepath, access=0x120089)
    f.create(disposition=CreateDisposition.FILE_OPEN, options=CreateOptions.FILE_NON_DIRECTORY_FILE)
    return f

def smb_copy_file(tree_in, path_in, tree_out, path_out):
    f_in = smb_open_file(tree_in, path_in)
    data = f_in.read(8*1024*1024)
    f_in.close()
    f_out = Open(tree_out, path_out, access=0x12019f)
    f_out.create(disposition=CreateDisposition.FILE_OVERWRITE_IF, options=CreateOptions.FILE_NON_DIRECTORY_FILE)
    f_out.write(data,0)
    f_out.close()

def smb_delete(tree, path):
    f = Open(tree, path, access=0x10000)
    f.create(disposition=CreateDisposition.FILE_OPEN, options=CreateOptions.FILE_NON_DIRECTORY_FILE)
    f.delete()
    f.close()

def smb_mkdirs(tree, path):
    parts = path.replace("\\","/").split("/")
    cur=""
    for p in parts:
        if not p: continue
        cur=cur+"/"+p
        try:
            d=Open(tree, cur, access=0x01)
            d.create(disposition=CreateDisposition.FILE_OPEN, options=CreateOptions.FILE_DIRECTORY_FILE)
            d.close()
        except:
            d=Open(tree, cur, access=0x01)
            d.create(disposition=CreateDisposition.FILE_CREATE, options=CreateOptions.FILE_DIRECTORY_FILE)
            d.close()

# -----------------------
# 3. Mesi in italiano
# -----------------------
mesi_italiano = {
    "01":"Gennaio","02":"Febbraio","03":"Marzo","04":"Aprile",
    "05":"Maggio","06":"Giugno","07":"Luglio","08":"Agosto",
    "09":"Settembre","10":"Ottobre","11":"Novembre","12":"Dicembre"
}

# -----------------------
# 4. Funzione estrazione EXIF
# -----------------------
def estrai_metadati(img_bytes):
    try:
        img = Image.open(img_bytes)
        exif_raw = img._getexif()
        if exif_raw is None:
            return None, "Nessun EXIF presente"
        exif = {}
        for tag, val in exif_raw.items():
            key = ExifTags.TAGS.get(tag, tag)
            exif[key]=val
        data = exif.get("DateTimeOriginal") or exif.get("DateTime")
        return {"data_creazione": data, **exif}, None
    except Exception as e:
        return None, str(e)

# -----------------------
# 5. Preparazione log e Excel
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
# 6. Lettura file da samba
# -----------------------
all_files = smb_listdir(tree_in)
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
# 7. Elaborazione file
# -----------------------
dati_excel=[]
count_ok=0
count_err=0

try:
    for file_name in file_list:
        fh = smb_open_file(tree_in, file_name)
        bytes_data = fh.read(8*1024*1024)
        fh.close()
        bio=BytesIO(bytes_data)

        metadati, err_exif = estrai_metadati(bio)

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
        dest_dir=os.path.join(anno, month_folder).replace("\\","/")
        dest_path=dest_dir+"/"+file_name
        smb_mkdirs(tree_out, dest_dir)

        if move_file:
            try:
                smb_copy_file(tree_in, file_name, tree_out, dest_path)
                smb_delete(tree_in, file_name)
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
            "destinazione": dest_path,
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
# 8. Salvataggio Excel su Samba
# -----------------------
df=pd.DataFrame(dati_excel)
anno_corrente=datetime.now().strftime("%Y")
excel_dir=os.path.join(share_output, anno_corrente).replace("\\","/")
smb_mkdirs(tree_out, excel_dir)
excel_path=os.path.join(excel_dir, excel_file).replace("\\","/")

excel_bytes=BytesIO()
df.to_excel(excel_bytes,index=False)
excel_bytes.seek(0)
f_out=Open(tree_out, excel_path, access=0x12019f)
f_out.create(disposition=CreateDisposition.FILE_OVERWRITE_IF,
             options=CreateOptions.FILE_NON_DIRECTORY_FILE)
f_out.write(excel_bytes.read(),0)
f_out.close()

# -----------------------
# 9. Log finale
# -----------------------
log_write(f"Totale file analizzati: {len(file_list)}")
log_write(f"File spostati correttamente: {count_ok}")
log_write(f"File con errore (non spostati): {count_err}")
with open(log_file,"w",encoding="utf-8") as f:
    f.write("\n".join(log_lines))

# -----------------------
# 10. Invio mail testo
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
