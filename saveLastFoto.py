import os
import shutil
import pandas as pd
from PIL import Image, ExifTags
from datetime import datetime
import configparser
import smtplib
import traceback
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders

# -------------------------
#  CARICA CONFIGURAZIONE
# -------------------------
config = configparser.ConfigParser()
config.read("config.ini")

input_folder = config["Local"]["input_folder"]
output_folder = config["Local"]["output_folder"]
move_file = config.getboolean("Local", "moveFile", fallback=True)

send_mail = config.getboolean("Email", "sendMail", fallback=True)
smtp_server = config["Email"]["smtp_server"]
smtp_port = config.getint("Email", "smtp_port")
smtp_user = config["Email"]["smtp_user"]
smtp_pass = config["Email"]["smtp_pass"]
mail_to = config["Email"]["mail_to"]

# -------------------------
# FILE OUTPUT
# -------------------------
excel_file = "log_metadati.xlsx"
log_file = excel_file.replace(".xlsx", "_log.txt")

# -------------------------
# MESI IN ITALIANO
# -------------------------
mesi_italiano = {
    "01": "Gennaio", "02": "Febbraio", "03": "Marzo", "04": "Aprile",
    "05": "Maggio", "06": "Giugno", "07": "Luglio", "08": "Agosto",
    "09": "Settembre", "10": "Ottobre", "11": "Novembre", "12": "Dicembre"
}

# -------------------------
# FUNZIONE ESTRAZIONE EXIF
# -------------------------
def estrai_metadati(file_path):
    try:
        img = Image.open(file_path)
        exif_data = img._getexif()
        if exif_data is None:
            return None, "Nessun EXIF trovato"

        exif = {}
        for tag, value in exif_data.items():
            decoded = ExifTags.TAGS.get(tag, tag)
            exif[decoded] = value

        data_creazione = exif.get("DateTimeOriginal") or exif.get("DateTime")
        return {"data_creazione": data_creazione, **exif}, None
    except Exception as e:
        return None, str(e)


# -------------------------
# FUNZIONE PRINCIPALE
# -------------------------
def main():
    timestamp_esecuzione = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
    dati = []

    total_files = 0
    moved_ok = 0
    moved_ko = 0

    # lista file locali
    file_list = [
        f for f in os.listdir(input_folder)
        if f.lower().endswith((".jpg", ".jpeg", ".png", ".tiff", ".bmp"))
    ]

    total_files = len(file_list)

    # log iniziale
    with open(log_file, "a", encoding="utf-8") as log:
        log.write("\n\n---------------------------------------\n")
        log.write(f"Esecuzione: {timestamp_esecuzione}\n")
        log.write("Configurazioni utilizzate:\n")
        for section in config.sections():
            log.write(f"[{section}]\n")
            for key, value in config[section].items():
                log.write(f"{key} = {value}\n")
        log.write("---------------------------------------\n")

    # elaborazione file
    for file_name in file_list:

        file_path = os.path.join(input_folder, file_name)
        metadati, errore = estrai_metadati(file_path)

        if metadati and metadati.get("data_creazione"):
            try:
                dt = datetime.strptime(metadati["data_creazione"], "%Y:%m:%d %H:%M:%S")
                anno = dt.strftime("%Y")
                mese_num = dt.strftime("%m")
                mese_nome = mesi_italiano[mese_num]
            except:
                anno = mese_num = mese_nome = "unknown"
        else:
            anno = mese_num = mese_nome = "unknown"

        dest_dir = os.path.join(output_folder, anno, f"{mese_num} {mese_nome}")
        dest_path = os.path.join(dest_dir, file_name)

        os.makedirs(dest_dir, exist_ok=True)

        # spostamento o simulazione
        if move_file:
            try:
                shutil.copy2(file_path, dest_path)
                os.remove(file_path)
                esito = "ok"
                moved_ok += 1
            except Exception as e:
                esito = f"ko: {e}"
                moved_ko += 1
        else:
            esito = "simulazione"
            moved_ok += 1

        dati.append({
            "data_esecuzione": timestamp_esecuzione,
            "directory_sorgente": input_folder,
            "file": file_name,
            "metadati": metadati if metadati else "",
            "errore_lettura": errore if errore else "",
            "destinazione": dest_path,
            "esito_spostamento": esito
        })

    # --- Excel ---
    if os.path.exists(excel_file):
        df_old = pd.read_excel(excel_file)
        df_new = pd.DataFrame(dati)
        df_finale = pd.concat([df_old, df_new], ignore_index=True)
    else:
        df_finale = pd.DataFrame(dati)

    df_finale.to_excel(excel_file, index=False)

    # --- log finale ---
    with open(log_file, "a", encoding="utf-8") as log:
        log.write(f"Totale file analizzati: {total_files}\n")
        log.write(f"File spostati con successo: {moved_ok}\n")
        log.write(f"File NON spostati: {moved_ko}\n")

    # --- invio mail ---
    if send_mail:
        subject = "Report elaborazione immagini"
        body = (
            f"Esecuzione: {timestamp_esecuzione}\n"
            f"Totale file analizzati: {total_files}\n"
            f"File spostati: {moved_ok}\n"
            f"File non spostati: {moved_ko}\n"
        )

        msg = MIMEMultipart()
        msg['From'] = smtp_user
        msg['To'] = mail_to
        msg['Subject'] = subject

        part = MIMEBase('application', 'octet-stream')
        with open(excel_file, "rb") as f:
            part.set_payload(f.read())
        encoders.encode_base64(part)
        part.add_header('Content-Disposition', f'attachment; filename="{excel_file}"')
        msg.attach(part)

        body_part = MIMEBase('text', 'plain')
        body_part.set_payload(body)
        msg.attach(body_part)

        with smtplib.SMTP(smtp_server, smtp_port) as server_mail:
            server_mail.starttls()
            server_mail.login(smtp_user, smtp_pass)
            server_mail.sendmail(smtp_user, mail_to, msg.as_string())

    print(f"✔ Elaborazione completata — File elaborati: {total_files}")


# -------------------------
# TRY/EXCEPT GENERALE
# -------------------------
if __name__ == "__main__":
    try:
        main()

    except Exception as e:
        timestamp = datetime.now().strftime("%d/%m/%Y %H:%M:%S")

        errore = (
            f"\n[{timestamp}] ERRORE IMPREVISTO\n"
            f"Tipo: {type(e).__name__}\n"
            f"Dettagli: {str(e)}\n"
            f"Traceback:\n{traceback.format_exc()}\n"
        )

        try:
            with open(log_file, "a", encoding="utf-8") as log:
                log.write(errore)
        except:
            print("ERRORE GRAVE: impossibile scrivere nel log.")

        print("Errore imprevisto. Log aggiornato. Mail NON inviata.")
