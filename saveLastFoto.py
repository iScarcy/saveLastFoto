import os
import configparser
import logging
from datetime import datetime
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import openpyxl
import smbclient

# ------------------------------------------------------------
# Caricamento configurazione
# ------------------------------------------------------------
CONFIG_FILE = "config.ini"
config = configparser.ConfigParser()
config.read(CONFIG_FILE)

server = config["Samba"]["server"]
share_input = config["Samba"]["share_input"]
share_output = config["Samba"]["share_output"]
username = config["Samba"]["username"]
password = config["Samba"]["password"]

send_mail = config["Email"].getboolean("sendMail", False)
smtp_server = config["Email"]["smtp_server"]
smtp_port = int(config["Email"]["smtp_port"])
smtp_user = config["Email"]["smtp_user"]
smtp_password = config["Email"]["smtp_password"]
email_to = config["Email"]["email_to"]

move_file = config["General"].getboolean("moveFile", True)

# ------------------------------------------------------------
# Imposta credenziali SMB
# ------------------------------------------------------------
smbclient.register_session(server, username=username, password=password)

# ------------------------------------------------------------
# Funzione per ottenere data/ora formattata
# ------------------------------------------------------------
def now_string():
    return datetime.now().strftime("%d/%m/%Y %H:%M:%S")

# ------------------------------------------------------------
# Funzione invio email
# ------------------------------------------------------------
def send_email(subject, body):
    msg = MIMEMultipart()
    msg["From"] = smtp_user
    msg["To"] = email_to
    msg["Subject"] = subject
    msg.attach(MIMEText(body, "plain"))

    with smtplib.SMTP(smtp_server, smtp_port) as server:
        server.starttls()
        server.login(smtp_user, smtp_password)
        server.send_message(msg)

# ------------------------------------------------------------
# Funzione principale
# ------------------------------------------------------------
def main():
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    excel_filename = f"output_{timestamp}.xlsx"
    log_filename = f"output_{timestamp}_log.txt"

    input_path = f"//{server}/{share_input}"
    output_path = f"//{server}/{share_output}"

    total_files = 0
    moved_success = 0
    moved_failed = 0

    logging.basicConfig(filename=log_filename, level=logging.INFO, format="%(message)s")

    logging.info(f"Esecuzione script: {now_string()}")
    logging.info("Configurazione usata:")
    for section in config.sections():
        logging.info(f"[{section}]")
        for key, value in config[section].items():
            logging.info(f"{key} = {value}")
        logging.info("")

    try:
        file_list = smbclient.listdir(input_path)
        total_files = len(file_list)

        wb = openpyxl.Workbook()
        ws = wb.active
        ws.append(["Nome File", "Risultato"])

        for filename in file_list:
            src = f"{input_path}/{filename}"
            dst = f"{output_path}/{filename}"

            if move_file:
                try:
                    smbclient.rename(src, dst)
                    ws.append([filename, "OK"])
                    moved_success += 1
                except Exception:
                    ws.append([filename, "ERRORE"])
                    moved_failed += 1
            else:
                ws.append([filename, "NON SPOSTATO (TEST MODE)"])

        wb.save(excel_filename)

        logging.info(f"Totale file analizzati: {total_files}")
        logging.info(f"File spostati correttamente: {moved_success}")
        logging.info(f"File non spostati: {moved_failed}")
        logging.info(f"Fine esecuzione: {now_string()}")

        if send_mail:
            email_body = (
                f"Esecuzione script del {now_string()}\n\n"
                f"Totale file analizzati: {total_files}\n"
                f"File spostati correttamente: {moved_success}\n"
                f"File non spostati: {moved_failed}\n"
                f"Modalità spostamento file: {'ATTIVA' if move_file else 'TEST MODE (nessun file spostato)'}\n"
            )
            send_email("Risultato elaborazione", email_body)

    except Exception as e:
        logging.error("\n*** ERRORE GENERALE ***")
        logging.error(str(e))
        logging.error(f"Errore avvenuto il: {now_string()}")
        print("Errore generale. Controlla il log.")


if __name__ == "__main__":
    main()
