import os
import configparser
import uuid
import sys
from smbprotocol.connection import Connection
from smbprotocol.session import Session
from smbprotocol.tree import TreeConnect
from smbprotocol.open import Open

# Percorso del file di configurazione nella cartella superiore
base_dir = os.path.dirname(os.path.abspath(__file__))
config_path = os.path.join(base_dir, "..", "config.ini")

config = configparser.ConfigParser()
config.read(config_path)

# Normalizza server: rimuove schema smb:// e parte ._smb._tcp
def normalize_server(s):
    if not s:
        return s
    s = s.strip()
    if s.lower().startswith("smb://"):
        s = s.split("://", 1)[1]
    s = s.replace("._smb._tcp", "")
    s = s.lstrip("/\\").strip(".")
    return s

server = normalize_server(config["Samba"]["server"])
share_in = config["Samba"]["share_input"]
user = config["Samba"]["username"]
pwd = config["Samba"]["password"]

# client_guid deve essere un uuid.UUID
client_guid = uuid.UUID(config["Samba"].get("client_guid", "123e4567-e89b-12d3-a456-426614174000"))

def normalize_share(s):
    s = s.strip()
    if not s:
        return s
    if s.startswith("/") or s.startswith("\\") or ":" in s:
        return os.path.basename(s.rstrip("/\\"))
    return s

share_in_normalized = normalize_share(share_in)

try:
    conn = Connection(guid=client_guid, server_name=server, port=445)
    conn.connect()

    sess = Session(conn, username=user, password=pwd)
    sess.connect()

    tree_path = fr"\\{server}\{share_in_normalized}"
    print(f"DEBUG: connecting to tree {tree_path} with GUID={client_guid}")
    sys.stdout.flush()
    tree = TreeConnect(sess, tree_path)
    tree.connect()

    directory = Open(tree, "")

    # Valori standard SMB2 (numerici) — usati direttamente per compatibilità:
    impersonation_level = 2
    desired_access = 0x120089
    file_attributes = 0x10      # attributo directory
    share_access = 0x07
    create_disposition = 1      # FILE_OPEN
    create_options = 0x00000001 # FILE_DIRECTORY_FILE

    directory.create(
        impersonation_level,
        desired_access,
        file_attributes,
        share_access,
        create_disposition,
        create_options,
    )

    print("Connessione OK. File presenti:")
   
    directory.close()
    tree.disconnect()
    sess.disconnect()
    conn.disconnect()

except Exception as e:
    print("ERRORE:", e)
    print(f"DEBUG INFO: server={server} share_in='{share_in}' normalized='{share_in_normalized}' GUID={client_guid}")
