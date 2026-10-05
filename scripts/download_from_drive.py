#!/usr/bin/env python3
"""
download_from_drive.py
======================

Descarga los ejercicios diarios desde una carpeta de Google Drive
hacia `data/ejercicios/`. Usa un Service Account (credenciales JSON
inyectadas como GitHub Secret `GDRIVE_SA_KEY`).

Env vars requeridas:
  GDRIVE_SA_KEY      JSON completo del Service Account
  GDRIVE_FOLDER_ID   ID de la carpeta de Drive con los ejercicios
"""

import os
import json
import io
from pathlib import Path

from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload

SCOPES = ["https://www.googleapis.com/auth/drive.readonly"]
DATA_DIR = Path(__file__).resolve().parent.parent / "data"
TARGET_DIR = DATA_DIR / "ejercicios"
# Pedidos de compra abiertos (fichero diario de Donet) -> data/pedidos/
PEDIDOS_DIR = DATA_DIR / "pedidos"


def main() -> int:
    sa_key_raw = os.environ.get("GDRIVE_SA_KEY")
    folder_id = os.environ.get("GDRIVE_FOLDER_ID")
    if not sa_key_raw or not folder_id:
        print("ERROR: GDRIVE_SA_KEY o GDRIVE_FOLDER_ID no configurados")
        return 1

    sa_info = json.loads(sa_key_raw)
    creds = service_account.Credentials.from_service_account_info(sa_info, scopes=SCOPES)
    svc = build("drive", "v3", credentials=creds, cache_discovery=False)

    TARGET_DIR.mkdir(parents=True, exist_ok=True)

    # Listar todos los ejercicio*.xlsx de la carpeta
    q = f"'{folder_id}' in parents and trashed = false and name contains 'ejercicio' and mimeType = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'"
    resp = svc.files().list(q=q, pageSize=200, fields="files(id, name, modifiedTime)").execute()
    files = resp.get("files", [])
    print(f"[drive] Encontrados {len(files)} ficheros en la carpeta")

    new_count = 0
    for f in files:
        fname = f["name"]
        local = TARGET_DIR / fname
        # Skip si ya existe y no ha cambiado (comparación simple por size)
        if local.exists():
            print(f"  ya existe: {fname}")
            continue
        print(f"  descargando: {fname}")
        request = svc.files().get_media(fileId=f["id"])
        buf = io.BytesIO()
        downloader = MediaIoBaseDownload(buf, request)
        done = False
        while not done:
            _, done = downloader.next_chunk()
        with open(local, "wb") as fp:
            fp.write(buf.getvalue())
        new_count += 1

    print(f"[drive] {new_count} ficheros nuevos descargados a {TARGET_DIR}")

    # --- Pedidos de compra abiertos (mismo folder de Drive) ---
    PEDIDOS_DIR.mkdir(parents=True, exist_ok=True)
    q2 = (f"'{folder_id}' in parents and trashed = false and "
          f"name contains 'pedidos de compra abiertos' and "
          f"mimeType = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'")
    resp2 = svc.files().list(q=q2, pageSize=200, fields="files(id, name, modifiedTime)").execute()
    files2 = resp2.get("files", [])
    print(f"[drive] Encontrados {len(files2)} ficheros de pedidos abiertos")
    new2 = 0
    for f in files2:
        local = PEDIDOS_DIR / f["name"]
        if local.exists():
            print(f"  ya existe: {f['name']}")
            continue
        print(f"  descargando: {f['name']}")
        request = svc.files().get_media(fileId=f["id"])
        buf = io.BytesIO()
        downloader = MediaIoBaseDownload(buf, request)
        done = False
        while not done:
            _, done = downloader.next_chunk()
        with open(local, "wb") as fp:
            fp.write(buf.getvalue())
        new2 += 1
    print(f"[drive] {new2} ficheros de pedidos nuevos descargados a {PEDIDOS_DIR}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
