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
TARGET_DIR = Path(__file__).resolve().parent.parent / "data" / "ejercicios"


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
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
