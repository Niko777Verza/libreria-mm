"""Sincronización persistente con Google Drive — Librería MM.

Estrategia híbrida (local-first):
- **En PC (sin secrets):** todas las funciones son no-op. Los datos viven en
  `data/` como siempre. No requiere instalar nada extra.
- **En Streamlit Cloud (con secrets):** al arrancar descarga los archivos
  (`inventario.xlsx`, `ventas.xlsx`, `deudores.xlsx`, `usuarios.xlsx`,
  `categorias.txt`, `categorias_ocultas.txt`) desde una carpeta de Drive;
  en cada guardado crítico sube el archivo correspondiente.

Autenticación: Service Account leída de `st.secrets["gcp_service_account"]`
más el ID de carpeta en `st.secrets["GDRIVE_FOLDER_ID"]`.
Los imports de Google son perezosos para no exigir las librerías en local.
"""
from __future__ import annotations

import io
import json
from pathlib import Path

import streamlit as st

BASE = Path(__file__).resolve().parent
DATA_DIR = BASE / "data"
IDS_PATH = DATA_DIR / ".drive_ids.json"  # caché fileId (local; ignorado por git vía data/*)

ARCHIVOS_SYNC = ["inventario.xlsx", "ventas.xlsx", "deudores.xlsx",
                 "usuarios.xlsx", "categorias.txt", "categorias_ocultas.txt"]
MIMETYPES = {".xlsx": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
             ".txt": "text/plain"}
SCOPES = ["https://www.googleapis.com/auth/drive"]


def _leer_secrets() -> dict | None:
    """Devuelve {'sa': {...}, 'folder': '...'} o None si no hay configuración."""
    try:
        sa = st.secrets.get("gcp_service_account")
        folder = st.secrets.get("GDRIVE_FOLDER_ID", "")
    except Exception:
        return None
    if not sa or not folder:
        return None
    try:
        return {"sa": dict(sa), "folder": str(folder).strip()}
    except (TypeError, ValueError):
        return None


def modo_nube() -> bool:
    """True si hay credenciales de Drive configuradas (Streamlit Cloud)."""
    return _leer_secrets() is not None


def estado_almacenamiento() -> str:
    return "☁️ Drive conectado" if modo_nube() else "💻 Datos locales"


@st.cache_resource
def _servicio():
    from google.oauth2 import service_account
    from googleapiclient.discovery import build

    cfg = _leer_secrets()
    if not cfg:
        raise RuntimeError("Sin credenciales de Google Drive en st.secrets.")
    creds = service_account.Credentials.from_service_account_info(cfg["sa"], scopes=SCOPES)
    return build("drive", "v3", credentials=creds, cache_discovery=False)


def _ids() -> dict:
    try:
        return json.loads(IDS_PATH.read_text(encoding="utf-8"))
    except (OSError, ValueError):
        return {}


def _guardar_ids(d: dict) -> None:
    try:
        DATA_DIR.mkdir(exist_ok=True)
        IDS_PATH.write_text(json.dumps(d), encoding="utf-8")
    except OSError:
        pass


def _mime(nombre: str) -> str:
    return MIMETYPES.get(Path(nombre).suffix.lower(), "application/octet-stream")


def _id_en_drive(service, folder: str, nombre: str) -> str | None:
    resp = service.files().list(
        q=f"'{folder}' in parents and name = '{nombre}' and trashed = false",
        fields="files(id, name)", pageSize=5, supportsAllDrives=True,
        includeItemsFromAllDrives=True).execute()
    files = resp.get("files", [])
    return files[0]["id"] if files else None


def subir(nombre: str) -> bool:
    """Sube (crea o sobrescribe) un archivo de data/ a la carpeta de Drive."""
    cfg = _leer_secrets()
    if not cfg:
        return False
    local = DATA_DIR / nombre
    if not local.exists():
        return False
    from googleapiclient.http import MediaFileUpload

    service = _servicio()
    ids = _ids()
    fid = ids.get(nombre) or _id_en_drive(service, cfg["folder"], nombre)
    media = MediaFileUpload(str(local), mimetype=_mime(nombre), resumable=False)
    if fid:
        service.files().update(fileId=fid, media_body=media,
                               supportsAllDrives=True).execute()
    else:
        creado = service.files().create(
            body={"name": nombre, "parents": [cfg["folder"]]},
            media_body=media, fields="id",
            supportsAllDrives=True).execute()
        fid = creado["id"]
    ids[nombre] = fid
    _guardar_ids(ids)
    return True


def descargar(nombre: str) -> bool:
    """Descarga un archivo de Drive a data/ local. True si existía en Drive."""
    cfg = _leer_secrets()
    if not cfg:
        return False
    from googleapiclient.http import MediaIoBaseDownload

    service = _servicio()
    ids = _ids()
    fid = ids.get(nombre) or _id_en_drive(service, cfg["folder"], nombre)
    if not fid:
        return False
    req = service.files().get_media(fileId=fid, supportsAllDrives=True)
    buf = io.BytesIO()
    downloader = MediaIoBaseDownload(buf, req)
    done = False
    while not done:
        _, done = downloader.next_chunk()
    DATA_DIR.mkdir(exist_ok=True)
    (DATA_DIR / nombre).write_bytes(buf.getvalue())
    ids[nombre] = fid
    _guardar_ids(ids)
    return True


def sincronizar_inicio() -> str:
    """Punto de entrada al arrancar la app. Devuelve 'nube' o 'local'.

    Regla: si el archivo existe en Drive se descarga (Drive manda);
    si solo existe local, se sube (siembra inicial en la nube).
    Nunca interrumpe la app: ante cualquier error retorna 'local'.
    """
    if not modo_nube():
        return "local"
    try:
        for nombre in ARCHIVOS_SYNC:
            if not descargar(nombre) and (DATA_DIR / nombre).exists():
                subir(nombre)
        return "nube"
    except Exception as e:  # noqa: BLE001 — la app debe abrir igual
        st.warning(f"⚠️ No se pudo sincronizar con Drive ({e}). Se usan datos locales.")
        return "local"


def subir_si_nube(nombre: str) -> bool:
    """Wrapper seguro para llamar tras cada guardado crítico. Nunca lanza."""
    if not modo_nube():
        return False
    try:
        return subir(nombre)
    except Exception as e:  # noqa: BLE001
        st.warning(f"⚠️ Guardado local OK, pero falló la subida a Drive ({nombre}): {e}")
        return False
