"""Sincronización persistente con Google Drive — Librería MM.

Estrategia híbrida (local-first, upload-first):
- **En PC (sin secrets):** todas las funciones son no-op. Los datos viven en
  `data/` como siempre. No requiere instalar nada extra.
- **En Streamlit Cloud (con secrets):** las acciones cotidianas guardan solo
  en local (`data/`). El botón "🔄 Sincronizar con Drive" compara marcas de
  tiempo y SUBE primero lo local reciente; solo descarga si Drive es
  claramente más nuevo, siempre tras respaldar en `data/_backup_sync/`.
  En empate gana LO LOCAL para nunca perder datos recientes.

Autenticación: Service Account leída de `st.secrets["gcp_service_account"]`
más el ID de carpeta en `st.secrets["GDRIVE_FOLDER_ID"]`.
Los imports de Google son perezosos para no exigir las librerías en local.
"""
from __future__ import annotations

import io
import json
import shutil
from datetime import datetime, timezone
from pathlib import Path

import streamlit as st

BASE = Path(__file__).resolve().parent
DATA_DIR = BASE / "data"
BACKUP_DIR = DATA_DIR / "_backup_sync"  # respaldos previos a cualquier descarga
IDS_PATH = DATA_DIR / ".drive_ids.json"  # caché fileId (local; ignorado por git vía data/*)

ARCHIVOS_SYNC = ["inventario.xlsx", "ventas.xlsx", "deudores.xlsx",
                 "usuarios.xlsx", "categorias.txt", "categorias_ocultas.txt"]
MIMETYPES = {".xlsx": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
             ".txt": "text/plain"}
SCOPES = ["https://www.googleapis.com/auth/drive"]

# Tolerancia para comparar relojes local vs. Drive. Si la diferencia está
# dentro de este margen, se considera empate y GANA LO LOCAL (upload).
TOLERANCIA_SEG = 30.0


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
    meta = _meta_remota(service, folder, nombre)
    return meta["id"] if meta else None


def _mtime_local(nombre: str) -> float | None:
    """mtime (epoch seg) del archivo local, o None si no existe."""
    try:
        p = DATA_DIR / nombre
        return p.stat().st_mtime if p.exists() else None
    except OSError:
        return None


def _meta_remota(service, folder: str, nombre: str) -> dict | None:
    """Metadatos remotos {id, modifiedTime, ...} o None si no existe en Drive."""
    resp = service.files().list(
        q=f"'{folder}' in parents and name = '{nombre}' and trashed = false",
        fields="files(id, name, modifiedTime, size)", pageSize=5,
        supportsAllDrives=True,
        includeItemsFromAllDrives=True).execute()
    files = resp.get("files", [])
    return files[0] if files else None


def _mtime_remota(meta: dict | None) -> float | None:
    """Convierte Drive modifiedTime (RFC3339) a epoch seg. None si no parseable."""
    if not meta:
        return None
    raw = meta.get("modifiedTime")
    if not raw:
        return None
    try:
        # Formatos: 2026-10-10T12:34:56.123Z  |  ...+00:00
        dt = datetime.fromisoformat(str(raw).replace("Z", "+00:00"))
        if dt.tzinfo is None:
            dt = dt.replace(tzinfo=timezone.utc)
        return dt.timestamp()
    except (ValueError, TypeError):
        return None


def _respaldar_local(nombre: str) -> str | None:
    """Copia de seguridad del archivo local antes de sobrescribirlo.

    Devuelve el nombre del respaldo o None si no había nada que respaldar.
    Nunca lanza excepciones.
    """
    try:
        origen = DATA_DIR / nombre
        if not origen.exists():
            return None
        BACKUP_DIR.mkdir(parents=True, exist_ok=True)
        sello = datetime.now().strftime("%Y%m%d_%H%M%S")
        destino = BACKUP_DIR / f"{nombre}.{sello}.bak"
        shutil.copy2(origen, destino)
        return destino.name
    except OSError:
        return None


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

    Regla local-first con comparación por timestamps (nunca pierde datos):
    - Si solo existe de un lado, se copia hacia el otro (siembra).
    - Si existe en ambos, gana el más reciente; en empate gana LO LOCAL.
    - Solo se descarga cuando Drive es claramente más nuevo (> tolerancia),
      y siempre tras respaldar el archivo local en data/_backup_sync/.
    Nunca interrumpe la app: ante cualquier error retorna 'local'.
    """
    if not modo_nube():
        return "local"
    try:
        cfg = _leer_secrets()
        service = _servicio()
        for nombre in ARCHIVOS_SYNC:
            try:
                local_mtime = _mtime_local(nombre)
                meta = _meta_remota(service, cfg["folder"], nombre)  # type: ignore[union-attr]
                remote_mtime = _mtime_remota(meta)
                if local_mtime is None and meta is None:
                    continue
                if local_mtime is not None and meta is None:
                    subir(nombre)
                elif local_mtime is None and meta is not None:
                    descargar(nombre)
                elif local_mtime is not None and remote_mtime is not None:
                    if local_mtime >= remote_mtime - TOLERANCIA_SEG:
                        subir(nombre)  # local más nuevo o empate -> nube se actualiza
                    else:
                        _respaldar_local(nombre)
                        descargar(nombre)
                elif local_mtime is not None:
                    # Sin timestamp remoto confiable: no se arriesga descarga.
                    subir(nombre)
            except Exception:
                continue
        return "nube"
    except Exception as e:  # noqa: BLE001 — la app debe abrir igual
        st.warning(f"⚠️ No se pudo sincronizar con Drive ({e}). Se usan datos locales.")
        return "local"


def sincronizacion_manual() -> tuple[bool, str]:
    """Sincronización manual del botón "🔄 Sincronizar con Drive". Nunca lanza.

    Dirección prioritaria: SUBIR (upload-first, local-first).
    1) Por cada archivo compara mtime local vs. modifiedTime de Drive.
    2) Si lo local es más reciente o igual (tolerancia incluida) -> SUBE.
       La nube se actualiza con lo último que hizo el usuario. Nunca se
       descarga encima de un archivo local reciente.
    3) Solo descarga cuando Drive es claramente más nuevo, y SIEMPRE tras
       crear un respaldo en data/_backup_sync/ para no perder nada.
    4) Si no hay remoto, siembra (sube); si no hay local, descarga (nada que perder).
    Devuelve (ok, mensaje_para_mostrar).
    """
    if not modo_nube():
        return False, "Sin conexión a Drive (modo local). Configura los secrets para sincronizar."
    try:
        cfg = _leer_secrets()
        assert cfg is not None
        service = _servicio()
        subidos, bajados, respaldos, iguales, errores = [], [], [], [], []
        for nombre in ARCHIVOS_SYNC:
            try:
                local_mtime = _mtime_local(nombre)
                try:
                    meta = _meta_remota(service, cfg["folder"], nombre)
                except Exception as e:  # error de red/API en este archivo
                    errores.append(f"{nombre} (Drive: {e})")
                    continue
                remote_mtime = _mtime_remota(meta)

                if local_mtime is None and meta is None:
                    continue  # no existe en ningún lado
                if local_mtime is not None and meta is None:
                    # Siembra inicial: solo existe local -> subir, jamás borrar.
                    if subir(nombre):
                        subidos.append(nombre)
                    continue
                if local_mtime is None and meta is not None:
                    # Solo existe en la nube y nada local que perder -> descargar.
                    if descargar(nombre):
                        bajados.append(nombre)
                    continue
                # Existe en ambos: decide por timestamps, empate -> LOCAL.
                if remote_mtime is None:
                    # Sin fecha remota confiable: proteger lo local, solo subir.
                    if subir(nombre):
                        subidos.append(nombre)
                    continue
                assert local_mtime is not None
                if local_mtime >= remote_mtime - TOLERANCIA_SEG:
                    if subir(nombre):
                        subidos.append(nombre)
                    else:
                        iguales.append(nombre)
                else:
                    # Drive claramente más nuevo: respaldar antes de tocar lo local.
                    bak = _respaldar_local(nombre)
                    if bak:
                        respaldos.append(f"{nombre}→{bak}")
                    if descargar(nombre):
                        bajados.append(nombre)
            except Exception as e:  # noqa: BLE001 — un archivo no frena los demás
                errores.append(f"{nombre} ({e})")
                continue
        if not subidos and not bajados:
            base = "Conectado a Drive: todo ya estaba actualizado (gana lo más reciente)."
            if errores:
                base += f" Avisos: {'; '.join(errores)}."
            return True, base
        partes = [f"☁️ Sincronización protegida (local primero)."]
        if subidos:
            partes.append(f"Subidos a Drive ({len(subidos)}): {', '.join(subidos)}.")
        if bajados:
            partes.append(f"Descargados de Drive ({len(bajados)}): {', '.join(bajados)}.")
        if respaldos:
            partes.append(f"Respaldos locales previos en data/_backup_sync/ ({len(respaldos)}).")
        if errores:
            partes.append(f"Avisos: {'; '.join(errores)}.")
        return True, " ".join(partes)
    except Exception as e:  # noqa: BLE001
        return False, f"No se pudo sincronizar con Drive: {e}"
