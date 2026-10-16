"""Sincronización persistente con Google Drive — Librería MM.

Estrategia con PRIORIDAD ABSOLUTA A LA NUBE (cloud-wins):
- **En PC (sin secrets):** todas las funciones son no-op. Los datos viven en
  `data/` como siempre. No requiere instalar nada extra.
- **Con secrets (nube configurada):** al arrancar la app y al pulsar
  "🔄 Sincronizar con Drive" se compara cada archivo local contra su
  `modifiedTime` de Drive. La nube es la versión oficial:
  - Si Drive es más reciente (o hay empate/conflicto) -> se DESCARGA,
    siempre tras respaldar lo local en `data/_backup_sync/`.
  - Solo se SUBE cuando lo local es claramente más reciente (venta o edición
    posterior a la última sincronización) o cuando el archivo no existe en
    Drive (siembra inicial).
  - Tras cada operación se registra el estado (mtimes de ambos lados) para
    no re-transferir lo ya igualado y resistir desfases de reloj.
- **Respaldo post-cobro:** al confirmar un ticket, la app guarda en local y
  lanza `respaldo_venta_fondo()`: solo `ventas.xlsx` + `inventario.xlsx` suben
  en un hilo daemon (sin bloquear la UI) y dejan constancia en
  `data/.respaldo_estado.json` para mostrar su resultado.

Autenticación: Service Account leída de `st.secrets["gcp_service_account"]`
más el ID de carpeta en `st.secrets["GDRIVE_FOLDER_ID"]`.
Los imports de Google son perezosos para no exigir las librerías en local.
"""
from __future__ import annotations

import io
import hashlib
import json
import os
import shutil
import threading
from datetime import datetime, timezone
from pathlib import Path

import streamlit as st

BASE = Path(__file__).resolve().parent
DATA_DIR = BASE / "data"
IDS_PATH = DATA_DIR / ".drive_ids.json"  # caché fileId + estado de sync (ignorado por git vía data/*)

ARCHIVOS_SYNC = ["inventario.xlsx", "ventas.xlsx", "deudores.xlsx",
                 "usuarios.xlsx", "categorias.txt", "categorias_ocultas.txt",
                 "caja_turnos.xlsx", "caja_movimientos.xlsx", "abonos.xlsx"]
# Subconjunto crítico que se respalda solo tras cada cobro (rápido y selectivo).
ARCHIVOS_VENTA = ["ventas.xlsx", "inventario.xlsx"]
ESTADO_RESPALDO_PATH = DATA_DIR / ".respaldo_estado.json"  # {ok, ts, detalle}
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
    cfg = _leer_secrets()
    if not cfg:
        raise RuntimeError("Sin credenciales de Google Drive en st.secrets.")
    return _servicio_con(cfg["sa"])


def _servicio_con(sa: dict):
    """Construye el cliente de Drive sin tocar st (apto para hilos secundarios)."""
    from google.oauth2 import service_account
    from googleapiclient.discovery import build

    creds = service_account.Credentials.from_service_account_info(sa, scopes=SCOPES)
    return build("drive", "v3", credentials=creds, cache_discovery=False)


def _ids() -> dict:
    return _cargar_store().get("ids", {})


def _guardar_ids(d: dict) -> None:
    store = _cargar_store()
    store["ids"] = d
    _escribir_store(store)


def _cargar_store() -> dict:
    """Lee .drive_ids.json con compatibilidad hacia atrás (antes era {nombre: id})."""
    try:
        d = json.loads(IDS_PATH.read_text(encoding="utf-8"))
    except (OSError, ValueError):
        return {"ids": {}, "sync": {}}
    if isinstance(d, dict) and "ids" in d:
        d.setdefault("sync", {})
        return d
    if isinstance(d, dict):
        return {"ids": d, "sync": {}}
    return {"ids": {}, "sync": {}}


def _escribir_store(store: dict) -> None:
    try:
        DATA_DIR.mkdir(exist_ok=True)
        IDS_PATH.write_text(json.dumps(store), encoding="utf-8")
    except OSError:
        pass


def _estado() -> dict:
    """Mapa {archivo: {l: mtime_local, r: ts_remoto}} de la última sincronización."""
    s = _cargar_store().get("sync", {})
    return s if isinstance(s, dict) else {}


def _guardar_estado(sync: dict) -> None:
    store = _cargar_store()
    store["sync"] = sync
    _escribir_store(store)


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


_CACHE_METAS: dict[str, dict] = {}


def _precargar_metas_remotas(service, folder: str) -> None:
    """Precarga metadatos de todos los archivos en la carpeta de Drive en UNA sola petición HTTP."""
    global _CACHE_METAS
    try:
        resp = service.files().list(
            q=f"'{folder}' in parents and trashed = false",
            fields="files(id, name, modifiedTime, size)", pageSize=100,
            supportsAllDrives=True,
            includeItemsFromAllDrives=True).execute()
        _CACHE_METAS = {f["name"]: f for f in resp.get("files", []) if "name" in f}
    except Exception:
        _CACHE_METAS = {}


def _meta_remota(service, folder: str, nombre: str) -> dict | None:
    """Metadatos remotos {id, modifiedTime, ...} o None si no existe en Drive."""
    global _CACHE_METAS
    if nombre in _CACHE_METAS:
        return _CACHE_METAS[nombre]
    try:
        resp = service.files().list(
            q=f"'{folder}' in parents and name = '{nombre}' and trashed = false",
            fields="files(id, name, modifiedTime, size)", pageSize=5,
            supportsAllDrives=True,
            includeItemsFromAllDrives=True).execute()
        files = resp.get("files", [])
        m = files[0] if files else None
        if m:
            _CACHE_METAS[nombre] = m
        return m
    except Exception:
        return None


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


def _dir_respaldos() -> Path:
    return DATA_DIR / "_backup_sync"


def _respaldar_local(nombre: str) -> str | None:
    """Copia de seguridad del archivo local antes de sobrescribirlo.

    Devuelve el nombre del respaldo o None si no había nada que respaldar.
    Conserva como máximo 20 respaldos por archivo. Nunca lanza excepciones.
    """
    try:
        origen = DATA_DIR / nombre
        if not origen.exists():
            return None
        _dir_respaldos().mkdir(parents=True, exist_ok=True)
        sello = datetime.now().strftime("%Y%m%d_%H%M%S")
        destino = _dir_respaldos() / f"{nombre}.{sello}.bak"
        shutil.copy2(origen, destino)
        _podar_respaldos(nombre)
        return destino.name
    except OSError:
        return None


def _podar_respaldos(nombre: str, mantener: int = 20) -> None:
    try:
        cands = sorted(_dir_respaldos().glob(f"{nombre}.*.bak"))
        for viejo in cands[:-mantener] if len(cands) > mantener else []:
            try:
                viejo.unlink()
            except OSError:
                pass
    except OSError:
        pass


def _subir_con(service, folder: str, nombre: str) -> str | None:
    """Núcleo de subida con un servicio ya construido. Devuelve fileId o None."""
    local = DATA_DIR / nombre
    if not local.exists():
        return None
    from googleapiclient.http import MediaFileUpload

    ids = _ids()
    fid = ids.get(nombre) or _id_en_drive(service, folder, nombre)
    media = MediaFileUpload(str(local), mimetype=_mime(nombre), resumable=False)
    if fid:
        service.files().update(fileId=fid, media_body=media,
                               supportsAllDrives=True).execute()
    else:
        creado = service.files().create(
            body={"name": nombre, "parents": [folder]},
            media_body=media, fields="id",
            supportsAllDrives=True).execute()
        fid = creado["id"]
    ids[nombre] = fid
    _guardar_ids(ids)
    _CACHE_METAS.pop(nombre, None)  # invalidar precarga (el id/ts cambiaron)
    return fid


def subir(nombre: str) -> bool:
    """Sube (crea o sobrescribe) un archivo de data/ a la carpeta de Drive."""
    cfg = _leer_secrets()
    if not cfg:
        return False
    local = DATA_DIR / nombre
    if not local.exists():
        return False
    return _subir_con(_servicio(), cfg["folder"], nombre) is not None


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
    _CACHE_METAS.pop(nombre, None)
    return True


def _hash_local(nombre: str) -> str | None:
    try:
        return hashlib.sha256((DATA_DIR / nombre).read_bytes()).hexdigest()
    except OSError:
        return None


def _marcar(service, folder: str, nombre: str) -> None:
    """Registra el estado post-sync y alinea el mtime local con el remoto.

    Alinear relojes evita re-transferencias y resiste desfases de reloj entre
    la PC y Google. Nunca lanza (es solo optimización).
    """
    try:
        meta = _meta_remota(service, folder, nombre)
        rts = _mtime_remota(meta)
        if rts:
            try:
                os.utime(DATA_DIR / nombre, (rts, rts))
            except OSError:
                pass
        sync = _estado()
        sync[nombre] = {"l": _mtime_local(nombre), "r": rts, "h": _hash_local(nombre)}
        _guardar_estado(sync)
    except Exception:
        pass


def _escribir_estado_respaldo(ok: bool, detalle: str) -> None:
    """Deja constancia del último respaldo post-cobro (la UI la muestra sin bloquear)."""
    try:
        DATA_DIR.mkdir(exist_ok=True)
        ESTADO_RESPALDO_PATH.write_text(
            json.dumps({"ok": ok, "ts": datetime.now(timezone.utc).timestamp(),
                        "detalle": detalle}, ensure_ascii=False),
            encoding="utf-8")
    except OSError:
        pass


def estado_ultimo_respaldo() -> dict | None:
    """Lee la constancia del último respaldo post-cobro, o None si nunca hubo."""
    try:
        d = json.loads(ESTADO_RESPALDO_PATH.read_text(encoding="utf-8"))
        return d if isinstance(d, dict) else None
    except (OSError, ValueError):
        return None


def _respaldar_venta_en_hilo(cfg: dict) -> None:
    """Trabajo del hilo secundario: sube ventas+stock y registra estado. Nunca lanza."""
    try:
        service = _servicio_con(cfg["sa"])
        subidos, errores = [], []
        for nombre in ARCHIVOS_VENTA:
            try:
                if _subir_con(service, cfg["folder"], nombre) is None:
                    errores.append(f"{nombre} (sin archivo local)")
                    continue
                _marcar(service, cfg["folder"], nombre)
                subidos.append(nombre)
            except Exception as e:  # noqa: BLE001 — un archivo no frena al otro
                errores.append(f"{nombre} ({e})")
        if errores and not subidos:
            _escribir_estado_respaldo(False, "; ".join(errores))
        else:
            detalle = f"Subidos: {', '.join(subidos)}" + (
                f". Avisos: {'; '.join(errores)}" if errores else "")
            _escribir_estado_respaldo(True, detalle)
    except Exception as e:  # noqa: BLE001 — p. ej. sin red al momento del cobro
        _escribir_estado_respaldo(False, str(e))


def respaldo_venta_fondo() -> bool:
    """Lanza el respaldo post-cobro (ventas+stock) en un hilo daemon. Nunca lanza.

    No bloquea la interfaz: el cobro ya quedó guardado en local y la subida
    corre en segundo plano con su propio cliente de Drive (sin tocar `st`).
    El botón manual sigue disponible como respaldo. Devuelve True si el hilo
    fue lanzado (hay nube configurada), False si es no-op local.
    """
    try:
        cfg = _leer_secrets()
        if not cfg:
            return False
        hilo = threading.Thread(target=_respaldar_venta_en_hilo, args=(cfg,),
                                name="respaldo-drive-venta", daemon=True)
        hilo.start()
        return True
    except Exception:
        return False


def _sincronizar_uno(service, folder: str, nombre: str) -> str:
    """Sincroniza un archivo con PRIORIDAD A LA NUBE.

    Devuelve: "subido" | "descargado" | "igual" | "omitido" | "conflicto".
    - "conflicto": ambos lados cambiaron -> se respalda y GANA LA NUBE.
    - Sin registro previo (primera vez): gana el más reciente por timestamp
      absoluto; en empate gana la nube. Lo local obsoleto jamás sube.
    """
    lts = _mtime_local(nombre)
    meta = _meta_remota(service, folder, nombre)
    rts = _mtime_remota(meta)

    if meta is None and lts is None:
        return "omitido"
    if meta is None:
        subir(nombre)  # siembra: solo existe local
        _marcar(service, folder, nombre)
        return "subido"
    if lts is None:
        descargar(nombre)  # solo existe en la nube
        _marcar(service, folder, nombre)
        return "descargado"
    if rts is None:
        subir(nombre)  # sin fecha remota confiable: no se puede probar obsolescencia
        _marcar(service, folder, nombre)
        return "subido"

    rec = _estado().get(nombre)
    if rec is None:
        # Bootstrap: comparación absoluta; empate -> nube.
        assert lts is not None
        if lts - rts > TOLERANCIA_SEG:
            subir(nombre)
            _marcar(service, folder, nombre)
            return "subido"
        _respaldar_local(nombre)
        descargar(nombre)
        _marcar(service, folder, nombre)
        return "descargado"

    try:
        # Cambio local: mtime distinto O contenido distinto (el hash cubre
        # ediciones dentro de la ventana de tolerancia del reloj).
        cambio_local = abs(lts - float(rec.get("l", lts))) > TOLERANCIA_SEG
        if not cambio_local and rec.get("h"):
            cambio_local = _hash_local(nombre) != rec.get("h")
        cambio_remoto = abs(rts - float(rec.get("r", rts))) > TOLERANCIA_SEG
    except (TypeError, ValueError):
        cambio_local = cambio_remoto = True

    if not cambio_local and not cambio_remoto:
        return "igual"
    if cambio_remoto and not cambio_local:
        _respaldar_local(nombre)
        descargar(nombre)
        _marcar(service, folder, nombre)
        return "descargado"
    if cambio_local and not cambio_remoto:
        subir(nombre)
        _marcar(service, folder, nombre)
        return "subido"
    # Conflicto real: cambió en ambos lados -> respalda y GANA LA NUBE.
    _respaldar_local(nombre)
    descargar(nombre)
    _marcar(service, folder, nombre)
    return "conflicto"


def sincronizar_inicio() -> str:
    """Punto de entrada al arrancar la app. Devuelve 'nube' o 'local'.

    Aplica prioridad-nube por archivo (ver _sincronizar_uno): al abrir, la
    interfaz ya refleja los datos oficiales de Drive. Nunca interrumpe la app.
    """
    if not modo_nube():
        return "local"
    try:
        cfg = _leer_secrets()
        assert cfg is not None
        service = _servicio()
        _precargar_metas_remotas(service, cfg["folder"])
        for nombre in ARCHIVOS_SYNC:
            try:
                _sincronizar_uno(service, cfg["folder"], nombre)
            except Exception:
                continue
        return "nube"
    except Exception as e:  # noqa: BLE001 — la app debe abrir igual
        st.warning(f"⚠️ No se pudo sincronizar con Drive ({e}). Se usan datos locales.")
        return "local"


def _error_amable(nombre: str, e: Exception) -> str:
    """Resume errores técnicos de Drive en instrucciones accionables."""
    txt = str(e)
    if "storageQuotaExceeded" in txt or "Service Accounts do not have storage quota" in txt:
        return (f"'{nombre}' aún no existe en Drive y la cuenta de servicio no puede crearlo: "
                f"súbelo manualmente una sola vez a la carpeta compartida y luego se actualiza solo")
    if "HttpError 404" in txt:
        return f"'{nombre}' no se encontró en Drive (¿se movió o eliminó?)"
    if "HttpError 403" in txt:
        return f"sin permiso sobre '{nombre}' (revisa que la carpeta esté compartida como Editor)"
    return txt[:300]


def sincronizacion_manual() -> tuple[bool, str]:
    """Sincronización manual del botón "🔄 Sincronizar con Drive". Nunca lanza.

    Dirección prioritaria: DESCARGA (la nube es la versión oficial).
    1) Por cada archivo compara estado local vs. Drive (timestamps + registro).
    2) Drive más reciente, empate o conflicto -> se DESCARGA tras respaldar
       lo local en data/_backup_sync/. Lo local obsoleto jamás sobrescribe.
    3) Solo se SUBE lo local claramente más reciente o si el archivo no
       existe en Drive (siembra).
    4) Lo ya igualado se omite (sin tráfico innecesario).
    La interfaz refleja lo descargado de inmediato (el llamador limpia caché).
    Devuelve (ok, mensaje_para_mostrar).
    """
    if not modo_nube():
        return False, "Sin conexión a Drive (modo local). Configura los secrets para sincronizar."
    try:
        cfg = _leer_secrets()
        assert cfg is not None
        service = _servicio()
        _precargar_metas_remotas(service, cfg["folder"])
        subidos, bajados, iguales, conflictos, errores = [], [], [], [], []
        for nombre in ARCHIVOS_SYNC:
            try:
                r = _sincronizar_uno(service, cfg["folder"], nombre)
            except Exception as e:  # noqa: BLE001 — un archivo no frena los demás
                errores.append(f"{nombre} ({_error_amable(nombre, e)})")
                continue
            {"subido": subidos, "descargado": bajados, "igual": iguales,
             "conflicto": conflictos, "omitido": []}[r].append(nombre)
        if not subidos and not bajados and not conflictos:
            base = ("Conectado a Drive: todo ya estaba actualizado con la nube oficial."
                    if not errores else "Sin cambios aplicados.")
            if errores:
                base += f" Avisos: {'; '.join(errores)}."
            return True, base
        partes = ["☁️ Sincronización con prioridad a la nube."]
        if bajados:
            partes.append(f"Descargados de Drive ({len(bajados)}): {', '.join(bajados)}.")
        if conflictos:
            partes.append(f"Conflictos resueltos a favor de la nube ({len(conflictos)}): "
                          f"{', '.join(conflictos)} (respaldo previo en data/_backup_sync/).")
        if subidos:
            partes.append(f"Subidos a Drive ({len(subidos)}): {', '.join(subidos)}.")
        if iguales:
            partes.append(f"Sin cambios ({len(iguales)}).")
        if errores:
            partes.append(f"Avisos: {'; '.join(errores)}.")
        return True, " ".join(partes)
    except Exception as e:  # noqa: BLE001
        return False, f"No se pudo sincronizar con Drive: {e}"
