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
    Drive (siembra). Las tablas de movimientos (ventas, deudores) no se
    reemplazan jamás: se FUSIONAN (unión sin duplicados por ID) y el resultado
    queda en ambos lados, así ningún dispositivo pisa al otro.
  - Tras cada operación se registra el estado (mtimes de ambos lados) para
    no re-transferir lo ya igualado y resistir desfases de reloj.

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
from datetime import datetime, timezone
from pathlib import Path

import pandas as pd
import streamlit as st

BASE = Path(__file__).resolve().parent
DATA_DIR = BASE / "data"
IDS_PATH = DATA_DIR / ".drive_ids.json"  # caché fileId + estado de sync (ignorado por git vía data/*)

ARCHIVOS_SYNC = ["inventario.xlsx", "ventas.xlsx", "deudores.xlsx",
                 "usuarios.xlsx", "categorias.txt", "categorias_ocultas.txt"]
# Tablas de movimientos con fusión (unión sin duplicados) en vez de reemplazo.
# El resto (inventario, usuarios, categorías) mantiene reemplazo con prioridad-nube.
MERGEABLE = {"ventas.xlsx", "deudores.xlsx"}
HOJA_POR_ARCHIVO = {"ventas.xlsx": "Ventas", "deudores.xlsx": "Deudores"}
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


def _norm_val(v) -> str:
    """Normaliza un valor de celda para comparar filas (63000 == 63000.0 == '63000')."""
    try:
        if pd.isna(v):
            return ""
    except (TypeError, ValueError):
        pass
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    if hasattr(v, "isoformat") and not isinstance(v, str):
        try:
            return str(v.isoformat(sep=" "))  # type: ignore[union-attr]
        except Exception:
            pass
    return str(v).strip()


def _leer_tabla(path: Path) -> pd.DataFrame:
    try:
        return pd.read_excel(path, sheet_name=0)
    except Exception:
        return pd.DataFrame()


def _leer_tabla_bytes(raw: bytes) -> pd.DataFrame:
    try:
        return pd.read_excel(io.BytesIO(raw), sheet_name=0)
    except Exception:
        return pd.DataFrame()


def _claves_filas(df: pd.DataFrame, cols: list) -> pd.Series:
    return pd.Series([tuple(_norm_val(v) for v in fila)
                      for fila in df.reindex(columns=cols).itertuples(index=False)],
                     index=df.index)


def _mismo_contenido(a: pd.DataFrame, b: pd.DataFrame) -> bool:
    cols = list(a.columns) + [c for c in b.columns if c not in a.columns]
    try:
        return _claves_filas(a, cols).sort_values().tolist() == \
               _claves_filas(b, cols).sort_values().tolist()
    except Exception:
        return False


def _fusionar_deudores(todo: pd.DataFrame) -> pd.DataFrame:
    """Resuelve filas con el mismo ID: gana la de mayor Monto_Pagado (abono más
    nuevo); en empate, la de Concepto más largo (más productos agregados);
    en empate, la remota (prioridad-nube). Se elige la fila completa para no
    romper la invariante Total - Pagado = Saldo."""
    def _elige(grupo: pd.DataFrame) -> pd.DataFrame:
        if len(grupo) == 1:
            return grupo.iloc[[0]]
        g = grupo.copy()
        g["_pag"] = pd.to_numeric(g.get("Monto_Pagado"), errors="coerce").fillna(0)
        g["_len"] = g.get("Concepto").fillna("").astype(str).str.len() \
            if "Concepto" in g.columns else 0
        g["_rem"] = (g.get("_origen") == "remoto").astype(int) \
            if "_origen" in g.columns else 0
        idx = g.sort_values(["_pag", "_len", "_rem"],
                            ascending=[False, False, False]).index[0]
        return grupo.loc[[idx]]

    id_col = "ID" if "ID" in todo.columns else todo.columns[0]
    try:
        out = pd.concat([_elige(g) for _, g in todo.groupby(id_col, dropna=False)],
                        ignore_index=True)
    except Exception:
        out = todo
    try:
        out = out.sort_values(id_col).reset_index(drop=True)
    except Exception:
        pass
    return out


def fusionar_tablas(local_df: pd.DataFrame, remoto_df: pd.DataFrame, nombre: str) -> pd.DataFrame:
    """Une dos versiones de una tabla sin perder ni duplicar registros."""
    cols = list(local_df.columns) + [c for c in remoto_df.columns if c not in local_df.columns]
    if not cols:
        return local_df.copy()
    loc = local_df.reindex(columns=cols).copy()
    rem = remoto_df.reindex(columns=cols).copy()
    loc["_origen"] = "local"
    rem["_origen"] = "remoto"
    todo = pd.concat([loc, rem], ignore_index=True)
    todo["_clave"] = _claves_filas(todo, cols)
    todo = todo.drop_duplicates(subset=["_clave"], keep="first")
    if nombre == "deudores.xlsx":
        todo = _fusionar_deudores(todo)
    elif nombre == "ventas.xlsx":
        try:
            todo["_f"] = pd.to_datetime(todo["Fecha"], errors="coerce")
            todo = todo.sort_values(["_f", "ID_Venta"]).reset_index(drop=True)
        except Exception:
            pass
    return todo.drop(columns=[c for c in ("_clave", "_origen", "_f") if c in todo.columns])


def _descargar_bytes(service, fid: str) -> bytes:
    from googleapiclient.http import MediaIoBaseDownload

    req = service.files().get_media(fileId=fid, supportsAllDrives=True)
    buf = io.BytesIO()
    downloader = MediaIoBaseDownload(buf, req)
    done = False
    while not done:
        _, done = downloader.next_chunk()
    return buf.getvalue()


def _escribir_tabla(df: pd.DataFrame, path: Path, hoja: str) -> None:
    DATA_DIR.mkdir(exist_ok=True)
    with pd.ExcelWriter(path, engine="openpyxl") as w:
        df.to_excel(w, sheet_name=hoja, index=False)


def fusionar_archivo(service, folder: str, nombre: str) -> str:
    """Fusiona la versión local con la de Drive (unión sin duplicados) y deja
    el resultado en AMBOS lados. Devuelve "fusionado" o "igual"."""
    ids = _ids()
    fid = ids.get(nombre) or _id_en_drive(service, folder, nombre)
    if not fid:
        raise RuntimeError(f"{nombre} no existe en Drive")
    remoto_bytes = _descargar_bytes(service, fid)
    local_df = _leer_tabla(DATA_DIR / nombre)
    remoto_df = _leer_tabla_bytes(remoto_bytes)
    fusionado = fusionar_tablas(local_df, remoto_df, nombre)
    if _mismo_contenido(fusionado, local_df) and _mismo_contenido(fusionado, remoto_df):
        _marcar(service, folder, nombre)
        return "igual"
    _respaldar_local(nombre)
    _escribir_tabla(fusionado, DATA_DIR / nombre, HOJA_POR_ARCHIVO[nombre])
    subir(nombre)
    _marcar(service, folder, nombre)
    return "fusionado"


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
    if nombre in MERGEABLE:
        # Tablas de movimientos: jamás se reemplazan; se fusionan (unión sin
        # duplicados) para que ningún dispositivo pise al otro.
        if rec is not None:
            try:
                cambio_local = abs(lts - float(rec.get("l", lts))) > TOLERANCIA_SEG
                if not cambio_local and rec.get("h"):
                    cambio_local = _hash_local(nombre) != rec.get("h")
                cambio_remoto = abs(rts - float(rec.get("r", rts))) > TOLERANCIA_SEG
            except (TypeError, ValueError):
                cambio_local = cambio_remoto = True
            if not cambio_local and not cambio_remoto:
                return "igual"
        return fusionar_archivo(service, folder, nombre)
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
        for nombre in ARCHIVOS_SYNC:
            try:
                _sincronizar_uno(service, cfg["folder"], nombre)
            except Exception:
                continue
        return "nube"
    except Exception as e:  # noqa: BLE001 — la app debe abrir igual
        st.warning(f"⚠️ No se pudo sincronizar con Drive ({e}). Se usan datos locales.")
        return "local"


def sincronizacion_manual() -> tuple[bool, str]:
    """Sincronización manual del botón "🔄 Sincronizar con Drive". Nunca lanza.

    Dirección prioritaria: DESCARGA (la nube es la versión oficial), con FUSIÓN
    en tablas de movimientos (ventas, deudores): en vez de reemplazar, se
    concatenan ambas versiones por ID/fila para no duplicar ni perder registros.
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
        subidos, bajados, iguales, conflictos, fusionados, errores = [], [], [], [], [], []
        for nombre in ARCHIVOS_SYNC:
            try:
                r = _sincronizar_uno(service, cfg["folder"], nombre)
            except Exception as e:  # noqa: BLE001 — un archivo no frena los demás
                errores.append(f"{nombre} ({e})")
                continue
            {"subido": subidos, "descargado": bajados, "igual": iguales,
             "conflicto": conflictos, "fusionado": fusionados, "omitido": []}[r].append(nombre)
        if not subidos and not bajados and not conflictos and not fusionados:
            base = ("Conectado a Drive: todo ya estaba actualizado con la nube oficial."
                    if not errores else "Sin cambios aplicados.")
            if errores:
                base += f" Avisos: {'; '.join(errores)}."
            return True, base
        partes = ["☁️ Sincronización con prioridad a la nube."]
        if fusionados:
            partes.append(f"Fusionados sin perder datos ({len(fusionados)}): {', '.join(fusionados)}.")
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
