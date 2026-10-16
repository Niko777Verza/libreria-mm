"""Módulo de Arqueo/Caja y Deudores Pro — Librería MM.

Dominio puro + almacenamiento (sin UI Streamlit, salvo caché): la
presentación vive en app.py. Cubre:
- Turnos de caja (apertura, movimientos de caja chica, cortes X y Z).
- Historial de abonos con recibo digital.
- Notificación de deudores por WhatsApp (enlace wa.me).

Moneda: Guaraníes enteros (sin decimales).
"""
from __future__ import annotations

import html
import re
from datetime import datetime
from pathlib import Path
from urllib.parse import quote

import pandas as pd
import streamlit as st

BASE = Path(__file__).resolve().parent
DATA_DIR = BASE / "data"
TURNOS_PATH = DATA_DIR / "caja_turnos.xlsx"
MOVS_PATH = DATA_DIR / "caja_movimientos.xlsx"
ABONOS_PATH = DATA_DIR / "abonos.xlsx"

COL_TURNO = ["ID_Turno", "Fecha_Apertura", "Fecha_Cierre", "Cajero_Doc", "Cajero_Nombre",
             "Caja", "Fondo_Inicial", "Estado", "Ventas_Efectivo", "Ventas_Tarjeta",
             "Ventas_Transferencia", "Ventas_Otro", "Entradas", "Retiros",
             "Esperado_Efectivo", "Contado_Efectivo", "Diferencia", "Notas_Cierre"]
COL_MOV = ["ID_Mov", "ID_Turno", "Fecha", "Tipo", "Concepto", "Monto",
           "Cajero_Doc", "Cajero_Nombre"]
COL_ABONO = ["ID_Abono", "ID_Deuda", "Cliente", "Fecha", "Monto",
             "Metodo", "Cajero_Doc", "Cajero_Nombre", "Notas"]

METODOS_ARQUEO = ["Efectivo", "Tarjeta", "Transferencia", "Otro"]


def ahora_str() -> str:
    return datetime.now().strftime("%Y-%m-%d %H:%M")


def _a_entero(x) -> int:
    try:
        return int(round(float(x)))
    except (TypeError, ValueError):
        return 0


def _leer(path: Path, hoja: str, columnas: list) -> pd.DataFrame:
    if not path.exists():
        return pd.DataFrame(columns=columnas)
    try:
        df = pd.read_excel(path, sheet_name=hoja)
        for c in columnas:
            if c not in df.columns:
                df[c] = pd.NA
        return df[list(columnas)].copy()
    except Exception:
        return pd.DataFrame(columns=columnas)


def _escribir(df: pd.DataFrame, path: Path, hoja: str, columnas: list) -> None:
    DATA_DIR.mkdir(exist_ok=True)
    with pd.ExcelWriter(path, engine="openpyxl") as w:
        df[list(columnas)].to_excel(w, sheet_name=hoja, index=False)
        ws = w.sheets[hoja]
        for col in ws.columns:
            ancho = max((len(str(c.value)) if c.value is not None else 0) for c in col) + 2
            ws.column_dimensions[col[0].column_letter].width = min(ancho, 46)


def asegurar_archivos() -> None:
    """Crea los xlsx de caja/abonos vacíos (con cabecera) si no existen."""
    for path, hoja, cols in ((TURNOS_PATH, "Turnos", COL_TURNO),
                             (MOVS_PATH, "Movimientos", COL_MOV),
                             (ABONOS_PATH, "Abonos", COL_ABONO)):
        if not path.exists():
            _escribir(pd.DataFrame(columns=cols), path, hoja, cols)


@st.cache_data(ttl=60)
def cargar_turnos() -> pd.DataFrame:
    df = _leer(TURNOS_PATH, "Turnos", COL_TURNO)
    for c in ("Fondo_Inicial", "Ventas_Efectivo", "Ventas_Tarjeta", "Ventas_Transferencia",
              "Ventas_Otro", "Entradas", "Retiros", "Esperado_Efectivo",
              "Contado_Efectivo", "Diferencia"):
        if c in df.columns:
            df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0).round(0).astype(int)
    return df


@st.cache_data(ttl=60)
def cargar_movimientos() -> pd.DataFrame:
    df = _leer(MOVS_PATH, "Movimientos", COL_MOV)
    if len(df) and "Monto" in df.columns:
        df["Monto"] = pd.to_numeric(df["Monto"], errors="coerce").fillna(0).round(0).astype(int)
    return df


@st.cache_data(ttl=60)
def cargar_abonos() -> pd.DataFrame:
    df = _leer(ABONOS_PATH, "Abonos", COL_ABONO)
    if len(df) and "Monto" in df.columns:
        df["Monto"] = pd.to_numeric(df["Monto"], errors="coerce").fillna(0).round(0).astype(int)
    return df


def guardar_turnos(df: pd.DataFrame) -> None:
    _escribir(df, TURNOS_PATH, "Turnos", COL_TURNO)
    cargar_turnos.clear()


def guardar_movimientos(df: pd.DataFrame) -> None:
    _escribir(df, MOVS_PATH, "Movimientos", COL_MOV)
    cargar_movimientos.clear()


def guardar_abonos(df: pd.DataFrame) -> None:
    _escribir(df, ABONOS_PATH, "Abonos", COL_ABONO)
    cargar_abonos.clear()


def turno_abierto(turnos: pd.DataFrame, caja: str) -> pd.Series | None:
    """Turno en estado Abierto para una caja, o None."""
    if not len(turnos):
        return None
    m = turnos[(turnos["Estado"] == "Abierto") & (turnos["Caja"] == caja)]
    return m.iloc[0] if len(m) else None


def resumen_turno(ventas: pd.DataFrame, turno: pd.Series) -> dict:
    """Totales por método de pago desde la apertura del turno (misma caja).

    Devuelve {Efectivo, Tarjeta, Transferencia, Otro, Total, Tickets}.
    """
    res = {m: 0 for m in METODOS_ARQUEO}
    res.update({"Total": 0, "Tickets": 0})
    if ventas is None or not len(ventas):
        return res
    v = ventas.copy()
    v["Fecha"] = pd.to_datetime(v["Fecha"], errors="coerce")
    ape = pd.to_datetime(turno["Fecha_Apertura"], errors="coerce")
    f = v[(v["Fecha"] >= ape) & (v["Caja"] == turno["Caja"])].copy()
    if not len(f):
        return res
    f["Total"] = pd.to_numeric(f["Total"], errors="coerce").fillna(0).astype(int)
    for m in METODOS_ARQUEO:
        res[m] = int(f.loc[f["Metodo_Pago"] == m, "Total"].sum())
    res["Total"] = int(f["Total"].sum())
    res["Tickets"] = int(f["ID_Venta"].nunique())
    return res


def movimientos_turno(movs: pd.DataFrame, id_turno: str) -> pd.DataFrame:
    if movs is None or not len(movs):
        return pd.DataFrame(columns=COL_MOV)
    return movs[movs["ID_Turno"] == id_turno].copy()


def totales_movimientos(movs_turno: pd.DataFrame) -> tuple[int, int]:
    """(entradas, retiros) en Gs. del turno."""
    if movs_turno is None or not len(movs_turno):
        return 0, 0
    m = movs_turno.copy()
    m["Monto"] = pd.to_numeric(m["Monto"], errors="coerce").fillna(0).astype(int)
    ent = int(m.loc[m["Tipo"] == "Entrada", "Monto"].sum())
    ret = int(m.loc[m["Tipo"] == "Retiro", "Monto"].sum())
    return ent, ret


def calcular_arqueo(turno: pd.Series, resumen: dict, entradas: int, retiros: int,
                    conteo: dict) -> dict:
    """Arqueo ciego: compara lo esperado vs. lo contado en mano.

    esperado_efectivo = fondo + ventas_efectivo + entradas - retiros.
    diferencia > 0 sobrante, < 0 faltante. Todo en enteros Gs.
    """
    fondo = _a_entero(turno["Fondo_Inicial"])
    esp_ef = fondo + int(resumen.get("Efectivo", 0)) + int(entradas) - int(retiros)
    cont_ef = _a_entero(conteo.get("Efectivo", 0))
    dif = cont_ef - esp_ef
    por_metodo = {}
    for m in METODOS_ARQUEO:
        if m == "Efectivo":
            por_metodo[m] = {"esperado": esp_ef, "contado": cont_ef, "diferencia": dif}
        else:
            esp = int(resumen.get(m, 0))
            cont = _a_entero(conteo.get(m, esp))  # por defecto se confirma el sistema
            por_metodo[m] = {"esperado": esp, "contado": cont, "diferencia": cont - esp}
    return {"esperado_efectivo": esp_ef, "contado_efectivo": cont_ef,
            "diferencia": dif, "por_metodo": por_metodo,
            "total_ventas": int(resumen.get("Total", 0)),
            "tickets": int(resumen.get("Tickets", 0))}


# ------------------------------------------------- WhatsApp Deudores Pro
def wa_numero(telefono: str | None) -> str | None:
    """Normaliza un teléfono paraguayo a formato wa.me (5959XXXXXXXX).

    Acepta '0981 123456', '0981123456', '+595981123456', '595981123456'.
    Devuelve None si no hay dígitos suficientes.
    """
    dig = re.sub(r"\D", "", str(telefono or ""))
    if dig.startswith("595"):
        dig = dig[3:]
    dig = dig.lstrip("0")
    if len(dig) < 9 or not dig.isdigit():
        return None
    return "595" + dig


def mensaje_cobro(cliente: str, saldo_gs: str, vencimiento: str, concepto: str) -> str:
    return (f"Hola {cliente}, le saludamos de Librería MM. "
            f"Le recordamos que cuenta con un saldo pendiente de {saldo_gs} "
            f"con vencimiento el {vencimiento}. Detalle: {concepto}. "
            f"¡Gracias por su preferencia! 📚")


def wa_link(numero_wa: str, mensaje: str) -> str:
    return f"https://wa.me/{numero_wa}?text={quote(mensaje)}"


# ------------------------------------------------- Comprobantes imprimibles
_ESTILO_TICKET = (
    "body{font-family:'Courier New',Courier,monospace;margin:0;padding:8px;color:#000;}"
    ".ticket{max-width:320px;margin:0 auto;}"
    "h2,p{text-align:center;margin:4px 0;font-size:14px;}h2{font-size:18px;}"
    "table{width:100%;border-collapse:collapse;font-size:13px;}"
    "th{border-top:1px dashed #000;border-bottom:1px dashed #000;text-align:left;padding:4px 2px;}"
    "td{padding:3px 2px;border-bottom:1px dotted #bbb;vertical-align:top;}"
    ".c{text-align:center;}.r{text-align:right;white-space:nowrap;}"
    ".total{font-size:17px;font-weight:bold;text-align:right;border-top:2px solid #000;"
    "margin-top:6px;padding-top:6px;}"
    ".ok{color:#0a0;font-weight:bold;}.bad{color:#c00;font-weight:bold;}"
    ".btn{display:block;width:100%;margin:10px 0;padding:10px;font-size:15px;font-weight:bold;"
    "background:#1F4E79;color:#fff;border:none;border-radius:6px;cursor:pointer;}"
    "@media print{.btn{display:none;}body{padding:0;}}"
)


def _fmt_gs(n) -> str:
    try:
        return "Gs. " + f"{int(round(float(n))):,}".replace(",", ".")
    except (TypeError, ValueError):
        return "Gs. 0"


def comprobante_cierre_html(turno: pd.Series, resumen: dict, entradas: int,
                            retiros: int, arq: dict, auto_print: bool = False) -> str:
    """Comprobante de Corte Z (80mm) con comparativa por método y diferencia."""
    filas = "\n".join(
        f"<tr><td>{m}</td><td class='r'>{_fmt_gs(v['esperado'])}</td>"
        f"<td class='r'>{_fmt_gs(v['contado'])}</td>"
        f"<td class='r'>{_fmt_gs(v['diferencia'])}</td></tr>"
        for m, v in arq["por_metodo"].items())
    dif = arq["diferencia"]
    clase = "ok" if dif == 0 else ("ok" if dif > 0 else "bad")
    leyenda = "SIN DIFERENCIA ✅" if dif == 0 else (f"SOBRANTE { _fmt_gs(dif)}" if dif > 0
                                                     else f"FALTANTE {_fmt_gs(-dif)}")
    auto = "<script>window.onload=function(){window.print();};</script>" if auto_print else ""
    return (f"<html><head><meta charset='utf-8'><style>{_ESTILO_TICKET}</style></head>"
            f"<body><div class='ticket'><h2>📚 Librería MM</h2>"
            f"<p><b>CORTE Z — ARQUEO DE CAJA</b><br>Turno <b>{html.escape(str(turno['ID_Turno']))}</b> · "
            f"{html.escape(str(turno['Caja']))}<br>Apertura: {html.escape(str(turno['Fecha_Apertura']))}<br>"
            f"Cierre: {ahora_str()}<br>Cajero: <b>{html.escape(str(turno['Cajero_Nombre']))}</b></p>"
            f"<p>Fondo inicial: {_fmt_gs(turno['Fondo_Inicial'])} · "
            f"Ventas: {_fmt_gs(arq['total_ventas'])} ({arq['tickets']} tickets)<br>"
            f"Entradas: {_fmt_gs(entradas)} · Retiros: {_fmt_gs(retiros)}</p>"
            f"<table><tr><th>Método</th><th>Esper.</th><th>Cont.</th><th>Dif.</th></tr>{filas}</table>"
            f"<div class='total {clase}'>{leyenda}</div>"
            f"<p>Firma cajero: ______________<br>¡Gracias!</p>"
            f"<button class='btn' onclick='window.print()'>🖨️ Imprimir cierre</button>"
            f"</div>{auto}</body></html>")


def recibo_abono_html(abono: dict, saldo_restante, total_deuda) -> str:
    """Recibo digital de un abono/pago parcial (80mm, imprimible)."""
    return (f"<html><head><meta charset='utf-8'><style>{_ESTILO_TICKET}</style></head>"
            f"<body><div class='ticket'><h2>📚 Librería MM</h2>"
            f"<p><b>RECIBO DE PAGO N° {html.escape(str(abono.get('ID_Abono', '')))}</b><br>"
            f"{html.escape(str(abono.get('Fecha', '')))}<br>"
            f"Deuda: <b>{html.escape(str(abono.get('ID_Deuda', '')))}</b> · "
            f"Cliente: <b>{html.escape(str(abono.get('Cliente', '')))}</b></p>"
            f"<table><tr><td>Total deuda</td><td class='r'>{_fmt_gs(total_deuda)}</td></tr>"
            f"<tr><td>Abono ({html.escape(str(abono.get('Metodo', 'Efectivo')))})</td>"
            f"<td class='r'>{_fmt_gs(abono.get('Monto', 0))}</td></tr>"
            f"<tr><td><b>Saldo restante</b></td><td class='r'><b>{_fmt_gs(saldo_restante)}</b></td></tr>"
            f"</table><p>Recibido por: <b>{html.escape(str(abono.get('Cajero_Nombre', '')))}</b></p>"
            f"<p>¡Gracias por su pago!</p>"
            f"<button class='btn' onclick='window.print()'>🖨️ Imprimir recibo</button>"
            f"</div></body></html>")
