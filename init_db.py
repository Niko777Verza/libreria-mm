"""Inicializa la base de datos local en Excel para Librería MM.

Crea data/inventario.xlsx, data/ventas.xlsx y data/deudores.xlsx
con el esquema esperado por app.py + datos de ejemplo.

Uso:
    python init_db.py          # crea solo si no existen
    python init_db.py --force  # regenera todo (¡borra cambios!)
"""
from __future__ import annotations

import sys
from datetime import datetime, timedelta
from pathlib import Path

import pandas as pd

BASE = Path(__file__).resolve().parent
DATA_DIR = BASE / "data"

COLUMNAS_INV = ["SKU", "Nombre", "Categoria", "Costo", "Precio_Sugerido",
                "Precio_Actual", "Stock_Actual", "Stock_Minimo", "Fecha_Actualizacion"]
COLUMNAS_VENTAS = ["ID_Venta", "Fecha", "SKU", "Producto", "Categoria", "Cantidad",
                   "Precio_Unitario", "Total", "Metodo_Pago", "Caja", "Notas"]
COLUMNAS_DEUD = ["ID", "Cliente", "Telefono", "Concepto", "Monto_Total",
                 "Monto_Pagado", "Saldo", "Fecha_Prestamo", "Fecha_Vencimiento", "Estado"]


def inventario_ejemplo() -> pd.DataFrame:
    hoy = datetime.now().strftime("%Y-%m-%d %H:%M")
    filas = [
        ("9789504981931", "El Principito — Antoine de Saint-Exupéry", "Libros", 3500, 25, 5),
        ("9789877387891", "Cien años de soledad — García Márquez", "Libros", 6800, 18, 4),
        ("9789501142132", "Martín Fierro — Edición escolar", "Libros", 2900, 30, 6),
        ("CUAD-RIV-100", "Cuaderno Rivadavia tapa dura 100h rayado", "Útiles", 4200, 40, 10),
        ("LAP-FAB-12", "Lápices de colores Faber x12 largos", "Útiles", 5100, 24, 8),
        ("BOL-BIC-AZ", "Bolígrafo Bic cristal azul", "Útiles", 450, 200, 50),
        ("CART-A4-500", "Resma A4 80g x500 hojas", "Útiles", 7800, 15, 5),
        ("MOCH-JAN-NEG", "Mochila Jansport negra 25L", "Útiles", 28000, 6, 2),
        ("FOT-A4-BN", "Fotocopia A4 blanco y negro (servicio)", "Fotocopias", 60, 9999, 0),
        ("IMP-COLOR", "Impresión color A4 (servicio)", "Servicios", 150, 9999, 0),
        ("ANILLADO", "Anillado hasta 100 hojas (servicio)", "Servicios", 800, 9999, 0),
        ("PLAST-A4", "Plastificado A4 (servicio)", "Servicios", 500, 9999, 0),
    ]
    regs = []
    for sku, nombre, cat, costo, stock, minimo in filas:
        sugerido = round(costo * 1.5, 2)
        # Precio_Actual con redondeo comercial de ejemplo (oferta / terminación en 00/90)
        actual = float(sugerido) if cat in ("Fotocopias", "Servicios") else float(round(sugerido / 100) * 100)
        regs.append([sku, nombre, cat, float(costo), sugerido, actual, int(stock), int(minimo), hoy])
    return pd.DataFrame(regs, columns=COLUMNAS_INV)


def ventas_ejemplo(inv: pd.DataFrame) -> pd.DataFrame:
    base = datetime.now().replace(hour=10, minute=30, second=0, microsecond=0)
    picks = [
        (0, 2, "Efectivo", "Caja 1", 0), (5, 5, "Efectivo", "Caja 1", 0),
        (3, 1, "Tarjeta", "Caja 1", 1), (4, 2, "Transferencia", "Caja 2", 1),
        (8, 20, "Efectivo", "Caja 1", 2), (6, 1, "Tarjeta", "Caja 2", 3),
        (1, 1, "Efectivo", "Caja 1", 5), (9, 8, "Efectivo", "Caja 2", 6),
    ]
    regs = []
    for i, (idx, cant, metodo, caja, dias_atras) in enumerate(picks):
        p = inv.iloc[idx]
        fecha = base - timedelta(days=dias_atras, hours=i)
        total = round(cant * float(p["Precio_Actual"]), 2)
        regs.append([f"V-{fecha.strftime('%Y%m%d')}-{1000+i}", fecha.strftime("%Y-%m-%d %H:%M"),
                     p["SKU"], p["Nombre"], p["Categoria"], int(cant),
                     float(p["Precio_Actual"]), total, metodo, caja, "Venta de ejemplo"])
    return pd.DataFrame(regs, columns=COLUMNAS_VENTAS)


def deudores_ejemplo() -> pd.DataFrame:
    hoy = datetime.now()
    filas = [
        ("D-0001", "María González", "381-555-0101", "Útiles escolares (cta. cte.)", 15500, 5000,
         (hoy - timedelta(days=10)).strftime("%Y-%m-%d"), (hoy + timedelta(days=20)).strftime("%Y-%m-%d"), "Parcial"),
        ("D-0002", "Juan Pérez", "381-555-0102", "Fotocopias apuntes facultad", 3200, 0,
         (hoy - timedelta(days=40)).strftime("%Y-%m-%d"), (hoy - timedelta(days=10)).strftime("%Y-%m-%d"), "Vencido"),
        ("D-0003", "Escuela San Martín", "381-555-0103", "Resmas A4 x10", 78000, 78000,
         (hoy - timedelta(days=30)).strftime("%Y-%m-%d"), hoy.strftime("%Y-%m-%d"), "Pagado"),
        ("D-0004", "Lucía Fernández", "381-555-0104", "Libro Cien años de soledad", 10200, 0,
         hoy.strftime("%Y-%m-%d"), (hoy + timedelta(days=30)).strftime("%Y-%m-%d"), "Pendiente"),
    ]
    regs = []
    for _id, cli, tel, conc, tot, pag, f1, f2, est in filas:
        regs.append([_id, cli, tel, conc, float(tot), float(pag), float(tot - pag), f1, f2, est])
    return pd.DataFrame(regs, columns=COLUMNAS_DEUD)


def guardar(df: pd.DataFrame, nombre: str, hoja: str) -> Path:
    DATA_DIR.mkdir(exist_ok=True)
    path = DATA_DIR / nombre
    with pd.ExcelWriter(path, engine="openpyxl") as w:
        df.to_excel(w, sheet_name=hoja, index=False)
        ws = w.sheets[hoja]
        for col in ws.columns:
            ancho = max((len(str(c.value)) if c.value is not None else 0) for c in col) + 2
            ws.column_dimensions[col[0].column_letter].width = min(ancho, 48)
    return path


def main() -> None:
    force = "--force" in sys.argv
    DATA_DIR.mkdir(exist_ok=True)
    inv = inventario_ejemplo()
    ven = ventas_ejemplo(inv)
    deu = deudores_ejemplo()
    for df, nombre, hoja in [(inv, "inventario.xlsx", "Inventario"),
                             (ven, "ventas.xlsx", "Ventas"),
                             (deu, "deudores.xlsx", "Deudores")]:
        dest = DATA_DIR / nombre
        if dest.exists() and not force:
            print(f"⏭️  {dest.name} ya existe, se conserva. Usa --force para regenerar.")
            continue
        guardar(df, nombre, hoja)
        print(f"✅ {dest} ({len(df)} filas, hoja '{hoja}')")
    print("\nListo. Ejecuta:  streamlit run app.py")


if __name__ == "__main__":
    main()
