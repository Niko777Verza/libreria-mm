"""Inicializa la base de datos local en Excel para Librería MM (Gs. sin decimales).

Crea data/inventario.xlsx, data/ventas.xlsx, data/deudores.xlsx y
data/categorias.txt con el esquema esperado por app.py + datos de ejemplo
en Guaraníes.

Uso:
    python init_db.py          # crea solo si no existen
    python init_db.py --force  # regenera todo (¡borra cambios!)
"""
from __future__ import annotations

import hashlib
import sys
from datetime import datetime, timedelta
from pathlib import Path

import pandas as pd

BASE = Path(__file__).resolve().parent
DATA_DIR = BASE / "data"

COLUMNAS_INV = ["SKU", "Nombre", "Categoria", "Costo", "Precio_Sugerido",
                "Precio_Actual", "Stock_Actual", "Stock_Minimo", "Fecha_Actualizacion"]
COLUMNAS_VENTAS = ["ID_Venta", "Fecha", "SKU", "Producto", "Categoria", "Cantidad",
                   "Precio_Unitario", "Total", "Metodo_Pago", "Caja", "Notas",
                   "Vendedor_Doc", "Vendedor_Nombre"]
COLUMNAS_DEUD = ["ID", "Cliente", "Telefono", "Concepto", "Monto_Total",
                 "Monto_Pagado", "Saldo", "Fecha_Prestamo", "Fecha_Vencimiento", "Estado"]
COLUMNAS_USU = ["Documento", "Nombre", "Password_Hash", "Rol", "Activo", "Fecha_Creacion"]
COLUMNAS_TURNO = ["ID_Turno", "Fecha_Apertura", "Fecha_Cierre", "Cajero_Doc", "Cajero_Nombre",
                  "Caja", "Fondo_Inicial", "Estado", "Ventas_Efectivo", "Ventas_Tarjeta",
                  "Ventas_Transferencia", "Ventas_Otro", "Entradas", "Retiros",
                  "Esperado_Efectivo", "Contado_Efectivo", "Diferencia", "Notas_Cierre"]
COLUMNAS_MOV = ["ID_Mov", "ID_Turno", "Fecha", "Tipo", "Concepto", "Monto",
                "Cajero_Doc", "Cajero_Nombre"]
COLUMNAS_ABONO = ["ID_Abono", "ID_Deuda", "Cliente", "Fecha", "Monto",
                  "Metodo", "Cajero_Doc", "Cajero_Nombre", "Notas"]


def inventario_ejemplo() -> pd.DataFrame:
    hoy = datetime.now().strftime("%Y-%m-%d %H:%M")
    # Precios en Guaraníes (enteros, sin decimales)
    filas = [
        ("9789504981931", "El Principito — Antoine de Saint-Exupéry", "Libros", 35000, 25, 5),
        ("9789877387891", "Cien años de soledad — García Márquez", "Libros", 68000, 18, 4),
        ("9789501142132", "Martín Fierro — Edición escolar", "Libros", 29000, 30, 6),
        ("CUAD-RIV-100", "Cuaderno Rivadavia tapa dura 100h rayado", "Útiles", 42000, 40, 10),
        ("LAP-FAB-12", "Lápices de colores Faber x12 largos", "Útiles", 51000, 24, 8),
        ("BOL-BIC-AZ", "Bolígrafo Bic cristal azul", "Útiles", 4500, 200, 50),
        ("CART-A4-500", "Resma A4 80g x500 hojas", "Útiles", 78000, 15, 5),
        ("MOCH-JAN-NEG", "Mochila Jansport negra 25L", "Útiles", 280000, 6, 2),
        ("FOT-A4-BN", "Fotocopia A4 blanco y negro (servicio)", "Fotocopias", 500, 9999, 0),
        ("IMP-COLOR", "Impresión color A4 (servicio)", "Servicios", 2000, 9999, 0),
        ("ANILLADO", "Anillado hasta 100 hojas (servicio)", "Servicios", 8000, 9999, 0),
        ("PLAST-A4", "Plastificado A4 (servicio)", "Servicios", 5000, 9999, 0),
    ]
    regs = []
    for sku, nombre, cat, costo, stock, minimo in filas:
        sugerido = int(round(costo * 1.5))  # Precio Sugerido = Costo * 1.5
        actual = sugerido if cat in ("Fotocopias", "Servicios") else int(round(sugerido / 1000) * 1000)
        regs.append([sku, nombre, cat, int(costo), sugerido, actual, int(stock), int(minimo), hoy])
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
        pu = int(p["Precio_Actual"])
        total = int(cant * pu)
        regs.append([f"V-{fecha.strftime('%Y%m%d')}-{1000+i}", fecha.strftime("%Y-%m-%d %H:%M"),
                     p["SKU"], p["Nombre"], p["Categoria"], int(cant),
                     pu, total, metodo, caja, "Venta de ejemplo", "5729679", "Cesar Verza"])
    return pd.DataFrame(regs, columns=COLUMNAS_VENTAS)


def deudores_ejemplo() -> pd.DataFrame:
    hoy = datetime.now()
    filas = [
        ("D-0001", "María González", "381-555-0101", "Útiles escolares (cta. cte.)", 155000, 50000,
         (hoy - timedelta(days=10)).strftime("%Y-%m-%d"), (hoy + timedelta(days=20)).strftime("%Y-%m-%d"), "Parcial"),
        ("D-0002", "Juan Pérez", "381-555-0102", "Fotocopias apuntes facultad", 32000, 0,
         (hoy - timedelta(days=40)).strftime("%Y-%m-%d"), (hoy - timedelta(days=10)).strftime("%Y-%m-%d"), "Vencido"),
        ("D-0003", "Escuela San Martín", "381-555-0103", "Resmas A4 x10", 780000, 780000,
         (hoy - timedelta(days=30)).strftime("%Y-%m-%d"), hoy.strftime("%Y-%m-%d"), "Pagado"),
        ("D-0004", "Lucía Fernández", "381-555-0104", "Libro Cien años de soledad", 102000, 0,
         hoy.strftime("%Y-%m-%d"), (hoy + timedelta(days=30)).strftime("%Y-%m-%d"), "Pendiente"),
    ]
    regs = []
    for _id, cli, tel, conc, tot, pag, f1, f2, est in filas:
        regs.append([_id, cli, tel, conc, int(tot), int(pag), int(tot - pag), f1, f2, est])
    return pd.DataFrame(regs, columns=COLUMNAS_DEUD)


def usuarios_ejemplo() -> pd.DataFrame:
    hoy = datetime.now().strftime("%Y-%m-%d %H:%M")
    return pd.DataFrame([
        ["5729679", "Cesar Verza",
         hashlib.sha256("Nicolas08.L".encode("utf-8")).hexdigest(),
         "SuperUsuario", True, hoy],
        ["1000001", "Cajero Ejemplo",
         hashlib.sha256("caja1234".encode("utf-8")).hexdigest(),
         "Cajero", True, hoy],
    ], columns=COLUMNAS_USU)


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
    usu = usuarios_ejemplo()
    for df, nombre, hoja in [(inv, "inventario.xlsx", "Inventario"),
                             (ven, "ventas.xlsx", "Ventas"),
                             (deu, "deudores.xlsx", "Deudores"),
                             (usu, "usuarios.xlsx", "Usuarios")]:
        dest = DATA_DIR / nombre
        if dest.exists() and not force:
            print(f"⏭️  {dest.name} ya existe, se conserva. Usa --force para regenerar.")
            continue
        guardar(df, nombre, hoja)
        print(f"✅ {dest} ({len(df)} filas, hoja '{hoja}')")
    # Archivos de caja/abonos: solo se crean vacíos si no existen (nunca se pisan)
    for df_v, nombre, hoja in [(pd.DataFrame(columns=COLUMNAS_TURNO), "caja_turnos.xlsx", "Turnos"),
                               (pd.DataFrame(columns=COLUMNAS_MOV), "caja_movimientos.xlsx", "Movimientos"),
                               (pd.DataFrame(columns=COLUMNAS_ABONO), "abonos.xlsx", "Abonos")]:
        dest = DATA_DIR / nombre
        if not dest.exists():
            guardar(df_v, nombre, hoja)
            print(f"✅ {dest} (vacío, hoja '{hoja}')")
    # Categorías iniciales persistidas
    cat_path = DATA_DIR / "categorias.txt"
    if not cat_path.exists() or force:
        cats = sorted(inv["Categoria"].unique().tolist() + ["Juguetes", "Regalería", "Otros"])
        cat_path.write_text("\n".join(sorted(set(cats))), encoding="utf-8")
        print(f"✅ {cat_path} ({len(cats)} categorías)")
    print("\nListo. Ejecuta:  streamlit run app.py")


if __name__ == "__main__":
    main()
