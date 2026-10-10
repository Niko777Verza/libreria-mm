"""Librería MM — Sistema de Inventario, Ventas/Caja, Deudores y Reportes.
Stack: Streamlit + Pandas + Excel local (data/*.xlsx).
Ejecución: streamlit run app.py
"""
from __future__ import annotations

import io
import re
from datetime import date, datetime
from pathlib import Path

import pandas as pd
import streamlit as st

# ---------------------------------------------------------------- Config
APP_TITULO = "📚 Librería MM — Gestión"
BASE = Path(__file__).resolve().parent
DATA_DIR = BASE / "data"
INV_PATH = DATA_DIR / "inventario.xlsx"
VENTAS_PATH = DATA_DIR / "ventas.xlsx"
DEUD_PATH = DATA_DIR / "deudores.xlsx"

MARGEN = 1.5  # Precio_Sugerido = Costo * 1.5
CATEGORIAS_DEFAULT = ["Libros", "Útiles", "Fotocopias", "Servicios", "Juguetes", "Regalería", "Otros"]
METODOS_PAGO = ["Efectivo", "Tarjeta", "Transferencia", "Cuenta corriente", "Mercado Pago", "Otro"]
CAJAS = ["Caja 1", "Caja 2", "Caja 3"]

COL_INV = ["SKU", "Nombre", "Categoria", "Costo", "Precio_Sugerido",
           "Precio_Actual", "Stock_Actual", "Stock_Minimo", "Fecha_Actualizacion"]
COL_VENTAS = ["ID_Venta", "Fecha", "SKU", "Producto", "Categoria", "Cantidad",
              "Precio_Unitario", "Total", "Metodo_Pago", "Caja", "Notas"]
COL_DEUD = ["ID", "Cliente", "Telefono", "Concepto", "Monto_Total",
            "Monto_Pagado", "Saldo", "Fecha_Prestamo", "Fecha_Vencimiento", "Estado"]

st.set_page_config(page_title="Librería MM", page_icon="📚", layout="wide")

# ------------------------------------------------------- Utilidades base
def ahora_str() -> str:
    return datetime.now().strftime("%Y-%m-%d %H:%M")


def recalcular_sugerido(costo: float) -> float:
    try:
        return round(float(costo) * MARGEN, 2)
    except (TypeError, ValueError):
        return 0.0


def df_vacio_inv() -> pd.DataFrame:
    return pd.DataFrame(columns=COL_INV)


def df_vacio_ventas() -> pd.DataFrame:
    return pd.DataFrame(columns=COL_VENTAS)


def df_vacio_deud() -> pd.DataFrame:
    return pd.DataFrame(columns=COL_DEUD)


def _leer_excel_seguro(path: Path, hoja: str, vacio: pd.DataFrame) -> pd.DataFrame:
    if not path.exists():
        return vacio.copy()
    try:
        df = pd.read_excel(path, sheet_name=hoja, dtype={"SKU": str})
        # Garantizar todas las columnas esperadas
        for c in vacio.columns:
            if c not in df.columns:
                df[c] = pd.NA
        return df[list(vacio.columns)].copy()
    except Exception as e:  # noqa: BLE001
        st.warning(f"No se pudo leer {path.name}: {e}. Se muestra tabla vacía.")
        return vacio.copy()


def _guardar_excel(df: pd.DataFrame, path: Path, hoja: str) -> None:
    DATA_DIR.mkdir(exist_ok=True)
    with pd.ExcelWriter(path, engine="openpyxl") as w:
        df.to_excel(w, sheet_name=hoja, index=False)
        ws = w.sheets[hoja]
        for col in ws.columns:
            ancho = max((len(str(c.value)) if c.value is not None else 0) for c in col) + 2
            ws.column_dimensions[col[0].column_letter].width = min(ancho, 46)


@st.cache_data(ttl=5)
def cargar_inventario() -> pd.DataFrame:
    df = _leer_excel_seguro(INV_PATH, "Inventario", df_vacio_inv())
    if len(df):
        df["SKU"] = df["SKU"].astype(str).str.strip()
        for c in ("Costo", "Precio_Sugerido", "Precio_Actual"):
            df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0)
        for c in ("Stock_Actual", "Stock_Minimo"):
            df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0).astype(int)
        # Recalcular sugerido siempre desde el costo (fórmula automática)
        df["Precio_Sugerido"] = (df["Costo"] * MARGEN).round(2)
    return df


@st.cache_data(ttl=5)
def cargar_ventas() -> pd.DataFrame:
    df = _leer_excel_seguro(VENTAS_PATH, "Ventas", df_vacio_ventas())
    if len(df):
        df["Fecha"] = pd.to_datetime(df["Fecha"], errors="coerce")
        for c in ("Cantidad",):
            df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0).astype(int)
        for c in ("Precio_Unitario", "Total"):
            df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0)
    return df


@st.cache_data(ttl=5)
def cargar_deudores() -> pd.DataFrame:
    df = _leer_excel_seguro(DEUD_PATH, "Deudores", df_vacio_deud())
    if len(df):
        for c in ("Monto_Total", "Monto_Pagado", "Saldo"):
            df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0)
    return df


def guardar_inventario(df: pd.DataFrame) -> None:
    df = df.copy()
    df["Precio_Sugerido"] = (pd.to_numeric(df["Costo"], errors="coerce").fillna(0) * MARGEN).round(2)
    df["Fecha_Actualizacion"] = ahora_str()
    _guardar_excel(df[COL_INV], INV_PATH, "Inventario")
    st.cache_data.clear()


def guardar_ventas(df: pd.DataFrame) -> None:
    _guardar_excel(df[COL_VENTAS], VENTAS_PATH, "Ventas")
    st.cache_data.clear()


def guardar_deudores(df: pd.DataFrame) -> None:
    _guardar_excel(df[COL_DEUD], DEUD_PATH, "Deudores")
    st.cache_data.clear()


def estado_stock(stock: int, minimo: int) -> str:
    if minimo <= 0:
        return "✅ OK"
    if stock <= 0:
        return "🔴 Sin stock"
    if stock <= minimo:
        return "🔴 Crítico"
    if stock <= minimo * 1.5:
        return "🟡 Bajo"
    return "✅ OK"


def resaltar_stock(df: pd.DataFrame) -> pd.io.formats.style.Styler:
    """Colorea filas: rojo si Stock <= Mínimo, amarillo si <= Mínimo*1.5."""
    def color_fila(row):
        try:
            s, m = int(row["Stock_Actual"]), int(row["Stock_Minimo"])
        except (ValueError, TypeError, KeyError):
            return [""] * len(row)
        if s <= m:
            return ["background-color: #f8d7da"] * len(row)  # rojo claro
        if m > 0 and s <= m * 1.5:
            return ["background-color: #fff3cd"] * len(row)  # amarillo claro
        return [""] * len(row)
    cols = [c for c in COL_INV if c in df.columns]
    return df[cols].style.apply(color_fila, axis=1)


def df_a_excel_bytes(hojas: dict[str, pd.DataFrame]) -> bytes:
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as w:
        for nombre, df in hojas.items():
            df.to_excel(w, sheet_name=nombre[:31] or "Hoja", index=False)
    buf.seek(0)
    return buf.getvalue()


# ------------------------------------------- Importación flexible (Excel viejo)
SINONIMOS = {
    "sku": "SKU", "codigo": "SKU", "código": "SKU", "cod": "SKU",
    "codigo_barras": "SKU", "cod_barras": "SKU", "codigobarra": "SKU", "id": "SKU",
    "nombre": "Nombre", "producto": "Nombre", "descripcion": "Nombre",
    "descripción": "Nombre", "detalle": "Nombre", "articulo": "Nombre", "artículo": "Nombre",
    "categoria": "Categoria", "categoría": "Categoria", "rubro": "Categoria", "tipo": "Categoria",
    "costo": "Costo", "precio_compra": "Costo", "precio_costo": "Costo", "compra": "Costo",
    "precio_sugerido": "Precio_Sugerido", "sugerido": "Precio_Sugerido",
    "precio": "Precio_Actual", "precio_venta": "Precio_Actual", "precio_actual": "Precio_Actual",
    "pvp": "Precio_Actual", "venta": "Precio_Actual",
    "stock": "Stock_Actual", "stock_actual": "Stock_Actual", "cantidad": "Stock_Actual",
    "existencia": "Stock_Actual", "disponible": "Stock_Actual",
    "minimo": "Stock_Minimo", "mínimo": "Stock_Minimo", "stock_minimo": "Stock_Minimo",
    "min": "Stock_Minimo", "alerta": "Stock_Minimo", "critico": "Stock_Minimo",
    "fecha": "Fecha", "dia": "Fecha", "día": "Fecha",
    "metodo_pago": "Metodo_Pago", "pago": "Metodo_Pago", "medio_pago": "Metodo_Pago",
    "forma_pago": "Metodo_Pago", "metodo": "Metodo_Pago",
    "cliente": "Cliente", "nombre_cliente": "Cliente", "deudor": "Cliente",
    "monto": "Monto_Total", "total": "Total", "deuda": "Monto_Total",
    "importe": "Monto_Total", "debe": "Saldo", "saldo": "Saldo",
}

def _norm(col: str) -> str:
    c = str(col).strip().lower()
    c = re.sub(r"[^a-z0-9áéíóúñ]+", "_", c).strip("_")
    c = (c.replace("á", "a").replace("é", "e").replace("í", "i")
           .replace("ó", "o").replace("ú", "u"))
    c = {"ano": "anio"}.get(c, c)
    return SINONIMOS.get(c, str(col).strip())


def normalizar_df_inventario(crudo: pd.DataFrame) -> pd.DataFrame:
    df = crudo.rename(columns={c: _norm(c) for c in crudo.columns}).copy()
    for c in COL_INV:
        if c not in df.columns:
            df[c] = pd.NA
    df["SKU"] = df["SKU"].astype(str).str.strip()
    df = df[df["SKU"].notna() & (df["SKU"] != "") & (df["SKU"].str.lower() != "nan")]
    df["Nombre"] = df["Nombre"].fillna("Sin descripción").astype(str)
    df["Categoria"] = df["Categoria"].fillna("Otros").astype(str)
    df["Costo"] = pd.to_numeric(df["Costo"], errors="coerce").fillna(0)
    df["Precio_Actual"] = pd.to_numeric(df["Precio_Actual"], errors="coerce")
    df["Precio_Sugerido"] = (df["Costo"] * MARGEN).round(2)
    # Si no traía precio actual, usar el sugerido
    df["Precio_Actual"] = df["Precio_Actual"].fillna(df["Precio_Sugerido"])
    df["Stock_Actual"] = pd.to_numeric(df["Stock_Actual"], errors="coerce").fillna(0).astype(int)
    df["Stock_Minimo"] = pd.to_numeric(df["Stock_Minimo"], errors="coerce").fillna(0).astype(int)
    df["Fecha_Actualizacion"] = ahora_str()
    return df[COL_INV].drop_duplicates(subset="SKU", keep="last").reset_index(drop=True)


def normalizar_df_ventas(crudo: pd.DataFrame, inv: pd.DataFrame) -> pd.DataFrame:
    df = crudo.rename(columns={c: _norm(c) for c in crudo.columns}).copy()
    # Ventas-Enero.xlsx típico: Fecha | Producto | Cantidad | Precio | Total | ...
    if "Nombre" in df.columns and "Producto" not in df.columns:
        df = df.rename(columns={"Nombre": "Producto"})
    if "Precio_Actual" in df.columns and "Precio_Unitario" not in df.columns:
        df = df.rename(columns={"Precio_Actual": "Precio_Unitario"})
    precios = {}
    if len(inv):
        precios = dict(zip(inv["SKU"].astype(str), zip(inv["Nombre"], inv["Categoria"], inv["Precio_Actual"])))
    regs = []
    for i, r in df.iterrows():
        sku = str(r.get("SKU", "") or "").strip()
        prod = str(r.get("Producto", "") or "").strip() or "Ítem importado"
        cat, pu = "Otros", 0.0
        if sku in precios:
            prod, cat, pu = precios[sku][0], precios[sku][1], float(precios[sku][2])
        elif prod != "Ítem importado" and len(inv):
            match = inv[inv["Nombre"].str.lower() == prod.lower()]
            if len(match):
                m = match.iloc[0]
                sku, cat, pu = str(m["SKU"]), str(m["Categoria"]), float(m["Precio_Actual"])
        cant = pd.to_numeric(pd.Series([r.get("Cantidad", 1)]), errors="coerce").fillna(1).astype(int).iloc[0]
        if cant <= 0:
            cant = 1
        pu = float(pd.to_numeric(pd.Series([r.get("Precio_Unitario", pu)]), errors="coerce").fillna(pu).iloc[0])
        fecha = pd.to_datetime(r.get("Fecha", datetime.now()), errors="coerce")
        if pd.isna(fecha):
            fecha = datetime.now()
        total = round(cant * pu, 2)
        regs.append([f"V-IMP-{fecha.strftime('%Y%m%d')}-{i+1:04d}", fecha,
                     sku or "IMPORTADO", prod, str(r.get("Categoria", cat)),
                     int(cant), pu, total,
                     str(r.get("Metodo_Pago", "Efectivo")), str(r.get("Caja", "Caja 1")),
                     str(r.get("Notas", "Importado desde Excel"))])
    out = pd.DataFrame(regs, columns=COL_VENTAS)
    out["Fecha"] = pd.to_datetime(out["Fecha"])
    return out


def normalizar_df_deudores(crudo: pd.DataFrame) -> pd.DataFrame:
    df = crudo.rename(columns={c: _norm(c) for c in crudo.columns}).copy()
    for c in COL_DEUD:
        if c not in df.columns:
            df[c] = pd.NA
    regs = []
    for i, r in df.iterrows():
        tot = float(pd.to_numeric(pd.Series([r.get("Monto_Total", r.get("Total", 0))]), errors="coerce").fillna(0).iloc[0])
        pag = float(pd.to_numeric(pd.Series([r.get("Monto_Pagado", 0)]), errors="coerce").fillna(0).iloc[0])
        regs.append([str(r.get("ID", f"D-IMP-{i+1:04d}")), str(r.get("Cliente", "Sin nombre")),
                     str(r.get("Telefono", "")), str(r.get("Concepto", "Deuda importada")),
                     tot, pag, round(tot - pag, 2),
                     str(r.get("Fecha_Prestamo", date.today())),
                     str(r.get("Fecha_Vencimiento", date.today())),
                     str(r.get("Estado", "Pendiente" if tot - pag > 0 else "Pagado"))])
    return pd.DataFrame(regs, columns=COL_DEUD)


# ------------------------------------------------------------------ Estado
if "carrito" not in st.session_state:
    st.session_state.carrito = []  # [{SKU, Nombre, Cantidad, Precio}]
if "pagina" not in st.session_state:
    st.session_state.pagina = "📊 Dashboard"

inv = cargar_inventario()
ventas = cargar_ventas()
deud = cargar_deudores()

# ------------------------------------------------------------------ Sidebar
with st.sidebar:
    st.title("📚 Librería MM")
    st.caption("Inventario · Caja · Deudores")
    st.session_state.pagina = st.radio(
        "Menú", ["📊 Dashboard", "📦 Inventario", "🧾 Ventas y Caja",
                 "🤝 Deudores", "🔄 Importar / Exportar", "📈 Reportes"],
        index=["📊 Dashboard", "📦 Inventario", "🧾 Ventas y Caja",
               "🤝 Deudores", "🔄 Importar / Exportar", "📈 Reportes"].index(st.session_state.pagina),
    )
    st.divider()
    n_crit = int(((inv["Stock_Actual"] <= inv["Stock_Minimo"]) & (len(inv) > 0)).sum()) if len(inv) else 0
    st.metric("⚠️ Productos en alerta", n_crit)
    if not INV_PATH.exists():
        st.warning("Aún no hay base local. Ejecuta `python init_db.py` o carga productos en Inventario.")

pg = st.session_state.pagina

# --------------------------------------------------------------- DASHBOARD
if pg == "📊 Dashboard":
    st.header("📊 Panel general")
    hoy = date.today()
    v = ventas.copy()
    if len(v):
        v["Fecha"] = pd.to_datetime(v["Fecha"], errors="coerce")
        v = v.dropna(subset=["Fecha"])
    v_hoy = v[v["Fecha"].dt.date == hoy] if len(v) else v
    v_mes = v[(v["Fecha"].dt.year == hoy.year) & (v["Fecha"].dt.month == hoy.month)] if len(v) else v

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("💰 Ventas hoy", f"${v_hoy['Total'].sum():,.0f}" if len(v_hoy) else "$0", f"{len(v_hoy)} ítems")
    c2.metric("🗓️ Ventas del mes", f"${v_mes['Total'].sum():,.0f}" if len(v_mes) else "$0")
    c3.metric("📦 Productos", len(inv))
    valor_costo = float((inv["Costo"] * inv["Stock_Actual"]).sum()) if len(inv) else 0
    c4.metric("🏷️ Valor inventario (costo)", f"${valor_costo:,.0f}")

    col_a, col_b = st.columns([2, 1])
    with col_a:
        st.subheader("Ventas por día (últimos 14 días)")
        if len(v):
            ult = v[v["Fecha"] >= pd.Timestamp(hoy) - pd.Timedelta(days=13)]
            if len(ult):
                diario = ult.groupby(ult["Fecha"].dt.date)["Total"].sum().reset_index()
                diario.columns = ["Fecha", "Total"]
                st.bar_chart(diario.set_index("Fecha"))
            else:
                st.info("Sin ventas en los últimos 14 días.")
        else:
            st.info("Todavía no hay ventas registradas. Ve a 🧾 Ventas y Caja.")
    with col_b:
        st.subheader("Ventas por método de pago")
        if len(v):
            st.bar_chart(v.groupby("Metodo_Pago")["Total"].sum())
        st.subheader("Top categorías")
        if len(v):
            st.bar_chart(v.groupby("Categoria")["Total"].sum().sort_values(ascending=False).head(5))

    st.subheader("🚨 Alertas de stock (Stock ≤ Mínimo)")
    if len(inv):
        alertas = inv[inv["Stock_Actual"] <= inv["Stock_Minimo"]].copy()
        if len(alertas):
            alertas["Estado"] = alertas.apply(lambda r: estado_stock(r["Stock_Actual"], r["Stock_Minimo"]), axis=1)
            st.dataframe(resaltar_stock(alertas), use_container_width=True)
        else:
            st.success("✅ Sin alertas: todo el stock está por encima del mínimo.")
    else:
        st.info("Carga productos para activar las alertas.")

# -------------------------------------------------------------- INVENTARIO
elif pg == "📦 Inventario":
    st.header("📦 Inventario")
    t1, t2 = st.tabs(["📋 Catálogo (editar)", "➕ Agregar producto"])

    with t2:
        with st.form("nuevo_prod", clear_on_submit=True):
            c1, c2, c3 = st.columns(3)
            sku = c1.text_input("SKU / Código de barras *")
            nombre = c2.text_input("Nombre del producto *")
            cat = c3.selectbox("Categoría", CATEGORIAS_DEFAULT + ["➕ Nueva…"])
            cat_nueva = st.text_input("Nueva categoría (si elegiste ➕ Nueva…)") if cat == "➕ Nueva…" else ""
            c4, c5, c6, c7 = st.columns(4)
            costo = c4.number_input("Costo de adquisición", min_value=0.0, step=10.0)
            st.caption(f"Precio sugerido automático: **${recalcular_sugerido(costo):,.2f}** (= Costo × 1.5)")
            precio = c5.number_input("Precio Actual de venta", min_value=0.0, step=10.0)
            stock = c6.number_input("Stock Actual", min_value=0, step=1)
            minimo = c7.number_input("Stock Mínimo (alerta)", min_value=0, step=1)
            if st.form_submit_button("💾 Guardar producto", type="primary"):
                categoria = cat_nueva.strip() or cat
                if not sku.strip() or not nombre.strip():
                    st.error("SKU y Nombre son obligatorios.")
                elif sku.strip() in set(inv["SKU"].astype(str)):
                    st.error(f"El SKU `{sku}` ya existe. Usa otro o edítalo en el catálogo.")
                else:
                    if precio == 0:
                        precio = recalcular_sugerido(costo)
                    nuevo = pd.DataFrame([[sku.strip(), nombre.strip(), categoria, float(costo),
                                           recalcular_sugerido(costo), float(precio),
                                           int(stock), int(minimo), ahora_str()]], columns=COL_INV)
                    guardar_inventario(pd.concat([inv, nuevo], ignore_index=True))
                    st.success(f"✅ `{nombre}` guardado. Precio sugerido: ${recalcular_sugerido(costo):,.2f}")
                    st.rerun()

    with t1:
        c1, c2, c3 = st.columns([2, 1, 1])
        busq = c1.text_input("🔎 Buscar (SKU o nombre)", "")
        fcat = c2.selectbox("Categoría", ["Todas"] + sorted(inv["Categoria"].dropna().unique().tolist()) if len(inv) else ["Todas"])
        solo_alertas = c3.checkbox("Solo alertas ⚠️")
        vista = inv.copy()
        if busq:
            vista = vista[vista["SKU"].str.contains(busq, case=False, na=False)
                          | vista["Nombre"].str.contains(busq, case=False, na=False)]
        if fcat != "Todas":
            vista = vista[vista["Categoria"] == fcat]
        if solo_alertas:
            vista = vista[vista["Stock_Actual"] <= vista["Stock_Minimo"]]
        vista["Estado"] = vista.apply(lambda r: estado_stock(r["Stock_Actual"], r["Stock_Minimo"]), axis=1)
        st.caption("💡 Edita **Precio_Actual, Stock_Actual y Stock_Minimo** directo en la tabla. "
                   "`Precio_Sugerido` se recalcula solo (Costo × 1.5). No edites el SKU.")
        edit = st.data_editor(
            vista[["SKU", "Nombre", "Categoria", "Costo", "Precio_Sugerido",
                   "Precio_Actual", "Stock_Actual", "Stock_Minimo", "Estado"]],
            use_container_width=True, num_rows="fixed",
            column_config={
                "SKU": st.column_config.TextColumn("SKU", disabled=True),
                "Precio_Sugerido": st.column_config.NumberColumn("Precio Sugerido (auto)", disabled=True, format="$%.2f"),
                "Precio_Actual": st.column_config.NumberColumn("Precio Actual (venta)", format="$%.2f"),
                "Costo": st.column_config.NumberColumn("Costo", format="$%.2f"),
                "Estado": st.column_config.TextColumn("Estado", disabled=True),
            },
            key="editor_inv",
        )
        col_g, col_d = st.columns([1, 1])
        if col_g.button("💾 Guardar cambios", type="primary"):
            base = inv.set_index("SKU")
            for _, r in edit.iterrows():
                sku = str(r["SKU"])
                if sku in base.index:
                    base.loc[sku, ["Nombre", "Categoria", "Costo", "Precio_Actual",
                                   "Stock_Actual", "Stock_Minimo"]] = [
                        r["Nombre"], r["Categoria"], float(r["Costo"] or 0),
                        float(r["Precio_Actual"] or 0), int(r["Stock_Actual"] or 0),
                        int(r["Stock_Minimo"] or 0)]
            guardar_inventario(base.reset_index())
            st.success("✅ Inventario actualizado (sugeridos recalculados).")
            st.rerun()
        with col_d.expander("🗑️ Eliminar producto"):
            if len(inv):
                sku_del = st.selectbox("SKU a eliminar", inv["SKU"].tolist())
                if st.button("Eliminar definitivamente"):
                    guardar_inventario(inv[inv["SKU"] != sku_del])
                    st.success(f"Eliminado `{sku_del}`.")
                    st.rerun()

        st.subheader("Vista con alertas de color")
        st.dataframe(resaltar_stock(vista), use_container_width=True)
        st.download_button("⬇️ Descargar inventario (.xlsx)",
                           df_a_excel_bytes({"Inventario": inv[COL_INV]}),
                           file_name=f"inventario-MM-{date.today()}.xlsx")

# ----------------------------------------------------------- VENTAS Y CAJA
elif pg == "🧾 Ventas y Caja":
    st.header("🧾 Ventas y Caja")
    if not len(inv):
        st.warning("Primero carga productos en 📦 Inventario.")
        st.stop()
    t1, t2 = st.tabs(["🛒 Nueva venta", "📜 Historial y totales"])

    with t1:
        st.subheader("Armar ticket")
        c1, c2, c3 = st.columns([2, 1, 1])
        opciones = {f"{r.SKU} — {r.Nombre} (${r.Precio_Actual:,.0f}, stock {r.Stock_Actual})": r.SKU
                    for r in inv.itertuples()}
        sel = c1.selectbox("Producto (SKU — nombre)", list(opciones.keys()))
        cant = c2.number_input("Cantidad", min_value=1, step=1, value=1)
        sku_sel = opciones[sel]
        fila = inv[inv["SKU"] == sku_sel].iloc[0]
        c3.metric("Subtotal", f"${float(fila['Precio_Actual']) * cant:,.0f}")
        if st.button("➕ Agregar al ticket"):
            if cant > int(fila["Stock_Actual"]) and int(fila["Stock_Actual"]) < 9000:
                st.error(f"Stock insuficiente: hay {int(fila['Stock_Actual'])} y pides {cant}.")
            else:
                st.session_state.carrito.append(
                    {"SKU": sku_sel, "Nombre": fila["Nombre"], "Cantidad": int(cant),
                     "Precio": float(fila["Precio_Actual"])})
                st.success(f"Agregado: {fila['Nombre']} x{cant}")
        if st.session_state.carrito:
            st.divider()
            car = pd.DataFrame(st.session_state.carrito)
            car["Subtotal"] = car["Cantidad"] * car["Precio"]
            st.dataframe(car, use_container_width=True)
            total = float(car["Subtotal"].sum())
            st.markdown(f"### Total: ${total:,.2f}")
            cc1, cc2, cc3 = st.columns(3)
            metodo = cc1.selectbox("Método de pago", METODOS_PAGO)
            caja = cc2.selectbox("Caja", CAJAS)
            notas = cc3.text_input("Notas (opcional)")
            ccc1, ccc2 = st.columns(2)
            if ccc1.button("✅ Cobrar e imprimir ticket", type="primary"):
                ticket = f"V-{datetime.now().strftime('%Y%m%d-%H%M%S')}"
                nuevas = [[ticket, ahora_str(), it["SKU"], it["Nombre"],
                           str(inv.loc[inv['SKU'] == it['SKU'], 'Categoria'].iloc[0]
                               if (inv['SKU'] == it['SKU']).any() else "Otros"),
                           it["Cantidad"], it["Precio"], round(it["Cantidad"] * it["Precio"], 2),
                           metodo, caja, notas] for it in st.session_state.carrito]
                guardar_ventas(pd.concat([ventas, pd.DataFrame(nuevas, columns=COL_VENTAS)],
                                         ignore_index=True))
                # Descuento automático de stock
                base = inv.set_index("SKU")
                for it in st.session_state.carrito:
                    if it["SKU"] in base.index and int(base.loc[it["SKU"], "Stock_Actual"]) < 9000:
                        base.loc[it["SKU"], "Stock_Actual"] = max(
                            0, int(base.loc[it["SKU"], "Stock_Actual"]) - it["Cantidad"])
                guardar_inventario(base.reset_index())
                st.session_state.carrito = []
                st.success(f"✅ Venta {ticket} registrada por ${total:,.2f} ({metodo} · {caja}). Stock descontado.")
                st.rerun()
            if ccc2.button("🧹 Vaciar ticket"):
                st.session_state.carrito = []
                st.rerun()
        else:
            st.info("El ticket está vacío. Agrega productos arriba.")

    with t2:
        st.subheader("Historial")
        v = ventas.copy()
        if not len(v):
            st.info("Sin ventas todavía.")
        else:
            v["Fecha"] = pd.to_datetime(v["Fecha"], errors="coerce")
            c1, c2, c3 = st.columns(3)
            d1 = c1.date_input("Desde", v["Fecha"].min().date())
            d2 = c2.date_input("Hasta", v["Fecha"].max().date())
            fm = c3.selectbox("Método", ["Todos"] + sorted(v["Metodo_Pago"].dropna().unique().tolist()))
            f = v[(v["Fecha"].dt.date >= d1) & (v["Fecha"].dt.date <= d2)]
            if fm != "Todos":
                f = f[f["Metodo_Pago"] == fm]
            st.dataframe(f.sort_values("Fecha", ascending=False), use_container_width=True)
            m1, m2, m3 = st.columns(3)
            m1.metric("Total período", f"${f['Total'].sum():,.0f}")
            m2.metric("Tickets", f["ID_Venta"].nunique())
            m3.metric("Ticket promedio",
                      f"${(f.groupby('ID_Venta')['Total'].sum().mean()):,.0f}" if len(f) else "$0")
            st.download_button("⬇️ Descargar reporte de ventas (.xlsx)",
                               df_a_excel_bytes({"Ventas": f[COL_VENTAS]}),
                               file_name=f"ventas-MM-{d1}_{d2}.xlsx")

# ---------------------------------------------------------------- DEUDORES
elif pg == "🤝 Deudores":
    st.header("🤝 Control de deudores (ex Deudores.xlsm)")
    t1, t2 = st.tabs(["📋 Cuentas", "➕ Nuevo / Abono"])
    with t2:
        with st.form("nuevo_deud", clear_on_submit=True):
            c1, c2 = st.columns(2)
            cli = c1.text_input("Cliente *")
            tel = c2.text_input("Teléfono")
            conc = st.text_input("Concepto *", placeholder="Ej. Útiles escolares / Fotocopias")
            c3, c4, c5 = st.columns(3)
            tot = c3.number_input("Monto total", min_value=0.0, step=100.0)
            fp = c4.date_input("Fecha préstamo", date.today())
            fv = c5.date_input("Vencimiento", date.today())
            if st.form_submit_button("💾 Registrar deuda", type="primary"):
                if not cli.strip() or tot <= 0:
                    st.error("Cliente y monto mayor a 0 son obligatorios.")
                else:
                    nid = f"D-{(len(deud) + 1):04d}"
                    nuevo = pd.DataFrame([[nid, cli.strip(), tel, conc, float(tot), 0.0,
                                           float(tot), str(fp), str(fv), "Pendiente"]], columns=COL_DEUD)
                    guardar_deudores(pd.concat([deud, nuevo], ignore_index=True))
                    st.success(f"✅ Deuda {nid} registrada.")
                    st.rerun()
        st.divider()
        st.subheader("💵 Registrar abono / pago")
        if len(deud):
            pendientes = deud[deud["Saldo"] > 0]
            if len(pendientes):
                sel = st.selectbox("Cuenta", pendientes.apply(
                    lambda r: f"{r['ID']} — {r['Cliente']} (saldo ${r['Saldo']:,.0f})", axis=1).tolist())
                _id = sel.split(" — ")[0]
                monto = st.number_input("Monto del abono", min_value=1.0, step=100.0)
                if st.button("Registrar abono"):
                    d = deud.copy()
                    i = d.index[d["ID"] == _id][0]
                    d.loc[i, "Monto_Pagado"] = float(d.loc[i, "Monto_Pagado"]) + float(monto)
                    d.loc[i, "Saldo"] = round(float(d.loc[i, "Monto_Total"]) - float(d.loc[i, "Monto_Pagado"]), 2)
                    if d.loc[i, "Saldo"] <= 0:
                        d.loc[i, "Saldo"] = 0.0
                        d.loc[i, "Estado"] = "Pagado"
                    else:
                        d.loc[i, "Estado"] = "Parcial"
                    guardar_deudores(d)
                    st.success(f"✅ Abono de ${monto:,.0f} registrado en {_id}.")
                    st.rerun()
            else:
                st.success("Sin saldos pendientes. 🎉")
    with t1:
        if not len(deud):
            st.info("Sin deudores. Registra el primero en la pestaña ➕ Nuevo / Abono.")
        else:
            d = deud.copy()
            d["Fecha_Vencimiento"] = pd.to_datetime(d["Fecha_Vencimiento"], errors="coerce")
            hoy = pd.Timestamp(date.today())
            d["Vencido"] = (d["Fecha_Vencimiento"] < hoy) & (d["Saldo"] > 0)
            d.loc[d["Vencido"], "Estado"] = "Vencido"
            c1, c2, c3, c4 = st.columns(4)
            c1.metric("💰 Por cobrar", f"${d['Saldo'].sum():,.0f}")
            c2.metric("🔴 Vencido", f"${d[d['Vencido']]['Saldo'].sum():,.0f}")
            c3.metric("Cuentas abiertas", int((d["Saldo"] > 0).sum()))
            c4.metric("Cuentas pagadas", int((d["Saldo"] <= 0).sum()))
            fest = st.selectbox("Filtrar", ["Todas", "Pendiente", "Parcial", "Vencido", "Pagado"])
            if fest != "Todas":
                d = d[d["Estado"] == fest]
            st.dataframe(d[COL_DEUD], use_container_width=True)
            st.download_button("⬇️ Descargar deudores (.xlsx)",
                               df_a_excel_bytes({"Deudores": d[COL_DEUD]}),
                               file_name=f"deudores-MM-{date.today()}.xlsx")

# ------------------------------------------------------- IMPORTAR/EXPORTAR
elif pg == "🔄 Importar / Exportar":
    st.header("🔄 Importar Excel viejo / Exportar actualizado")
    st.info("Acepta `.xlsx`, `.xlsm` y `.xls`. La app detecta columnas aunque se llamen "
            "distinto (ej. `Código`→SKU, `Descripción`→Nombre, `Precio`→Precio Actual). "
            "Inspirado en `Ventas-Enero.xlsx` y `Deudores.xlsm`.")

    c1, c2, c3 = st.columns(3)
    with c1:
        st.subheader("📦 Importar inventario")
        f = st.file_uploader("Archivo de productos", type=["xlsx", "xlsm", "xls"], key="imp_inv")
        if f:
            crudo = pd.read_excel(f, sheet_name=0, dtype=str)
            st.write("Vista previa del archivo:", crudo.head())
            norm = normalizar_df_inventario(crudo)
            st.write("Cómo quedará normalizado:", norm.head())
            modo = st.radio("Modo", ["Agregar/actualizar por SKU", "Reemplazar todo"], key="m_inv")
            if st.button("Importar inventario", type="primary"):
                if modo == "Reemplazar todo":
                    guardar_inventario(norm)
                else:
                    base = inv.set_index("SKU") if len(inv) else pd.DataFrame(columns=COL_INV).set_index("SKU")
                    for _, r in norm.set_index("SKU").iterrows():
                        base.loc[r.name] = r
                    guardar_inventario(base.reset_index())
                st.success(f"✅ {len(norm)} productos importados.")
                st.rerun()
    with c2:
        st.subheader("🧾 Importar ventas históricas")
        f = st.file_uploader("Ej. Ventas-Enero.xlsx", type=["xlsx", "xlsm", "xls"], key="imp_ven")
        if f:
            crudo = pd.read_excel(f, sheet_name=0, dtype=str)
            st.write("Vista previa:", crudo.head())
            norm = normalizar_df_ventas(crudo, inv)
            st.write("Normalizado:", norm.head())
            descontar = st.checkbox("Descontar stock al importar", value=False)
            if st.button("Importar ventas", type="primary"):
                guardar_ventas(pd.concat([ventas, norm], ignore_index=True))
                if descontar:
                    base = inv.set_index("SKU")
                    for _, r in norm.iterrows():
                        if r["SKU"] in base.index:
                            base.loc[r["SKU"], "Stock_Actual"] = max(
                                0, int(base.loc[r["SKU"], "Stock_Actual"]) - int(r["Cantidad"]))
                    guardar_inventario(base.reset_index())
                st.success(f"✅ {len(norm)} líneas de venta importadas (${norm['Total'].sum():,.0f}).")
                st.rerun()
    with c3:
        st.subheader("🤝 Importar deudores")
        f = st.file_uploader("Ej. Deudores.xlsm", type=["xlsx", "xlsm", "xls"], key="imp_deu")
        if f:
            crudo = pd.read_excel(f, sheet_name=0, dtype=str)
            st.write("Vista previa:", crudo.head())
            norm = normalizar_df_deudores(crudo)
            st.write("Normalizado:", norm.head())
            if st.button("Importar deudores", type="primary"):
                guardar_deudores(pd.concat([deud, norm], ignore_index=True))
                st.success(f"✅ {len(norm)} deudores importados.")
                st.rerun()

    st.divider()
    st.subheader("📤 Exportar todo (respaldo)")
    st.download_button("⬇️ Descargar RESPALDO COMPLETO (.xlsx: Inventario+Ventas+Deudores)",
                       df_a_excel_bytes({"Inventario": inv[COL_INV], "Ventas": ventas[COL_VENTAS],
                                         "Deudores": deud[COL_DEUD]}),
                       file_name=f"respaldo-libreria-MM-{date.today()}.xlsx")

# ---------------------------------------------------------------- REPORTES
else:
    st.header("📈 Reportes")
    v = ventas.copy()
    if not len(v):
        st.info("Sin ventas para reportar todavía.")
        st.stop()
    v["Fecha"] = pd.to_datetime(v["Fecha"], errors="coerce")
    v = v.dropna(subset=["Fecha"])
    c1, c2 = st.columns(2)
    d1 = c1.date_input("Desde", v["Fecha"].min().date())
    d2 = c2.date_input("Hasta", v["Fecha"].max().date())
    f = v[(v["Fecha"].dt.date >= d1) & (v["Fecha"].dt.date <= d2)]
    if not len(f):
        st.warning("Sin ventas en ese rango.")
        st.stop()
    k1, k2, k3, k4 = st.columns(4)
    k1.metric("Total", f"${f['Total'].sum():,.0f}")
    k2.metric("Unidades", int(f["Cantidad"].sum()))
    k3.metric("Tickets", f["ID_Venta"].nunique())
    k4.metric("Ticket promedio", f"${f.groupby('ID_Venta')['Total'].sum().mean():,.0f}")
    c1, c2 = st.columns(2)
    with c1:
        st.subheader("Por categoría")
        st.bar_chart(f.groupby("Categoria")["Total"].sum().sort_values(ascending=False))
        st.subheader("Por método de pago")
        st.bar_chart(f.groupby("Metodo_Pago")["Total"].sum())
    with c2:
        st.subheader("Por caja")
        st.bar_chart(f.groupby("Caja")["Total"].sum())
        st.subheader("Top 10 productos")
        top = f.groupby("Producto").agg(Unidades=("Cantidad", "sum"), Total=("Total", "sum")).sort_values(
            "Total", ascending=False).head(10)
        st.dataframe(top, use_container_width=True)
    # Margen estimado por venta (Precio_Actual vs Costo)
    if len(inv):
        costos = dict(zip(inv["SKU"].astype(str), inv["Costo"]))
        ff = f.copy()
        ff["Costo_U"] = ff["SKU"].map(costos).fillna(0)
        ff["Ganancia"] = (ff["Precio_Unitario"] - ff["Costo_U"]) * ff["Cantidad"]
        st.metric("💹 Ganancia bruta estimada del período",
                  f"${ff['Ganancia'].sum():,.0f}")
    st.download_button("⬇️ Descargar reporte del período (.xlsx)",
                       df_a_excel_bytes({"Resumen": f.groupby("Producto").agg(
                           Unidades=("Cantidad", "sum"), Total=("Total", "sum")).reset_index(),
                           "Detalle": f[COL_VENTAS]}),
                       file_name=f"reporte-MM-{d1}_{d2}.xlsx")

st.divider()
st.caption(f"Librería MM · datos en {DATA_DIR} · Precio sugerido = Costo × {MARGEN} · {ahora_str()}")
