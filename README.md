# Librería MM — Sistema de Inventario, Ventas y Caja
# Stack: Python + Streamlit + Pandas + Excel local

Aplicación para migrar el sistema de Excel (`Ventas-Enero.xlsx`, `Deudores.xlsm`)
a una app con inventario, ventas con descuento automático de stock,
alertas de reposición, deudores y reportes exportables a `.xlsx`.

## 1. Estructura del proyecto

```
libreria-mm/
├── app.py              # Aplicación principal Streamlit
├── init_db.py          # Genera la base inicial en Excel + datos de ejemplo
├── requirements.txt    # Dependencias
├── README.md           # Este archivo
└── data/               # Se crea sola al ejecutar (no borrar)
    ├── inventario.xlsx # Hoja: Inventario
    ├── ventas.xlsx     # Hoja: Ventas
    └── deudores.xlsx   # Hoja: Deudores
```

### Esquema de datos

**`Inventario`** (`data/inventario.xlsx`)

| Columna | Tipo | Descripción |
|---|---|---|
| `SKU` | texto (único) | Código de barras / identificador |
| `Nombre` | texto | Descripción del libro / artículo / servicio |
| `Categoria` | texto | Libros, Útiles, Fotocopias, Servicios, etc. |
| `Costo` | número | Precio de compra al proveedor |
| `Precio_Sugerido` | número (auto) | `Costo * 1.5` — se recalcula solo |
| `Precio_Actual` | número (editable) | Precio final de venta (ofertas / redondeos) |
| `Stock_Actual` | entero | Cantidad disponible |
| `Stock_Minimo` | entero | Nivel crítico de reposición |
| `Fecha_Actualizacion` | fecha | Última modificación |

**`Ventas`** (`data/ventas.xlsx`)

| Columna | Descripción |
|---|---|
| `ID_Venta` | `V-YYYYMMDD-HHMMSS` (grupo/ticket) |
| `Fecha` | Fecha y hora de la venta |
| `SKU`, `Producto`, `Categoria` | Copia del ítem vendido |
| `Cantidad` | Unidades vendidas (descuenta stock) |
| `Precio_Unitario` | Se toma de `Precio_Actual` |
| `Total` | `Cantidad * Precio_Unitario` |
| `Metodo_Pago` | Efectivo, Tarjeta, Transferencia, Cuenta corriente |
| `Caja` | Caja 1, Caja 2, etc. |
| `Notas` | Opcional |

**`Deudores`** (`data/deudores.xlsx`, migrado desde `Deudores.xlsm`)

| Columna | Descripción |
|---|---|
| `ID` | `D-0001` correlativo |
| `Cliente`, `Telefono` | Datos del deudor |
| `Concepto` | Qué llevó / SKU relacionado |
| `Monto_Total`, `Monto_Pagado`, `Saldo` | `Saldo = Total - Pagado` (auto) |
| `Fecha_Prestamo`, `Fecha_Vencimiento` | Fechas |
| `Estado` | Pendiente / Parcial / Pagado / Vencido |

## 2. Instalación paso a paso (Windows)

### Paso 1 — Instalar Python 3.10 o superior
1. Ir a https://www.python.org/downloads/
2. Descargar Python 3.11 o 3.12.
3. **Importante:** en el instalador marcar ✅ `Add python.exe to PATH`.
4. Verificar en PowerShell:
   ```powershell
   python --version
   ```

### Paso 2 — Descargar este proyecto
Copiar la carpeta `libreria-mm` a, por ejemplo:
```
C:\LibreriaMM\
```

### Paso 3 — Crear entorno virtual (recomendado) e instalar dependencias
```powershell
cd C:\LibreriaMM
python -m venv venv
.\venv\Scripts\Activate.ps1
pip install --upgrade pip
pip install -r requirements.txt
```

> Si PowerShell bloquea la activación:
> ```powershell
> Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
> ```

### Paso 4 — Inicializar la base de datos local
```powershell
python init_db.py
```
Esto crea `data/inventario.xlsx`, `data/ventas.xlsx` y `data/deudores.xlsx`
con 12 productos de ejemplo, 8 ventas y 4 deudores para probar.
Si ya existen, **no los sobrescribe** (usa `--force` para regenerar):
```powershell
python init_db.py --force
```

### Paso 5 — Ejecutar la aplicación
```powershell
streamlit run app.py
```
Abrir en el navegador: http://localhost:8501

## 3. Uso diario

1. **Inventario:** agregar productos (SKU único), editar `Precio_Actual`,
   `Stock_Actual` y `Stock_Minimo` directo en la tabla. `Precio_Sugerido`
   se recalcula solo como `Costo * 1.5`.
2. **Ventas y Caja:** armar carrito por SKU/nombre → cobrar con método
   de pago y caja → el stock se descuenta automáticamente.
3. **Deudores:** Altas, abonos parciales y control de vencidos.
4. **Importar / Exportar:** subir `Ventas-Enero.xlsx` o `Deudores.xlsm`
   viejos (la app mapea columnas aunque tengan otros nombres) y descargar
   inventario y reportes actualizados en `.xlsx`.
5. **Reportes:** ventas diarias/mensuales, por categoría, método de pago,
   ticket promedio y top productos.

## 4. Respaldo

Copiar la carpeta `data/` a un pendrive o Drive al cierre de caja.
Recomendado: una copia por día con fecha, ej. `data-2026-10-10.zip`.
