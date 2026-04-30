import pandas as pd
import os
#import webbrowser
from datetime import datetime
from pathlib import Path

usuario = os.getlogin()
fecha_actualizacion = datetime.now().strftime("%d/%m/%Y %H:%M")

# =========================
# RUTA DINÁMICA
# =========================
carpeta = Path(fr"C:\Users\{usuario}\IMSS-BIENESTAR\División de Procesamiento de información - Comando Florence Nightingale\Proyectos\74 Limpieza de bases de abasto\Data")

archivos = list(carpeta.glob("reporte_metas_y_flags_*.xlsx"))
if not archivos:
    raise FileNotFoundError("No se encontraron archivos")

ruta = str(max(archivos, key=lambda x: x.stat().st_mtime))
print(f"Usando archivo: {ruta}")

# =========================
# CATÁLOGO
# =========================
clues_catalogo = pd.read_parquet(
    fr"C:\Users\{usuario}\IMSS-BIENESTAR\División de Procesamiento de información - Repositorio de Datos\CLUES\clues.parquet"
)

catalogo_limpio = clues_catalogo.drop_duplicates(subset="clues_imb")

# =========================
# TABLA PRINCIPAL
# =========================
df = pd.read_excel(ruta, sheet_name="Tabla_entidad_flags")

metas = df.assign(
    inventario_completo=(
        (df["clues_material_curacion_060"] / df["meta_de_clues"])
        .where(df["meta_de_clues"] != 0, 0) * 100
    ).round(1)
)

metas.columns = metas.columns.str.replace("_", " ", regex=False)

cols_color = ["pct avance", "inventario completo"]
metas[cols_color] = metas[cols_color].astype(float).round(2)

# =========================
# CLUES
# =========================
clues = pd.read_excel(ruta, sheet_name="Tabla_clues_flags")
clues = clues.drop(columns=["nombre_comercial"], errors="ignore")

# detectar columna entidad en catálogo
col_entidad_catalogo = next(
    (c for c in catalogo_limpio.columns if "entidad" in c.lower()),
    None
)

# columnas a usar en merge
cols_catalogo = ["clues_imb", "nombre_de_la_unidad"]
if col_entidad_catalogo:
    cols_catalogo.append(col_entidad_catalogo)

# merge
clues = clues.merge(
    catalogo_limpio[cols_catalogo],
    on="clues_imb",
    how="left",
    validate="m:1"
)

# =========================
# NORMALIZACIÓN
# =========================
clues.columns = (
    clues.columns
    .str.replace("_", " ", regex=False)
    .str.lower()
)

# =========================
# FLAGS
# =========================
cols_flags = [
    "reporto en cpm y ca",
    "reporto medicamentos 010 040",
    "reporto material curacion 060",
    "reporto otros 030 070 080"
]

clues["conteo"] = clues[cols_flags].sum(axis=1)

# =========================
# DETECTAR ENTIDAD YA LIMPIA
# =========================
col_entidad = next(
    (c for c in clues.columns if "entidad" in c),
    None
)

# columnas finales dinámicas
columnas_salida = ["clues imb", "nombre de la unidad", "conteo"]

if col_entidad:
    columnas_salida.insert(2, col_entidad)

# =========================
# SEGMENTACIÓN
# =========================
no_reportaron = clues[clues["conteo"] == 0][columnas_salida]
incompletos = clues[
    (clues["conteo"] > 0) & (clues["conteo"] < 4)
][columnas_salida]

# ordenar
no_reportaron = no_reportaron.sort_values(["conteo", "clues imb"])
incompletos = incompletos.sort_values(["conteo", "clues imb"])

# =========================
# FUNCIONES
# =========================
def semaforo(valor):
    return (
        "#D41111" if valor < 50 else
        "#F1D54A" if valor < 75 else
        "#88A91E" if valor < 100 else
        "#0D5D2A"
    )

def color_texto(bg):
    return "white" if bg in ["#D41111", "#0D5D2A"] else "black"

def tabla_principal_html(df):
    filas = ""
    for _, row in df.iterrows():
        fila = "<tr>"
        for col in df.columns:
            valor = row[col]
            if col in cols_color:
                color = semaforo(valor)
                txt = color_texto(color)
                display = f"{valor:.2f}"
            else:
                color = "white"
                txt = "black"
                display = valor

            fila += f'<td style="background:{color};color:{txt};text-align:center;">{display}</td>'
        fila += "</tr>"
        filas += fila
    return filas

def tabla_simple_html(df):
    filas = ""
    for _, row in df.iterrows():
        fila = "<tr>"
        for col in df.columns:
            fila += f"<td>{row[col]}</td>"
        fila += "</tr>"
        filas += fila
    return filas

tabla_principal = tabla_principal_html(metas)
tabla_no = tabla_simple_html(no_reportaron)
tabla_inc = tabla_simple_html(incompletos)

headers_principal = ''.join([f"<th>{col}</th>" for col in metas.columns])
headers_no = ''.join([f"<th>{col}</th>" for col in no_reportaron.columns])
headers_inc = ''.join([f"<th>{col}</th>" for col in incompletos.columns])

# =========================
# HTML
# =========================
html = f"""
<!DOCTYPE html>
<html>
<head>
<meta charset="UTF-8">
<title>Reporte Inventario</title>

<style>
body {{ font-family: Arial; background:#f4f6f9; margin:0; }}
.header {{ background:#7a1f2b; color:white; padding:15px; }}
.container {{ padding:20px; }}

table {{ border-collapse: collapse; width:100%; margin-bottom:20px; }}
th {{ background:#7a1f2b; color:white; padding:8px; }}
td {{ padding:6px; text-align:center; }}

.simple td {{ background:white; color:black; }}

.btn {{
    background:#7a1f2b;
    color:white;
    padding:8px;
    border:none;
    cursor:pointer;
}}

@media print {{
    @page {{ size: landscape; margin: 10mm; }}
    * {{
        -webkit-print-color-adjust: exact !important;
        print-color-adjust: exact !important;
    }}
    .btn {{ display:none; }}
}}
</style>

<script>
function imprimirPDF() {{
    window.print();
}}
</script>

</head>

<body>

<div class="header">
<h1>Reporte de Inventario</h1>
</div>

<div class="container">

<button class="btn" onclick="imprimirPDF()">Descargar PDF</button>
<p>Actualización: {fecha_actualizacion}</p>

<h2>Vista General</h2>
<table>
<thead><tr>{headers_principal}</tr></thead>
<tbody>{tabla_principal}</tbody>
</table>

<h2>CLUES que NO reportaron ({len(no_reportaron)})</h2>
<table class="simple">
<thead><tr>{headers_no}</tr></thead>
<tbody>{tabla_no}</tbody>
</table>

<h2>CLUES incompletos ({len(incompletos)})</h2>
<table class="simple">
<thead><tr>{headers_inc}</tr></thead>
<tbody>{tabla_inc}</tbody>
</table>

</div>
</body>
</html>
"""

# =========================
# GUARDAR Y ABRIR
# =========================
descargas = os.path.join(os.path.expanduser("~"), "Downloads")
ruta_html = os.path.join(descargas, "reporte_inventario.html")

with open(ruta_html, "w", encoding="utf-8") as f:
    f.write(html)

print(f"Reporte generado en: {ruta_html}")

#webbrowser.open("file://" + os.path.realpath(ruta_html))