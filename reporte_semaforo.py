import pandas as pd
import os
#import webbrowser
from datetime import datetime
from pathlib import Path


def reparar_mojibake(valor):
    if not isinstance(valor, str):
        return valor

    texto = valor.strip()
    if not texto:
        return texto

    if any(marca in texto for marca in ("Ã", "Â", "â", "ð")):
        try:
            return texto.encode("latin1").decode("utf-8")
        except UnicodeError:
            return texto

    return texto


def normalizar_texto_df(df_entrada):
    columnas_texto = df_entrada.select_dtypes(include=["object", "string"]).columns
    if len(columnas_texto) == 0:
        return df_entrada

    df_entrada[columnas_texto] = df_entrada[columnas_texto].apply(
        lambda serie: serie.map(reparar_mojibake)
    )
    return df_entrada


usuario = os.getlogin()
fecha_actualizacion = datetime.now().strftime("%d/%m/%Y %H:%M")

# =========================
# RUTA DINÁMICA
# =========================
posibles_carpetas = [
    Path(fr"C:\Users\{usuario}\IMSS-BIENESTAR\División de Procesamiento de información - Comando Florence Nightingale\Proyectos\74 Limpieza de bases de abasto\Data"),
    Path(fr"C:\Users\{usuario}\OneDrive - IMSS-BIENESTAR\División de Procesamiento de información - Comando Florence Nightingale\Proyectos\74 Limpieza de bases de abasto\Data")
]

carpeta = next((p for p in posibles_carpetas if p.exists()), None)
if carpeta is None:
    raise FileNotFoundError("No se encontró ninguna carpeta válida")

archivos = list(carpeta.glob("reporte_metas_y_flags_*.xlsx"))
if not archivos:
    raise FileNotFoundError("No se encontraron archivos")

ruta = str(max(archivos, key=lambda x: x.stat().st_mtime))
print(f"Usando archivo: {ruta}")

# =========================
# CATÁLOGO
# =========================
#clues_catalogo = pd.read_parquet(
#    fr"C:\Users\{usuario}\IMSS-BIENESTAR\División de Procesamiento de información - Repositorio de Datos\CLUES\clues.parquet"
#)
posibles_clues = [
    Path(fr"C:\Users\{usuario}\IMSS-BIENESTAR\División de Procesamiento de información - Repositorio de Datos\CLUES\clues.parquet"),
    Path(fr"C:\Users\{usuario}\OneDrive - IMSS-BIENESTAR\División de Procesamiento de información - Repositorio de Datos\CLUES\clues.parquet")
]

ruta_clues = next((p for p in posibles_clues if p.exists()), None)

if ruta_clues is None:
    raise FileNotFoundError("No se encontró clues.parquet en ninguna ruta esperada")

clues_catalogo = pd.read_parquet(ruta_clues)
clues_catalogo = normalizar_texto_df(clues_catalogo)


catalogo_limpio = clues_catalogo.drop_duplicates(subset="clues_imb")

col_entidad_catalogo = next(
    (c for c in catalogo_limpio.columns if "entidad" in c.lower()),
    None
)

# CLUES a excluir completamente del reporte.
claves_excluidas = {"GRIMB000012"}

# =========================
# TABLA PRINCIPAL
# =========================
df = pd.read_excel(ruta, sheet_name="Tabla_entidad_flags")
df = normalizar_texto_df(df)

if claves_excluidas and col_entidad_catalogo:
    df["entidad"] = df["entidad"].astype(str).str.strip()

    clues_excluidas = pd.read_excel(ruta, sheet_name="Tabla_clues_flags")
    clues_excluidas = normalizar_texto_df(clues_excluidas)
    clues_excluidas = clues_excluidas[clues_excluidas["clues_imb"].isin(claves_excluidas)]

    if not clues_excluidas.empty:
        clues_excluidas = clues_excluidas.merge(
            catalogo_limpio[["clues_imb", col_entidad_catalogo]],
            on="clues_imb",
            how="left",
            validate="m:1"
        )
        clues_excluidas[col_entidad_catalogo] = clues_excluidas[col_entidad_catalogo].astype(str).str.strip()

        clues_excluidas["clues_con_inventario"] = (
            clues_excluidas[
                [
                    "reporto_medicamentos_010_040",
                    "reporto_material_curacion_060",
                    "reporto_otros_030_070_080"
                ]
            ].sum(axis=1) > 0
        ).astype(int)

        ajuste = (
            clues_excluidas
            .groupby(col_entidad_catalogo, dropna=False)
            .agg(
                meta_de_clues_ajuste=("clues_imb", "size"),
                clues_con_inventario_ajuste=("clues_con_inventario", "sum"),
                clues_medicamentos_010_040_ajuste=("reporto_medicamentos_010_040", "sum"),
                clues_material_curacion_060_ajuste=("reporto_material_curacion_060", "sum")
            )
            .reset_index()
        )

        df = df.merge(ajuste, left_on="entidad", right_on=col_entidad_catalogo, how="left")

        cols_ajustables = [
            "meta_de_clues",
            "clues_con_inventario",
            "clues_medicamentos_010_040",
            "clues_material_curacion_060"
        ]
        for col in cols_ajustables:
            col_ajuste = f"{col}_ajuste"
            df[col] = (df[col] - df[col_ajuste].fillna(0)).clip(lower=0)

        if col_entidad_catalogo != "entidad":
            df = df.drop(columns=[col_entidad_catalogo], errors="ignore")

        df = df.drop(columns=[
            "meta_de_clues_ajuste",
            "clues_con_inventario_ajuste",
            "clues_medicamentos_010_040_ajuste",
            "clues_material_curacion_060_ajuste"
        ], errors="ignore")

df["pct_avance"] = (
    (df["clues_con_inventario"] / df["meta_de_clues"]).where(df["meta_de_clues"] != 0, 0)
    .mul(100)
    .round(2)
)

df["pct_completo"] = (
    (
        (df["clues_medicamentos_010_040"] + df["clues_material_curacion_060"])
        /
        (df["meta_de_clues"] * 2)
    )
    .where(df["meta_de_clues"] != 0, 0)
    .mul(100)
    .round(1)
)

metas = df.assign(
    inventario_completo=(
        (
            df["clues_medicamentos_010_040"] +
            df["clues_material_curacion_060"]
        )
        /
        (df["meta_de_clues"] * 2)
    ).where(df["meta_de_clues"] != 0, 0)
     .mul(100)
     .round(1)
)

metas.columns = metas.columns.str.replace("_", " ", regex=False)

cols_color = ["pct avance", "inventario completo"]
metas[cols_color] = metas[cols_color].astype(float).round(2)

# =========================
# CLUES
# =========================
clues = pd.read_excel(ruta, sheet_name="Tabla_clues_flags")
clues = normalizar_texto_df(clues)
clues = clues.drop(columns=["nombre_comercial"], errors="ignore")

clues = clues[~clues["clues_imb"].isin(claves_excluidas)]

cols_catalogo = ["clues_imb", "nombre_de_la_unidad"]
if col_entidad_catalogo:
    cols_catalogo.append(col_entidad_catalogo)

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
# FLAGS (TU LÓGICA ORIGINAL SE RESPETA)
# =========================
cols_flags = [
    "reporto medicamentos 010 040",
    "reporto material curacion 060"
]

clues["conteo"] = clues[cols_flags].sum(axis=1)

# =========================
# NUEVO: SUMA DE REPORTES
# =========================
cols_reportes = [
    "reporto medicamentos 010 040",
    "reporto material curacion 060",
    "reporto otros 030 070 080"
]

clues["suma_reportes"] = clues[cols_reportes].sum(axis=1)

# =========================
# DETECTAR ENTIDAD
# =========================
col_entidad = next(
    (c for c in clues.columns if "entidad" in c),
    None
)

columnas_salida = ["clues imb", "nombre de la unidad", "conteo"]

if col_entidad:
    columnas_salida.insert(2, col_entidad)

# =========================
# SEGMENTACIÓN CORREGIDA
# =========================

# NO REPORTARON (nada en absoluto)
no_reportaron = clues[
    clues["suma_reportes"] == 0
][columnas_salida]

# INCOMPLETOS (hay reporte pero con brecha real)
incompletos = clues[
    (clues["suma_reportes"] > 0) &
    (clues["conteo"] > 0) &
    (clues["conteo"] < 2)
][columnas_salida]

# ordenar
no_reportaron = no_reportaron.sort_values(["conteo", "clues imb"])
incompletos = incompletos.sort_values(["conteo", "clues imb"])

# =========================
# FUNCIONES HTML
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
descargas = os.path.join(os.path.expanduser("~"), "Downloads\semaforo_entidades-\semaforo_entidades-")
ruta_html = os.path.join(descargas, "index.html")

with open(ruta_html, "w", encoding="utf-8") as f:
    f.write(html)

print(f"Reporte generado en: {ruta_html}")

#webbrowser.open("file://" + os.path.realpath(ruta_html))