import pandas as pd
# import webbrowser
import os
usuario = os.getlogin()
from datetime import datetime

fecha_actualizacion = datetime.now().strftime("%d/%m/%Y %H:%M")
ruta = fr"C:\Users\{usuario}\IMSS-BIENESTAR\División de Procesamiento de información - Comando Florence Nightingale\Proyectos\74 Limpieza de bases de abasto\Data\reporte_metas_y_flags_2026-04-24.xlsx"

df = pd.read_excel(ruta, sheet_name="Tabla_entidad_flags")
metas = df.assign(
    inventario_completo = (
        (df["clues_material_curacion_060"] / df["meta_de_clues"])
        .where(df["meta_de_clues"] != 0, 0)
        * 100
    ).round(1)
)
# quitamos los _ para mejor visualizacion en el mapa 
metas.columns = metas.columns.str.replace("_", " ", regex=False)
# mapeo para nombres largos de las entidades, pa que se vea chidori 
mapeo_entidades = {
    "MICHOACAN DE OCAMPO": "MICHOACAN",
    "VERACRUZ DE IGNACIO DE LA LLAVE": "VERACRUZ"
}

metas["entidad"] = metas["entidad"].replace(mapeo_entidades)

cols_color = ["pct avance", "inventario completo"]

for col in cols_color:
    metas[col] = metas[col].astype(float).round(2)
def semaforo(valor):
    return (
        "#D41111" if valor < 50 else
        "#F1D54A" if valor < 75 else
        "#88A91E" if valor < 100 else
        "#0D5D2A"
    )
def color_texto(bg):
    return "white" if bg in ["#D41111", "#0D5D2A"] else "black"
# html filas
filas_html = ""

for _, row in metas.iterrows():
    fila = "<tr>"
    
    for col in metas.columns:
        valor = row[col]
        
        if col in cols_color:
            color = semaforo(valor)
            txt_color = color_texto(color)
            display = f"{valor:.2f}"
        else:
            color = "white"
            txt_color = "black"
            display = valor
        
        fila += f'''
        <td style="
            background-color:{color};
            color:{txt_color};
            text-align:center;
            font-weight:500;
        ">
            {display}
        </td>
        '''
    
    fila += "</tr>"
    filas_html += fila
html = f"""
<!DOCTYPE html>
<html>
<head>
<meta charset="UTF-8">
<title>Reporte Inventario</title>

<style>
body {{
    font-family: 'Segoe UI', Arial;
    margin: 0;
    background-color: #f4f6f9;
}}

.header {{
    background: #7a1f2b;
    color: white;
    padding: 15px 25px;
}}

.header h1 {{
    margin: 0;
    font-size: 22px;
}}

.subheader {{
    font-size: 12px;
    opacity: 0.8;
}}

.container {{
    padding: 20px;
}}

.tabla-container {{
    background: white;
    padding: 15px;
    border-radius: 10px;
    box-shadow: 0 2px 6px rgba(0,0,0,0.1);
    overflow-x: auto;
}}

table {{
    border-collapse: collapse;
    width: 100%;
    font-size: 13px;
}}

th {{
    background-color: #7a1f2b;
    color: white;
    padding: 8px;
}}

td {{
    padding: 6px;
    text-align: center;
}}

tr:hover {{
    background-color: #f1f1f1;
}}

.update {{
    text-align: right;
    font-size: 12px;
    color: #777;
    margin-bottom: 10px;
}}

.btn {{
    background-color: #7a1f2b;
    color: white;
    border: none;
    padding: 8px 14px;
    border-radius: 6px;
    cursor: pointer;
    margin-bottom: 10px;
}}

.btn:hover {{
    background-color: #5a1720;
}}

/*  MODO IMPRESIÓN */
@media print {{

    @page {{
        size: landscape;
        margin: 10mm;
    }}

    /*  fuerza colores */
    * {{
        -webkit-print-color-adjust: exact !important;
        print-color-adjust: exact !important;
    }}

    body * {{
        visibility: hidden;
    }}

    .tabla-container, .tabla-container * {{
        visibility: visible;
    }}

    .tabla-container {{
        position: absolute;
        left: 0;
        top: 0;
        width: 100%;
        box-shadow: none;
        background: white;
    }}

    td {{
        -webkit-print-color-adjust: exact;
        print-color-adjust: exact;
    }}

    .btn {{
        display: none;
    }}
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
    <h1>Reporte de Inventario Hospitales por Entidad</h1>
    <div class="subheader">IMSS-BIENESTAR · Monitoreo Operativo</div>
</div>

<div class="container">

    <div class="tabla-container">

        <button class="btn" onclick="imprimirPDF()">Descargar PDF</button>

        <div class="update">
            Última actualización: {fecha_actualizacion}
        </div>

        <table>
            <thead>
                <tr>
                    {''.join([f"<th>{col}</th>" for col in metas.columns])}
                </tr>
            </thead>
            <tbody>
                {filas_html}
            </tbody>
        </table>

    </div>

</div>

</body>
</html>
"""

descargas = os.path.join(os.path.expanduser("~"), f"C:\\Users\\{usuario}\\Downloads\\semaforo_entidades-")

ruta_html = os.path.join(descargas, "index.html")

with open(ruta_html, "w", encoding="utf-8") as f:
    f.write(html)

# webbrowser.open("file://" + os.path.realpath(ruta_html))