import pandas as pd
from pathlib import Path
from openpyxl.styles import PatternFill, Font, Alignment
from openpyxl.utils import get_column_letter

BASE = Path(__file__).parent.parent

RAW_PATHS = {
    "Jumbo":     BASE / "data" / "raw" / "jumbo_raw.csv",
    "Farmacity": BASE / "data" / "raw" / "farmacity_raw.csv",
    "Disco":     BASE / "data" / "raw" / "disco_raw.csv",
}

OUTPUT = BASE / "data" / "planilla_shampoo.xlsx"

# Colores por fuente
COLORES = {
    "Jumbo":     {"header": "1A5276", "alt": "D6EAF8", "formula_h": "1F618D", "formula_alt": "AED6F1"},
    "Farmacity": {"header": "1E8449", "alt": "D5F5E3", "formula_h": "196F3D", "formula_alt": "A9DFBF"},
    "Disco":     {"header": "7D6608", "alt": "FEF9E7", "formula_h": "9A7D0A", "formula_alt": "F9E79F"},
}

# Diccionario de datos
DICCIONARIO = [
    ("fuente",         "Texto",    "PK parcial", "Origen del dato: Jumbo, Farmacity o Disco"),
    ("producto_id",    "Entero",   "FK",         "ID interno del producto en el sistema VTEX del retailer"),
    ("sku_id",         "Entero",   "PK",         "ID del SKU (variante única del producto) — clave primaria"),
    ("nombre",         "Texto",    "-",          "Nombre completo del producto tal como aparece en el sitio"),
    ("marca",          "Texto",    "-",          "Marca del producto (puede tener inconsistencias de case en raw)"),
    ("precio_ars",     "Decimal",  "-",          "Precio de venta en pesos argentinos (ARS)"),
    ("precio_lista",   "Decimal",  "-",          "Precio de lista original antes de descuentos (ARS)"),
    ("disponible",     "Booleano", "-",          "TRUE = disponible para compra al momento del scraping"),
    ("volumen_ml",     "Decimal",  "-",          "Volumen en ml extraído del nombre con regex. NULL si no se encontró"),
    ("precio_por_ml",  "Decimal",  "-",          "precio_ars / volumen_ml. NULL si no hay volumen"),
    ("categoria_raw",  "Texto",    "-",          "Categoría jerárquica según el árbol del retailer"),
    ("tipo_producto",  "Texto",    "-",          "Clasificado por Python: shampoo / acondicionador / tratamiento"),
    ("linea_tipo",     "Texto",    "-",          "Clasificado por Python: profesional / estandar"),
    ("tipo_cabello",   "Texto",    "-",          "Clasificado por Python: general / rizado / tratado / anticaspa / seco_danado / bebe / graso"),
    ("url",            "Texto",    "-",          "URL del producto en el sitio del retailer"),
]

# Columnas de fórmulas Excel a agregar en hojas limpias
# Cada entrada: (header, formula_template)
# {row} se reemplaza por el número de fila real
# Las letras de columna corresponden al raw (A=fuente...O=url)
FORMULAS_LIMPIEZA = [
    (
        "nombre_limpio",
        "=TRIM(CLEAN(PROPER(D{row})))",
        "TRIM+CLEAN+PROPER sobre 'nombre': elimina espacios, caracteres no imprimibles y capitaliza"
    ),
    (
        "marca_normalizada",
        "=PROPER(TRIM(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(E{row},\"TRESEMME\",\"Tresemme\"),\"HEAD & SHOULDERS\",\"Head & Shoulders\"),\"ELVIVE\",\"Elvive\"),\"FRUCTIS\",\"Fructis\")))",
        "PROPER+TRIM+SUBSTITUTE sobre 'marca': normaliza case y corrige nombres especificos"
    ),
    (
        "es_duplicado",
        "=COUNTIF($D$2:$D$10000,D{row})>1",
        "COUNTIF sobre 'nombre': detecta si hay mas de 1 fila con el mismo nombre (posible duplicado)"
    ),
    (
        "disponible_texto",
        "=IF(H{row}=TRUE,\"Disponible\",\"No disponible\")",
        "IF sobre 'disponible': convierte booleano a texto legible"
    ),
    (
        "precio_ml_calculado",
        "=IF(AND(ISNUMBER(I{row}),I{row}>0),ROUND(F{row}/I{row},2),\"Sin volumen\")",
        "IF+AND+ISNUMBER+ROUND: calcula precio/ml solo cuando el volumen es un numero valido"
    ),
    (
        "rango_precio",
        "=IF(F{row}<5000,\"Bajo\",IF(F{row}<15000,\"Medio\",\"Alto\"))",
        "IF anidado sobre 'precio_ars': segmenta en rangos Bajo/Medio/Alto"
    ),
]


def set_header_style(cell, hex_color):
    cell.fill = PatternFill("solid", fgColor=hex_color)
    cell.font = Font(color="FFFFFF", bold=True, size=10)
    cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)


def autofit(ws, max_w=55):
    for col in ws.columns:
        col_letter = get_column_letter(col[0].column)
        max_len = max((len(str(c.value)) if c.value else 0) for c in col)
        ws.column_dimensions[col_letter].width = min(max_len + 3, max_w)


def escribir_raw(writer, df, nombre_hoja, colores):
    df.to_excel(writer, sheet_name=nombre_hoja, index=False)
    ws = writer.sheets[nombre_hoja]

    # Header
    for cell in ws[1]:
        set_header_style(cell, colores["header"])
    ws.row_dimensions[1].height = 26

    # Filas alternas
    fill = PatternFill("solid", fgColor=colores["alt"])
    for i, row in enumerate(ws.iter_rows(min_row=2, max_row=ws.max_row)):
        if i % 2 == 0:
            for cell in row:
                cell.fill = fill

    ws.freeze_panes = "A2"
    autofit(ws)


def escribir_limpio(writer, df_raw, nombre_hoja, colores):
    """
    Hoja de limpieza: raw data (solo disponibles) + columnas con fórmulas Excel.
    Metodología de clase: dejar original intacto y trabajar en hoja paralela.
    """
    # Solo disponibles (equivale a aplicar Filtro en Excel por disponible=TRUE)
    df = df_raw[df_raw["disponible"] == True].copy().reset_index(drop=True)

    df.to_excel(writer, sheet_name=nombre_hoja, index=False)
    ws = writer.sheets[nombre_hoja]

    n_cols_raw = len(df.columns)  # columnas del raw (hasta col O = 15)

    # ── Headers raw ──────────────────────────────────────────────────
    for cell in ws[1]:
        set_header_style(cell, colores["header"])

    # ── Headers de fórmulas (columnas extras) ────────────────────────
    formula_fill_h = PatternFill("solid", fgColor=colores["formula_h"])
    for i, (header, _, _) in enumerate(FORMULAS_LIMPIEZA):
        col_idx = n_cols_raw + 1 + i          # 1-based
        col_letter = get_column_letter(col_idx)
        cell = ws.cell(row=1, column=col_idx, value=header)
        set_header_style(cell, colores["formula_h"])

    ws.row_dimensions[1].height = 26

    # ── Filas de datos con fórmulas ──────────────────────────────────
    alt_fill     = PatternFill("solid", fgColor=colores["alt"])
    formula_fill = PatternFill("solid", fgColor=colores["formula_alt"])

    for i, row_obj in enumerate(ws.iter_rows(min_row=2, max_row=ws.max_row)):
        row_num = i + 2
        # Color alterno en columnas raw
        if i % 2 == 0:
            for cell in row_obj:
                cell.fill = alt_fill

        # Insertar fórmulas en columnas extras
        for j, (_, formula_tpl, _) in enumerate(FORMULAS_LIMPIEZA):
            col_idx = n_cols_raw + 1 + j
            formula = formula_tpl.replace("{row}", str(row_num))
            cell = ws.cell(row=row_num, column=col_idx, value=formula)
            if i % 2 == 0:
                cell.fill = formula_fill

    ws.freeze_panes = "A2"
    autofit(ws)


def escribir_diccionario(writer, colores_base="1F3864"):
    rows = []
    for campo, tipo, clave, desc in DICCIONARIO:
        rows.append({"Campo": campo, "Tipo de dato": tipo, "Clave": clave, "Descripción": desc})
    df = pd.DataFrame(rows)
    df.to_excel(writer, sheet_name="diccionario_datos", index=False)
    ws = writer.sheets["diccionario_datos"]
    for cell in ws[1]:
        set_header_style(cell, colores_base)
    ws.row_dimensions[1].height = 26
    autofit(ws, max_w=80)
    ws.column_dimensions["D"].width = 72


def escribir_guia_formulas(writer):
    """Hoja extra: guía de qué fórmula hace qué (útil para el informe)."""
    rows = []
    for header, formula_tpl, desc in FORMULAS_LIMPIEZA:
        ejemplo = formula_tpl.replace("{row}", "2")
        rows.append({"Columna generada": header, "Fórmula (fila 2)": ejemplo, "Qué hace": desc})
    df = pd.DataFrame(rows)
    df.to_excel(writer, sheet_name="guia_formulas_limpieza", index=False)
    ws = writer.sheets["guia_formulas_limpieza"]
    for cell in ws[1]:
        set_header_style(cell, "6C3483")
    ws.row_dimensions[1].height = 26
    autofit(ws, max_w=90)
    ws.column_dimensions["B"].width = 75
    ws.column_dimensions["C"].width = 75


def main():
    with pd.ExcelWriter(OUTPUT, engine="openpyxl") as writer:
        for fuente, path in RAW_PATHS.items():
            df = pd.read_csv(path)
            colores = COLORES[fuente]

            # Hoja raw
            nombre_raw = f"raw_{fuente}"
            escribir_raw(writer, df, nombre_raw, colores)
            print(f"  {nombre_raw}: {len(df)} filas")

            # Hoja limpieza con fórmulas
            nombre_limpio = f"limpio_{fuente}"
            disponibles = df["disponible"].sum()
            escribir_limpio(writer, df, nombre_limpio, colores)
            print(f"  {nombre_limpio}: {disponibles} filas disponibles + {len(FORMULAS_LIMPIEZA)} columnas con fórmulas")

        escribir_diccionario(writer)
        escribir_guia_formulas(writer)
        print(f"  diccionario_datos + guia_formulas_limpieza")

    print(f"\nArchivo generado: {OUTPUT}")
    print(f"Hojas: raw_Jumbo, limpio_Jumbo, raw_Farmacity, limpio_Farmacity, raw_Disco, limpio_Disco, diccionario_datos, guia_formulas_limpieza")


if __name__ == "__main__":
    main()
