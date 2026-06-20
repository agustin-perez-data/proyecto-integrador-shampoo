import pandas as pd
from collections import OrderedDict
from pathlib import Path
from openpyxl.styles import PatternFill, Font, Alignment
from openpyxl.utils import get_column_letter

BASE = Path(__file__).parent.parent

CLEAN_PATH = BASE / "data" / "clean" / "dataset_limpio.csv"

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


def calc_stats(series):
    """Métricas de estadística descriptiva según clase 4 UADE."""
    s = series.dropna()
    if len(s) == 0:
        return OrderedDict()
    q1 = float(s.quantile(0.25))
    q3 = float(s.quantile(0.75))
    iri = round(q3 - q1, 2)
    lim_inf = round(q1 - 1.5 * iri, 2)
    lim_sup = round(q3 + 1.5 * iri, 2)
    try:
        moda = round(float(s.mode().iloc[0]), 2)
    except Exception:
        moda = None
    return OrderedDict([
        ("n",                       int(len(s))),
        ("Minimo",                  round(float(s.min()), 2)),
        ("Maximo",                  round(float(s.max()), 2)),
        ("Rango",                   round(float(s.max() - s.min()), 2)),
        ("Media",                   round(float(s.mean()), 2)),
        ("Mediana",                 round(float(s.median()), 2)),
        ("Moda",                    moda),
        ("Varianza",                round(float(s.var()), 2)),
        ("Desvio Estandar",         round(float(s.std()), 2)),
        ("Asimetria",               round(float(s.skew()), 4)),
        ("Curtosis",                round(float(s.kurt()), 4)),
        ("Q1",                      round(q1, 2)),
        ("Q2 (Mediana)",            round(float(s.median()), 2)),
        ("Q3",                      round(q3, 2)),
        ("RI (Q3 - Q1)",            iri),
        ("Limite Inf. Outliers",    lim_inf),
        ("Limite Sup. Outliers",    lim_sup),
        ("N Outliers Inferiores",   int((s < lim_inf).sum())),
        ("N Outliers Superiores",   int((s > lim_sup).sum())),
    ])


# Fórmulas Sheets equivalentes para referencia del alumno
FORMULAS_REF = {
    "n":                      "COUNTA(rango)",
    "Minimo":                 "MIN(rango)",
    "Maximo":                 "MAX(rango)",
    "Rango":                  "MAX(rango)-MIN(rango)",
    "Media":                  "AVERAGE(rango)",
    "Mediana":                "MEDIAN(rango)",
    "Moda":                   "MODE.SNGL(rango)",
    "Varianza":               "VAR(rango)",
    "Desvio Estandar":        "STDEV(rango)",
    "Asimetria":              "SKEW(rango)",
    "Curtosis":               "KURT(rango)",
    "Q1":                     "QUARTILE(rango,1)",
    "Q2 (Mediana)":           "QUARTILE(rango,2)",
    "Q3":                     "QUARTILE(rango,3)",
    "RI (Q3 - Q1)":           "Q3-Q1",
    "Limite Inf. Outliers":   "Q1-1.5*RI",
    "Limite Sup. Outliers":   "Q3+1.5*RI",
    "N Outliers Inferiores":  "COUNTIF(rango,\"<\"&limite_inf)",
    "N Outliers Superiores":  "COUNTIF(rango,\">\"&limite_sup)",
}


def escribir_estadisticas(writer, df):
    """Hoja de estadística descriptiva — clase 4 UADE."""
    wb = writer.book
    ws = wb.create_sheet("estadistica_descriptiva")

    # ── Paleta de colores ───────────────────────────────────────────────
    C_TITLE   = "1C2833"  # encabezado principal
    C_GLOBAL  = "1A5276"  # secciones globales
    C_H1      = "6C3483"  # H1: linea_tipo
    C_H2      = "117A65"  # H2: tipo_cabello
    C_H3      = "784212"  # H3: fuente
    C_COL_HDR = "566573"  # cabeceras de columna
    C_ALT     = "EAF2FF"  # filas pares

    def hdr(cell, color, bold=True, size=10, align="left"):
        cell.fill = PatternFill("solid", fgColor=color)
        cell.font = Font(color="FFFFFF", bold=bold, size=size)
        cell.alignment = Alignment(horizontal=align, vertical="center", wrap_text=True)

    def val_cell(cell, alt=False, right=False, italic=False, gray=False):
        if alt:
            cell.fill = PatternFill("solid", fgColor=C_ALT)
        fc = "777777" if gray else "000000"
        cell.font = Font(size=9, italic=italic, color=fc)
        cell.alignment = Alignment(horizontal="right" if right else "left", vertical="center")

    current_row = [1]  # mutable para uso en closures

    def r():
        return current_row[0]

    def next_row(n=1):
        current_row[0] += n

    # ── Título principal ────────────────────────────────────────────────
    ws.merge_cells(start_row=r(), start_column=1, end_row=r(), end_column=7)
    c = ws.cell(row=r(), column=1,
                value="Estadistica Descriptiva — Dataset Shampoo & Acondicionador (Clase 4 UADE)")
    hdr(c, C_TITLE, size=12)
    ws.row_dimensions[r()].height = 26
    next_row(2)

    def write_global_section(titulo, col_color, variable, series, col_formula_note):
        """Tabla de 3 columnas: Metrica | Valor | Formula Sheets."""
        # Encabezado sección
        ws.merge_cells(start_row=r(), start_column=1, end_row=r(), end_column=3)
        hdr(ws.cell(row=r(), column=1, value=titulo), col_color, size=10)
        ws.row_dimensions[r()].height = 20
        next_row()

        # Sub-encabezados columna
        for col, txt in enumerate(["Metrica", "Valor", "Formula Sheets (referencia)"], 1):
            hdr(ws.cell(row=r(), column=col, value=txt), C_COL_HDR, size=9, align="center")
        ws.row_dimensions[r()].height = 17
        next_row()

        stats = calc_stats(series)
        for i, (metric, value) in enumerate(stats.items()):
            alt = (i % 2 == 0)
            val_cell(ws.cell(row=r(), column=1, value=metric), alt=alt)
            val_cell(ws.cell(row=r(), column=2, value=value),  alt=alt, right=True)
            val_cell(ws.cell(row=r(), column=3, value=FORMULAS_REF.get(metric, "")),
                     alt=alt, italic=True, gray=True)
            ws.row_dimensions[r()].height = 15
            next_row()

        # nota interpretación asimetría
        note = ws.cell(row=r(), column=1,
                       value=f"Nota: Asimetria={stats.get('Asimetria','?')}  "
                             f"({'cola derecha (precios altos)' if stats.get('Asimetria',0)>0 else 'cola izquierda'}). "
                             f"Outliers totales: {stats.get('N Outliers Inferiores',0)+stats.get('N Outliers Superiores',0)}")
        ws.merge_cells(start_row=r(), start_column=1, end_row=r(), end_column=3)
        note.font = Font(size=8, italic=True, color="444444")
        note.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
        ws.row_dimensions[r()].height = 22
        next_row(2)

    def write_grouped_section(titulo, col_color, variable, grupos_dict):
        """Tabla multi-columna: Metrica | Grupo1 | Grupo2 | ... con métricas clave."""
        grupos = list(grupos_dict.keys())
        n_grupos = len(grupos)
        total_cols = 1 + n_grupos  # Metrica + grupos

        ws.merge_cells(start_row=r(), start_column=1, end_row=r(), end_column=total_cols)
        hdr(ws.cell(row=r(), column=1, value=titulo), col_color, size=10)
        ws.row_dimensions[r()].height = 20
        next_row()

        # Cabeceras
        hdr(ws.cell(row=r(), column=1, value="Metrica"), C_COL_HDR, size=9, align="center")
        for j, g in enumerate(grupos, 2):
            hdr(ws.cell(row=r(), column=j, value=g), C_COL_HDR, size=9, align="center")
        ws.row_dimensions[r()].height = 17
        next_row()

        # Calcular stats por grupo
        all_stats = {g: calc_stats(grupos_dict[g]) for g in grupos}

        # Métricas seleccionadas para secciones agrupadas (subset)
        metricas_clave = [
            "n", "Media", "Mediana", "Desvio Estandar",
            "Minimo", "Maximo", "Rango",
            "Q1", "Q3", "RI (Q3 - Q1)",
            "Limite Inf. Outliers", "Limite Sup. Outliers",
            "N Outliers Inferiores", "N Outliers Superiores",
        ]

        for i, metric in enumerate(metricas_clave):
            alt = (i % 2 == 0)
            val_cell(ws.cell(row=r(), column=1, value=metric), alt=alt)
            for j, g in enumerate(grupos, 2):
                v = all_stats[g].get(metric, "")
                val_cell(ws.cell(row=r(), column=j, value=v), alt=alt, right=True)
            ws.row_dimensions[r()].height = 15
            next_row()

        next_row()  # espacio entre secciones

    # ── Sección 1: precio_ars global ─────────────────────────────────
    write_global_section(
        titulo=f"1. precio_ars — Analisis Global  (n={len(df):,} registros)",
        col_color=C_GLOBAL,
        variable="precio_ars",
        series=df["precio_ars"],
        col_formula_note="F",
    )

    # ── Sección 2: precio_por_ml global ──────────────────────────────
    df_ml = df[df["precio_por_ml"].notna()]
    write_global_section(
        titulo=f"2. precio_por_ml — Analisis Global  (n={len(df_ml):,} con volumen informado)",
        col_color=C_GLOBAL,
        variable="precio_por_ml",
        series=df_ml["precio_por_ml"],
        col_formula_note="J",
    )

    # ── Sección 3: H1 — precio_por_ml por linea_tipo ─────────────────
    write_grouped_section(
        titulo="3. H1: precio_por_ml por linea_tipo  (Profesional vs Estandar)",
        col_color=C_H1,
        variable="precio_por_ml",
        grupos_dict={
            linea: df_ml[df_ml["linea_tipo"] == linea]["precio_por_ml"]
            for linea in sorted(df_ml["linea_tipo"].dropna().unique())
        },
    )

    # ── Sección 4: H2 — precio_por_ml por tipo_cabello ───────────────
    write_grouped_section(
        titulo="4. H2: precio_por_ml por tipo_cabello",
        col_color=C_H2,
        variable="precio_por_ml",
        grupos_dict={
            tc: df_ml[df_ml["tipo_cabello"] == tc]["precio_por_ml"]
            for tc in sorted(df_ml["tipo_cabello"].dropna().unique())
        },
    )

    # ── Sección 5: H3 — precio_ars por fuente ────────────────────────
    write_grouped_section(
        titulo="5. H3: precio_ars por fuente  (Jumbo vs Farmacity vs Disco)",
        col_color=C_H3,
        variable="precio_ars",
        grupos_dict={
            f: df[df["fuente"] == f]["precio_ars"]
            for f in ["Jumbo", "Farmacity", "Disco"]
        },
    )

    # ── Anchos de columna ─────────────────────────────────────────────
    ws.column_dimensions["A"].width = 26
    ws.column_dimensions["B"].width = 16
    ws.column_dimensions["C"].width = 34
    ws.column_dimensions["D"].width = 16
    ws.column_dimensions["E"].width = 16
    ws.column_dimensions["F"].width = 16
    ws.column_dimensions["G"].width = 16
    ws.column_dimensions["H"].width = 16
    ws.freeze_panes = "A3"


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

        df_clean = pd.read_csv(CLEAN_PATH)
        escribir_estadisticas(writer, df_clean)
        print(f"  estadistica_descriptiva: {len(df_clean)} registros analizados")
        print(f"  diccionario_datos + guia_formulas_limpieza")

    print(f"\nArchivo generado: {OUTPUT}")
    print(f"Hojas: raw_Jumbo, limpio_Jumbo, raw_Farmacity, limpio_Farmacity, raw_Disco, limpio_Disco, diccionario_datos, guia_formulas_limpieza")


if __name__ == "__main__":
    main()
