import pandas as pd
from pathlib import Path
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

BASE = Path(__file__).parent.parent

RAW = {
    "raw_Jumbo":     BASE / "data" / "raw" / "jumbo_raw.csv",
    "raw_Farmacity": BASE / "data" / "raw" / "farmacity_raw.csv",
    "raw_Disco":     BASE / "data" / "raw" / "disco_raw.csv",
}

CLEAN_PATH = BASE / "data" / "clean" / "dataset_limpio.csv"
OUTPUT     = BASE / "data" / "planilla_shampoo.xlsx"

DICCIONARIO = [
    ("fuente",         "Texto",   "Origen del dato: Jumbo, Farmacity o Disco"),
    ("producto_id",    "Entero",  "ID interno del producto en el sistema VTEX del retailer"),
    ("sku_id",         "Entero",  "ID interno del SKU (variante específica del producto)"),
    ("nombre",         "Texto",   "Nombre completo del producto tal como aparece en el sitio"),
    ("marca",          "Texto",   "Marca del producto, normalizada a Title Case"),
    ("precio_ars",     "Decimal", "Precio de venta en pesos argentinos (ARS)"),
    ("precio_lista",   "Decimal", "Precio de lista original antes de descuentos (ARS)"),
    ("disponible",     "Booleano","True = disponible para compra; False = sin stock / discontinuado"),
    ("volumen_ml",     "Decimal", "Volumen del producto en mililitros, extraído del nombre con regex"),
    ("precio_por_ml",  "Decimal", "Precio por mililitro = precio_ars / volumen_ml"),
    ("categoria_raw",  "Texto",   "Categoría jerárquica del producto según el árbol del retailer"),
    ("tipo_producto",  "Texto",   "Clasificación: shampoo / acondicionador / tratamiento"),
    ("linea_tipo",     "Texto",   "Clasificación: profesional / estandar (por keywords en nombre/marca)"),
    ("tipo_cabello",   "Texto",   "Tipo de cabello objetivo: general / rizado / tratado / anticaspa / seco_danado / bebe / graso"),
    ("es_combo",       "Booleano","True si el producto es un pack o kit de varios artículos"),
    ("url",            "Texto",   "URL del producto en el sitio del retailer"),
]

HEADER_FILL  = PatternFill("solid", fgColor="1F3864")
HEADER_FONT  = Font(color="FFFFFF", bold=True, size=10)
ALT_FILL     = PatternFill("solid", fgColor="EEF2F8")
BORDER_SIDE  = Side(style="thin", color="CCCCCC")
CELL_BORDER  = Border(bottom=Border(bottom=BORDER_SIDE).bottom)

SECTION_COLORS = {
    "Jumbo":     "D6EAF8",
    "Farmacity": "D5F5E3",
    "Disco":     "FEF9E7",
}


def estilo_header(ws):
    for cell in ws[1]:
        cell.fill   = HEADER_FILL
        cell.font   = HEADER_FONT
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.row_dimensions[1].height = 28


def autofit(ws, max_width=50):
    for col in ws.columns:
        col_letter = get_column_letter(col[0].column)
        max_len = max((len(str(c.value)) if c.value else 0) for c in col)
        ws.column_dimensions[col_letter].width = min(max_len + 3, max_width)


def filas_alternas(ws, fill_color, start_row=2):
    fill = PatternFill("solid", fgColor=fill_color)
    for i, row in enumerate(ws.iter_rows(min_row=start_row, max_row=ws.max_row)):
        if i % 2 == 0:
            for cell in row:
                cell.fill = fill


def escribir_hoja(writer, df, nombre_hoja, fill_color):
    df.to_excel(writer, sheet_name=nombre_hoja, index=False)
    ws = writer.sheets[nombre_hoja]
    estilo_header(ws)
    filas_alternas(ws, fill_color)
    autofit(ws)
    ws.freeze_panes = "A2"


def main():
    df_clean = pd.read_csv(CLEAN_PATH, index_col="id")

    with pd.ExcelWriter(OUTPUT, engine="openpyxl") as writer:

        # --- Hojas raw ---
        for hoja, path in RAW.items():
            fuente = hoja.split("_")[1]
            df = pd.read_csv(path)
            color = SECTION_COLORS.get(fuente, "F5F5F5")
            escribir_hoja(writer, df, hoja, color)
            print(f"  {hoja}: {len(df)} filas")

        # --- Hojas limpias por fuente ---
        for fuente in ["Jumbo", "Farmacity", "Disco"]:
            df_f = df_clean[df_clean["fuente"] == fuente].copy()
            color = SECTION_COLORS.get(fuente, "F5F5F5")
            escribir_hoja(writer, df_f.reset_index(drop=True), f"limpio_{fuente}", color)
            print(f"  limpio_{fuente}: {len(df_f)} filas")

        # --- Diccionario de datos ---
        df_dic = pd.DataFrame(DICCIONARIO, columns=["Campo", "Tipo", "Descripción"])
        df_dic.to_excel(writer, sheet_name="diccionario_datos", index=False)
        ws = writer.sheets["diccionario_datos"]
        estilo_header(ws)
        autofit(ws, max_width=80)
        ws.column_dimensions["C"].width = 70
        print(f"  diccionario_datos: {len(df_dic)} campos")

    print(f"\nArchivo generado: {OUTPUT}")


if __name__ == "__main__":
    main()
