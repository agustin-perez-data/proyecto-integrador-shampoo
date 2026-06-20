import re
import pandas as pd
from pathlib import Path

INPUT  = Path(__file__).parent.parent / "data" / "raw" / "dataset_combinado_raw.csv"
OUTPUT = Path(__file__).parent.parent / "data" / "clean" / "dataset_limpio.csv"

# Palabras que indican producto NO capilar
KEYWORDS_EXCLUIR = [
    "para pisos", "suavizante ropa", "suavizante tela", "lavandina",
    "frisbee", "juguete perro", "hueso perro", "pets fun",
    "jabon liquido bebe", "jabón líquido bebe", "aceite bebe", "aceite bebé",
    "crema hidratante", "colonia", "perfume", "desodorante",
    "mascara de pestañas", "máscaras de pestañas", "rimmel",
    "chips", "llanofertil", "fertilizante",
    "crema facial", "gel de limpieza facial", "emulsion facial",
]

# Normalización de marcas
MARCA_MAP = {
    "PANTENE": "Pantene",
    "ELVIVE": "Elvive",
    "SEDAL": "Sedal",
    "DOVE": "Dove",
    "FRUCTIS": "Fructis",
    "HEAD & SHOULDERS": "Head & Shoulders",
    "HEAD&SHOULDERS": "Head & Shoulders",
    "TRESEMMÉ": "Tresemmé",
    "TRESEMME": "Tresemmé",
    "HERBAL ESSENCES": "Herbal Essences",
    "PLUSBELLE": "Plusbelle",
    "CAPILATIS": "Capilatis",
    "WELLA": "Wella",
    "KERASTASE": "Kérastase",
    "KÉRASTASE": "Kérastase",
    "JOHNSON": "Johnson's",
    "JOHNSON'S": "Johnson's",
    "JOHNSONS": "Johnson's",
    "SUNSILK": "Sunsilk",
    "GARNIER": "Garnier",
}


def normalizar_marca(marca: str) -> str:
    upper = marca.strip().upper()
    for key, val in MARCA_MAP.items():
        if key in upper:
            return val
    return marca.strip().title()


def es_no_capilar(nombre: str) -> bool:
    texto = nombre.lower()
    return any(kw in texto for kw in KEYWORDS_EXCLUIR)


def main():
    df = pd.read_csv(INPUT)
    total_inicial = len(df)
    log = []

    # 1. Solo disponibles
    df = df[df["disponible"] == True].copy()
    log.append(f"Solo disponibles: {len(df)} (eliminados: {total_inicial - len(df)})")

    # 2. Precio mínimo razonable
    df = df[df["precio_ars"] >= 500].copy()
    log.append(f"Precio >= $500: {len(df)}")

    # 3. Eliminar productos no capilares
    antes = len(df)
    df = df[~df["nombre"].apply(es_no_capilar)].copy()
    log.append(f"Sin no capilares: {len(df)} (eliminados: {antes - len(df)})")

    # 4. Volumen ml inválido → NULL
    antes_vol = df["volumen_ml"].notna().sum()
    df.loc[df["volumen_ml"] > 5000, "volumen_ml"] = None
    df.loc[df["volumen_ml"] > 5000, "precio_por_ml"] = None
    # Recalcular precio_por_ml donde volumen es válido
    mask = df["volumen_ml"].notna() & (df["volumen_ml"] > 0)
    df.loc[mask, "precio_por_ml"] = (df.loc[mask, "precio_ars"] / df.loc[mask, "volumen_ml"]).round(4)
    log.append(f"Volumen ml corregido: {df['volumen_ml'].notna().sum()} con valor")

    # 5. Eliminar duplicados exactos (nombre + fuente)
    antes = len(df)
    df = df.drop_duplicates(subset=["nombre", "fuente"]).copy()
    log.append(f"Sin duplicados nombre+fuente: {len(df)} (eliminados: {antes - len(df)})")

    # 6. Normalizar marca
    df["marca"] = df["marca"].apply(normalizar_marca)

    # 7. Flag combo/kit
    df["es_combo"] = df["nombre"].str.lower().str.contains("combo|kit|pack x|set ").fillna(False)

    # 8. Limpiar encoding raro en nombre
    df["nombre"] = df["nombre"].str.encode("utf-8", errors="replace").str.decode("utf-8")

    # 9. Reset index
    df = df.reset_index(drop=True)
    df.index += 1
    df.index.name = "id"

    OUTPUT.parent.mkdir(parents=True, exist_ok=True)
    df.to_csv(OUTPUT, encoding="utf-8-sig")

    print("=== LIMPIEZA COMPLETADA ===")
    for paso in log:
        print(" •", paso)
    print(f"\nDataset final: {len(df)} registros")
    print(f"\nPor fuente:")
    print(df["fuente"].value_counts().to_string())
    print(f"\nPor tipo_producto:")
    print(df["tipo_producto"].value_counts().to_string())
    print(f"\nPor linea_tipo:")
    print(df["linea_tipo"].value_counts().to_string())
    print(f"\nPor tipo_cabello:")
    print(df["tipo_cabello"].value_counts().to_string())
    print(f"\nPrecio ARS (disponibles):")
    print(df["precio_ars"].describe().round(2).to_string())
    print(f"\nPrecio/ml (registros con volumen):")
    print(df["precio_por_ml"].dropna().describe().round(4).to_string())
    print(f"\nGuardado: {OUTPUT}")

    return df


if __name__ == "__main__":
    main()
