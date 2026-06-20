import sys
import pandas as pd
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))
from scraper_jumbo import main as scrape_jumbo
from scraper_farmacity import main as scrape_farmacity

OUTPUT = Path(__file__).parent.parent / "data" / "raw" / "dataset_combinado_raw.csv"


def main():
    df_jumbo = scrape_jumbo()
    df_farm = scrape_farmacity()

    df = pd.concat([df_jumbo, df_farm], ignore_index=True)
    df.to_csv(OUTPUT, index=False, encoding="utf-8-sig")

    print(f"\n{'='*50}")
    print(f"TOTAL registros combinados: {len(df)}")
    print(f"  - Jumbo:      {len(df_jumbo)}")
    print(f"  - Farmacity:  {len(df_farm)}")
    print(f"\nDistribución por tipo_producto:")
    print(df["tipo_producto"].value_counts().to_string())
    print(f"\nDistribución por linea_tipo:")
    print(df["linea_tipo"].value_counts().to_string())
    print(f"\nDistribución por tipo_cabello:")
    print(df["tipo_cabello"].value_counts().to_string())
    print(f"\nGuardado en: {OUTPUT}")


if __name__ == "__main__":
    main()
