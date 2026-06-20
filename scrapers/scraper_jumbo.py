import sys
import pandas as pd
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))
from scraper_base import SEARCH_TERMS, fetch_vtex, vtex_a_filas

BASE_URL = "https://www.jumbo.com.ar/api/catalog_system/pub/products/search/"
OUTPUT = Path(__file__).parent.parent / "data" / "raw" / "jumbo_raw.csv"


def main():
    print("=== Scraping Jumbo ===")
    todos = {}
    for term in SEARCH_TERMS:
        print(f"\nBuscando: '{term}'")
        items = fetch_vtex(BASE_URL, term, max_pages=15)
        for it in items:
            todos[it["productId"]] = it

    print(f"\nProductos únicos Jumbo: {len(todos)}")
    filas = vtex_a_filas(list(todos.values()), "Jumbo")

    # Deduplicar por sku_id
    df = pd.DataFrame(filas).drop_duplicates(subset="sku_id")
    OUTPUT.parent.mkdir(parents=True, exist_ok=True)
    df.to_csv(OUTPUT, index=False, encoding="utf-8-sig")
    print(f"Guardado: {OUTPUT} ({len(df)} filas)")
    return df


if __name__ == "__main__":
    main()
