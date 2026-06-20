import re
import time
import requests
import pandas as pd


HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/125.0.0.0 Safari/537.36"
    ),
    "Accept": "application/json",
}

SEARCH_TERMS = [
    "shampoo",
    "acondicionador",
    "shampoo rulos",
    "shampoo anticaspa",
    "shampoo pelo seco",
    "shampoo profesional",
    "enjuague capilar",
    "tratamiento capilar",
    "shampoo bebe",
    "shampoo keratina",
    "dove cabello",
    "pantene",
    "sedal",
    "head shoulders",
    "fructis",
    "tresemme",
    "herbal essences",
    "elvive",
]

ML_VOLUME_RE = re.compile(
    r"(\d[\d.,]*)\s*(ml|mililitros?|l\b|litros?|gr?\b|gramos?|kg\b|kilos?)",
    re.IGNORECASE,
)

LINEAS_PRO = [
    "pro-v", "pro v", "sedal pro", "professional", "profesional",
    "expert", "repair", "elvive", "total repair", "liso total",
    "fructis pro", "pantene pro", "tresemme pro",
]

TIPOS_CABELLO = {
    "rizado": ["rulos", "rizos", "rizado", "curly", "curl", "ondulado"],
    "tratado": [
        "teñido", "tintura", "color", "tratado", "decolorado",
        "quimico", "quimicamente", "keratina", "alisado", "lacio",
    ],
    "anticaspa": ["anticaspa", "dandruff", "caspa"],
    "seco_danado": ["seco", "dañado", "danado", "reseco", "quebradizo", "hidrat"],
    "graso": ["graso", "grasa", "oleoso"],
    "bebe": ["bebe", "bebé", "niño", "infantil", "kids", "suave"],
    "general": [],
}


def extraer_volumen_ml(nombre: str) -> float | None:
    m = ML_VOLUME_RE.search(nombre)
    if not m:
        return None
    valor_str = m.group(1).replace(",", ".")
    valor = float(valor_str)
    unidad = m.group(2).lower()
    if unidad.startswith("l") and not unidad.startswith("liso"):
        return valor * 1000
    if unidad in ("kg", "kilo", "kilos"):
        return valor * 1000
    return valor


def clasificar_linea(nombre: str, marca: str) -> str:
    texto = (nombre + " " + marca).lower()
    for kw in LINEAS_PRO:
        if kw in texto:
            return "profesional"
    return "estandar"


def clasificar_tipo_cabello(nombre: str) -> str:
    texto = nombre.lower()
    for tipo, keywords in TIPOS_CABELLO.items():
        if tipo == "general":
            continue
        for kw in keywords:
            if kw in texto:
                return tipo
    return "general"


def clasificar_producto(nombre: str) -> str:
    texto = nombre.lower()
    if any(kw in texto for kw in ["acondicionador", "enjuague", "rinse", "conditioner"]):
        return "acondicionador"
    if any(kw in texto for kw in ["tratamiento", "mascarilla", "mascara", "mascara", "pack", "kit"]):
        return "tratamiento"
    return "shampoo"


def fetch_vtex(base_url: str, term: str, page_size: int = 50, max_pages: int = 20) -> list[dict]:
    productos = {}
    for page in range(max_pages):
        offset = page * page_size
        url = f"{base_url}{term}?_from={offset}&_to={offset + page_size - 1}"
        try:
            r = requests.get(url, headers=HEADERS, timeout=15)
            data = r.json()
            if not data:
                break
            for item in data:
                pid = item.get("productId")
                if pid and pid not in productos:
                    productos[pid] = item
            print(f"  [{term}] página {page+1}: {len(data)} items, acumulado: {len(productos)}")
            if len(data) < page_size:
                break
            time.sleep(0.4)
        except Exception as e:
            print(f"  Error en {url}: {e}")
            break
    return list(productos.values())


def vtex_a_filas(raw_items: list[dict], fuente: str) -> list[dict]:
    filas = []
    for item in raw_items:
        nombre = item.get("productName", "")
        marca = item.get("brand", "")
        categorias = item.get("categories", [])
        link = item.get("link", "")

        for sku in item.get("items", []):
            sku_nombre = sku.get("nameComplete") or sku.get("name") or nombre
            for seller in sku.get("sellers", []):
                oferta = seller.get("commertialOffer", {})
                precio = oferta.get("Price")
                disponible = oferta.get("IsAvailable", False)
                if not precio:
                    precio = oferta.get("ListPrice")
                if not precio:
                    continue

                volumen = extraer_volumen_ml(sku_nombre)
                precio_por_ml = round(precio / volumen, 4) if volumen else None

                filas.append({
                    "fuente": fuente,
                    "producto_id": item.get("productId"),
                    "sku_id": sku.get("itemId"),
                    "nombre": sku_nombre,
                    "marca": marca,
                    "precio_ars": precio,
                    "precio_lista": oferta.get("ListPrice"),
                    "disponible": disponible,
                    "volumen_ml": volumen,
                    "precio_por_ml": precio_por_ml,
                    "categoria_raw": categorias[0].replace("/", " > ").strip() if categorias else "",
                    "tipo_producto": clasificar_producto(sku_nombre),
                    "linea_tipo": clasificar_linea(sku_nombre, marca),
                    "tipo_cabello": clasificar_tipo_cabello(sku_nombre),
                    "url": link,
                })
    return filas
