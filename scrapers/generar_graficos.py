import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import matplotlib.ticker as mticker
import numpy as np
import pandas as pd
from pathlib import Path

BASE = Path(__file__).parent.parent
CLEAN_PATH = BASE / "data" / "clean" / "dataset_limpio.csv"
CHARTS_DIR = BASE / "charts"
CHARTS_DIR.mkdir(exist_ok=True)

# Paleta consistente con el Excel
C_JUMBO     = "#1A5276"
C_FARMACITY = "#1E8449"
C_DISCO     = "#7D6608"
C_PROF      = "#6C3483"
C_ESTANDAR  = "#2980B9"
C_ACCENT    = "#E74C3C"
C_ACCENT2   = "#F39C12"

PALETTE_FUENTE = {"Jumbo": C_JUMBO, "Farmacity": C_FARMACITY, "Disco": C_DISCO}

plt.rcParams.update({
    "font.family": "sans-serif",
    "font.size": 10,
    "axes.spines.top": False,
    "axes.spines.right": False,
    "axes.titlesize": 12,
    "axes.titleweight": "bold",
    "axes.titlepad": 14,
    "axes.labelsize": 10,
    "figure.dpi": 150,
    "figure.facecolor": "white",
})


def _conclusion(fig, texto, y=-0.06):
    fig.text(0.5, y, texto, ha="center", fontsize=8.5,
             color="#444444", fontstyle="italic",
             wrap=True, transform=fig.transFigure)


def _fmt_ars(ax, axis="y"):
    fmt = mticker.FuncFormatter(lambda v, _: f"${v:,.0f}")
    if axis == "y":
        ax.yaxis.set_major_formatter(fmt)
    else:
        ax.xaxis.set_major_formatter(fmt)


def save(fig, name):
    path = CHARTS_DIR / f"{name}.png"
    fig.savefig(path, bbox_inches="tight", facecolor="white")
    plt.close(fig)
    print(f"  {name}.png")
    return path


# ─────────────────────────────────────────────────────────────────────────────
# Gráfico 1 — Torta: distribución de productos por fuente
# ─────────────────────────────────────────────────────────────────────────────
def g1_distribucion_fuente(df):
    conteo = df["fuente"].value_counts().reindex(["Farmacity", "Jumbo", "Disco"])
    colores = [PALETTE_FUENTE[f] for f in conteo.index]

    fig, ax = plt.subplots(figsize=(7, 5))
    wedges, _, autotexts = ax.pie(
        conteo.values,
        labels=conteo.index,
        autopct="%1.1f%%",
        colors=colores,
        startangle=90,
        pctdistance=0.72,
        wedgeprops={"edgecolor": "white", "linewidth": 2},
    )
    for at in autotexts:
        at.set_fontsize(11)
        at.set_color("white")
        at.set_fontweight("bold")

    ax.set_title("Distribución de productos disponibles por fuente")

    lider = conteo.idxmax()
    pct = 100 * conteo.max() / conteo.sum()
    _conclusion(fig,
        f"Conclusión: {lider} concentra el {pct:.1f}% de los productos disponibles, "
        f"más del doble que Jumbo y Disco. La amplitud del catálogo varía significativamente entre retailers.")
    return save(fig, "01_distribucion_fuente")


# ─────────────────────────────────────────────────────────────────────────────
# Gráfico 2 — Barras agrupadas: precio_ars por fuente (H3)
# ─────────────────────────────────────────────────────────────────────────────
def g2_precio_por_fuente(df):
    stats = (df.groupby("fuente")["precio_ars"]
               .agg(media="mean", mediana="median")
               .reindex(["Jumbo", "Farmacity", "Disco"]))
    x = np.arange(3)
    w = 0.35
    colores = [PALETTE_FUENTE[f] for f in stats.index]

    fig, ax = plt.subplots(figsize=(8, 5))
    b1 = ax.bar(x - w/2, stats["media"],   w, label="Media",   color=colores, alpha=0.92)
    b2 = ax.bar(x + w/2, stats["mediana"], w, label="Mediana", color=colores, alpha=0.50,
                hatch="//", edgecolor="white")

    ax.set_xticks(x)
    ax.set_xticklabels(stats.index)
    ax.set_ylabel("Precio ARS")
    ax.set_title("H3 — Precio promedio y mediana por fuente")
    ax.legend()
    _fmt_ars(ax)

    for bar in b1:
        ax.text(bar.get_x() + bar.get_width()/2, bar.get_height() + 80,
                f"${bar.get_height():,.0f}", ha="center", va="bottom", fontsize=8.5)

    caro  = stats["media"].idxmax()
    barat = stats["media"].idxmin()
    ratio = stats.loc[caro, "media"] / stats.loc[barat, "media"]
    _conclusion(fig,
        f"Conclusión: {caro} presenta el precio promedio más alto. "
        f"La diferencia con {barat} es {ratio:.1f}x. "
        f"En todos los casos la mediana < media, señal de outliers hacia precios altos.")
    return save(fig, "02_precio_por_fuente")


# ─────────────────────────────────────────────────────────────────────────────
# Gráfico 3 — Columnas: precio/ml profesional vs estándar (H1)
# ─────────────────────────────────────────────────────────────────────────────
def g3_h1_linea(df):
    df_ml = df[df["precio_por_ml"].notna() & df["linea_tipo"].notna()]
    stats = (df_ml.groupby("linea_tipo")["precio_por_ml"]
                  .agg(media="mean", mediana="median", n="count")
                  .reindex(["profesional", "estandar"]))

    fig, ax = plt.subplots(figsize=(7, 5))
    colores = [C_PROF, C_ESTANDAR]
    bars = ax.bar(["Profesional", "Estandar"], stats["media"],
                  color=colores, width=0.45, edgecolor="white")
    ax.bar(["Profesional", "Estandar"], stats["mediana"],
           color=colores, width=0.45, alpha=0.4, hatch="//",
           edgecolor="white", label="Mediana")

    for bar, (_, row) in zip(bars, stats.iterrows()):
        ax.text(bar.get_x() + bar.get_width()/2, bar.get_height() + 0.2,
                f"${row['media']:.1f}/ml\nn={int(row['n'])}",
                ha="center", va="bottom", fontsize=9)

    ax.set_ylabel("Precio ARS por ml")
    ax.set_title("H1 — Precio por ml: linea Profesional vs Estandar")
    ax.legend()

    ratio = stats.loc["profesional", "media"] / stats.loc["estandar", "media"]
    _conclusion(fig,
        f"Conclusión: La linea profesional cuesta en promedio {ratio:.1f}x mas por ml que la estandar, "
        f"lo que valida la hipotesis H1. La diferencia en medianas es menor, "
        f"indicando que algunos productos profesionales de alto precio elevan la media.")
    return save(fig, "03_h1_linea_profesional")


# ─────────────────────────────────────────────────────────────────────────────
# Gráfico 4 — Barras horizontales: precio/ml por tipo de cabello (H2)
# ─────────────────────────────────────────────────────────────────────────────
def g4_h2_tipo_cabello(df):
    df_ml = df[df["precio_por_ml"].notna() & df["tipo_cabello"].notna()]
    stats = (df_ml.groupby("tipo_cabello")["precio_por_ml"]
                  .agg(media="mean", n="count")
                  .sort_values("media", ascending=True))

    COLORES_TC = {
        "general":    "#AAB7B8",
        "graso":      "#7F8C8D",
        "anticaspa":  "#E67E22",
        "bebe":       "#1ABC9C",
        "seco_danado":"#E74C3C",
        "tratado":    "#2980B9",
        "rizado":     "#8E44AD",
    }
    colores = [COLORES_TC.get(i, "#888888") for i in stats.index]

    fig, ax = plt.subplots(figsize=(8, 5))
    bars = ax.barh(stats.index, stats["media"], color=colores, edgecolor="white")

    for bar, (idx, row) in zip(bars, stats.iterrows()):
        ax.text(bar.get_width() + 0.15, bar.get_y() + bar.get_height()/2,
                f"${row['media']:.1f}  (n={int(row['n'])})",
                va="center", fontsize=8.5)

    ax.set_xlabel("Precio ARS por ml")
    ax.set_title("H2 — Precio por ml segun tipo de cabello")
    ax.set_xlim(0, stats["media"].max() * 1.35)

    top = stats.index[-1]
    bot = stats.index[0]
    ratio = stats.loc[top, "media"] / stats.loc[bot, "media"]
    _conclusion(fig,
        f"Conclusion: '{top}' tiene el precio/ml mas alto y '{bot}' el mas bajo "
        f"(diferencia {ratio:.1f}x). Los productos para necesidades especificas "
        f"(cabello rizado, tratado, danado) presentan mayor precio por ml, validando H2 parcialmente.")
    return save(fig, "04_h2_tipo_cabello")


# ─────────────────────────────────────────────────────────────────────────────
# Gráfico 5 — Histograma: distribución de precio_ars
# ─────────────────────────────────────────────────────────────────────────────
def g5_histograma_precio(df):
    p99 = df["precio_ars"].quantile(0.99)
    df_p = df[df["precio_ars"] <= p99]
    media = df["precio_ars"].mean()
    mediana = df["precio_ars"].median()

    fig, ax = plt.subplots(figsize=(9, 5))
    ax.hist(df_p["precio_ars"], bins=45, color=C_JUMBO, alpha=0.85, edgecolor="white")
    ax.axvline(media,   color=C_ACCENT,  linestyle="--", linewidth=1.8, label=f"Media: ${media:,.0f}")
    ax.axvline(mediana, color=C_ACCENT2, linestyle="-",  linewidth=1.8, label=f"Mediana: ${mediana:,.0f}")

    ax.set_xlabel("Precio ARS")
    ax.set_ylabel("Frecuencia")
    ax.set_title("Distribucion de precios — precio_ars (sin top 1%)")
    ax.legend()
    _fmt_ars(ax, "x")

    skew = df["precio_ars"].skew()
    _conclusion(fig,
        f"Conclusion: Distribucion con asimetria positiva ({skew:.2f}). "
        f"La mayoria de productos se concentra en precios bajos-medios, "
        f"con cola larga hacia precios altos. Media (${media:,.0f}) > Mediana (${mediana:,.0f}), "
        f"efecto tipico de outliers superiores.")
    return save(fig, "05_histograma_precio")


# ─────────────────────────────────────────────────────────────────────────────
# Gráfico 6 — Barras horizontales: top 10 marcas
# ─────────────────────────────────────────────────────────────────────────────
def g6_top_marcas(df):
    top = df["marca"].value_counts().head(10).sort_values(ascending=True)

    fig, ax = plt.subplots(figsize=(9, 6))
    bars = ax.barh(top.index, top.values, color=C_DISCO, edgecolor="white", alpha=0.9)

    for bar in bars:
        ax.text(bar.get_width() + 1, bar.get_y() + bar.get_height()/2,
                str(int(bar.get_width())), va="center", fontsize=9)

    ax.set_xlabel("Cantidad de productos disponibles")
    ax.set_title("Top 10 marcas por cantidad de productos disponibles")
    ax.set_xlim(0, top.max() * 1.15)

    top1, n1 = top.index[-1], top.iloc[-1]
    top3 = top.iloc[-3:].sum()
    _conclusion(fig,
        f"Conclusion: {top1} lidera con {n1} productos ({100*n1/len(df):.1f}% del dataset). "
        f"Las 3 principales marcas concentran {100*top3/len(df):.1f}% de los productos disponibles. "
        f"El mercado presenta alta concentracion en pocas marcas masivas.")
    return save(fig, "06_top_marcas")


# ─────────────────────────────────────────────────────────────────────────────
# Gráfico 7 — Columnas apiladas: tipo de producto por fuente
# ─────────────────────────────────────────────────────────────────────────────
def g7_tipo_producto_fuente(df):
    pivot = (df.groupby(["fuente", "tipo_producto"]).size()
               .unstack(fill_value=0)
               .reindex(["Jumbo", "Farmacity", "Disco"]))

    tipos   = pivot.columns.tolist()
    colores = {
        "shampoo":       C_JUMBO,
        "acondicionador": C_FARMACITY,
        "tratamiento":   C_DISCO,
    }

    fig, ax = plt.subplots(figsize=(8, 5))
    bottom = np.zeros(3)
    for t in tipos:
        vals = pivot[t].values.astype(float)
        color = colores.get(t, "#888888")
        bars = ax.bar(pivot.index, vals, bottom=bottom,
                      label=t.capitalize(), color=color, edgecolor="white", width=0.5)
        for bar, v, b in zip(bars, vals, bottom):
            if v >= 20:
                ax.text(bar.get_x() + bar.get_width()/2, b + v/2,
                        str(int(v)), ha="center", va="center",
                        fontsize=9, color="white", fontweight="bold")
        bottom += vals

    ax.set_ylabel("Cantidad de productos")
    ax.set_title("Composicion del catalogo por tipo de producto y fuente")
    ax.legend(loc="upper right")

    _conclusion(fig,
        "Conclusion: El shampoo domina el catalogo en los tres retailers. "
        "Farmacity ofrece la mayor proporcion de acondicionadores y tratamientos, "
        "coherente con su perfil de farmacia orientada al cuidado capilar.")
    return save(fig, "07_tipo_producto_fuente")


# ─────────────────────────────────────────────────────────────────────────────
# Gráfico 8 — Dispersión: precio_ars vs volumen_ml
# ─────────────────────────────────────────────────────────────────────────────
def g8_dispersion_precio_volumen(df):
    df_p = df[
        df["volumen_ml"].notna() &
        (df["volumen_ml"] <= 2000) &
        (df["precio_ars"] <= df["precio_ars"].quantile(0.97))
    ]

    fig, ax = plt.subplots(figsize=(9, 6))
    for fuente, color in PALETTE_FUENTE.items():
        sub = df_p[df_p["fuente"] == fuente]
        ax.scatter(sub["volumen_ml"], sub["precio_ars"],
                   alpha=0.35, s=18, color=color, label=fuente)

    x = df_p["volumen_ml"].values
    y = df_p["precio_ars"].values
    m, b = np.polyfit(x, y, 1)
    xs = np.linspace(x.min(), x.max(), 200)
    ax.plot(xs, m * xs + b, color=C_ACCENT, linewidth=1.8,
            linestyle="--", label=f"Tendencia (y={m:.1f}x+{b:.0f})")

    ax.set_xlabel("Volumen (ml)")
    ax.set_ylabel("Precio ARS")
    ax.set_title("Relacion entre volumen (ml) y precio ARS por fuente")
    ax.legend()
    _fmt_ars(ax)

    corr = df_p[["volumen_ml", "precio_ars"]].corr().iloc[0, 1]
    _conclusion(fig,
        f"Conclusion: Correlacion positiva moderada entre volumen y precio (r={corr:.2f}). "
        f"A mayor contenido, mayor precio, aunque con alta dispersion. "
        f"Los tres retailers muestran patrones similares sin diferencias sistematicas por fuente.")
    return save(fig, "08_dispersion_precio_volumen")


def main():
    df = pd.read_csv(CLEAN_PATH)
    print(f"Dataset: {len(df)} registros — generando 8 graficos...")

    g1_distribucion_fuente(df)
    g2_precio_por_fuente(df)
    g3_h1_linea(df)
    g4_h2_tipo_cabello(df)
    g5_histograma_precio(df)
    g6_top_marcas(df)
    g7_tipo_producto_fuente(df)
    g8_dispersion_precio_volumen(df)

    print(f"\nListo. Graficos en: {CHARTS_DIR}")


if __name__ == "__main__":
    main()
