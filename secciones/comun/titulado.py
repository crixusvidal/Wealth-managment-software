"""Copia de una figura con nombre (título, subtítulo) arriba y fuente/notas abajo dentro de la imagen, para compartirla fuera del informe.
Mismo diseño que `titulado` de industria/graficos/industria.py: barra roja + título en mayúsculas, subtítulo gris y pie con la fuente."""
import textwrap

import matplotlib.pyplot as plt
from matplotlib.patches import Rectangle

from estilo import NEGRO, ROJO, TINTA_SECUNDARIA

ANCHO = 5.0     # más ancho que la columna del informe: el texto del título y de la fuente cabe sin achicarlo


def titulado(png, titulo, subtitulo, notas, salida):
    img = plt.imread(png)
    alto_img = ANCHO * img.shape[0] / img.shape[1]
    lin_t = textwrap.wrap(titulo.upper(), 40)
    lin_s = textwrap.wrap(subtitulo, 70) if subtitulo else []
    lin_n = [l for nota in notas for l in textwrap.wrap(nota, 84)]
    cab = 0.2 + 0.24 * len(lin_t) + 0.17 * len(lin_s) + 0.12
    pie = 0.14 + 0.15 * len(lin_n)
    h = alto_img + cab + pie
    fig = plt.figure(figsize=(ANCHO, h))
    fig.patch.set_facecolor("white")
    ax = fig.add_axes([0, pie / h, 1, alto_img / h])
    ax.imshow(img)
    ax.axis("off")
    y = h - 0.16
    fig.add_artist(Rectangle((0.03, (y - 0.24 * len(lin_t) + 0.02) / h), 0.05 / ANCHO, (0.24 * len(lin_t) - 0.02) / h, color=ROJO, transform=fig.transFigure))
    for linea in lin_t:
        fig.text(0.03 + 0.1 / ANCHO, y / h, linea, fontsize=12, fontweight="bold", color=NEGRO, va="top")
        y -= 0.24
    y -= 0.04
    for linea in lin_s:
        fig.text(0.03, y / h, linea, fontsize=8.5, color=TINTA_SECUNDARIA, va="top")
        y -= 0.17
    for k, linea in enumerate(lin_n):
        fig.text(0.03, (pie - 0.08 - k * 0.15) / h, linea, fontsize=7.5, color=TINTA_SECUNDARIA, va="top")
    fig.savefig(salida, dpi=400, facecolor="white")
    plt.close(fig)
