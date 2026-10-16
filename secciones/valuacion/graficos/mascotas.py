"""Siluetas vectoriales del oso (bear) y del toro (bull) atacando, en colores institucionales.

Oso amarillo (#ecb100) erguido con las garras en alto; toro negro embistiendo con la cabeza baja.
Detalles en rojo (ojos y boca), blanco (garras, cuernos y dientes) y azul institucional (contorno del oso).
Se construyen con shapely (uniones de elipses y trazos gruesos suavizados) y se dibujan como parches de matplotlib.

Uso como módulo:
    from mascotas import dibujar_oso, dibujar_toro
    dibujar_oso(ax, x, y, alto)      # (x, y) = esquina inferior izquierda; alto en unidades del eje
Uso suelto (hoja de prueba):  python valuacion/graficos/mascotas.py
"""
import sys
from pathlib import Path

import matplotlib.pyplot as plt
from matplotlib.patches import PathPatch, Polygon
from matplotlib.path import Path as MPath
from shapely import affinity
from shapely.geometry import LineString, Point
from shapely.ops import unary_union

sys.path.insert(0, str(Path(__file__).resolve().parents[2] / "comun"))  # paleta institucional
import estilo  # noqa: E402,F401
from estilo import AMARILLO, AZUL, NEGRO, ROJO

BLANCO = "#ffffff"


# Primitivas ---------------------------------------------------------------------------------------
def elipse(cx, cy, rx, ry, rot=0.0):
    return affinity.rotate(affinity.scale(Point(cx, cy).buffer(1, 48), rx, ry), rot, origin=(cx, cy))


def trazo(puntos, ancho):
    return LineString(puntos).buffer(ancho / 2, cap_style=1, join_style=1, quad_segs=24)


def suave(geom, r=1.2):
    return geom.buffer(r, quad_segs=24).buffer(-r, quad_segs=24)


def parche(geom, **kw):
    """Shapely (Polygon/MultiPolygon, con huecos) → PathPatch de matplotlib."""
    partes = geom.geoms if hasattr(geom, "geoms") else [geom]
    verts, codes = [], []
    for p in partes:
        for anillo in [p.exterior, *p.interiors]:
            xy = list(anillo.coords)
            verts += xy
            codes += [MPath.MOVETO] + [MPath.LINETO] * (len(xy) - 2) + [MPath.CLOSEPOLY]
    return PathPatch(MPath(verts, codes), **kw)


def poner(ax, geom, x, y, escala, **kw):
    """Dibuja geom (coordenadas de diseño 0–100) con su esquina inferior izquierda en (x, y)."""
    g = affinity.translate(affinity.scale(geom, escala, escala, origin=(0, 0)), x, y)
    ax.add_patch(parche(g, **kw))


# Toro (embistiendo hacia la derecha) --------------------------------------------------------------
def pata(muslo, rodilla, pie, ancho=7.0):
    return unary_union([trazo([muslo, rodilla], ancho), trazo([rodilla, pie], ancho * 0.62),
                        elipse(pie[0] + 1.2, pie[1] - 0.6, 3.4, 2.4, -20)])


def toro():
    cuerpo = unary_union([
        elipse(62, 34, 21, 17, -8),            # pecho
        elipse(37, 33, 19, 12.5, -4),          # grupa
        elipse(63, 46, 13, 10.5, -14),         # joroba
        trazo([(72, 40), (83, 29)], 19),       # cuello
        elipse(86, 25, 12.5, 8.2, -42),        # cabeza baja
        elipse(94, 18.5, 6.4, 4.8, -42),       # hocico
        pata((70, 24), (79, 15), (88, 11)),    # delantera adelantada
        pata((60, 24), (60, 13), (52, 7)),     # delantera atrasada
        pata((30, 25), (20, 17), (10, 13)),    # trasera extendida
        pata((38, 23), (33, 12), (24, 5)),     # trasera plegada
        trazo([(19, 38), (12, 45), (9, 52)], 2.4), elipse(8.5, 54, 2.4, 4.6, 8),   # cola con mechón
    ])
    cuerno = unary_union([trazo([(80, 34), (89, 39.5), (98, 38.5)], 4.4), trazo([(98, 38.5), (104, 43.5)], 2.4),
                          trazo([(78, 36), (84, 45), (92, 50)], 3.6), trazo([(92, 50), (97, 56)], 2.0)])
    return suave(cuerpo, 1.4), cuerno


def dibujar_toro(ax, x, y, alto):
    cuerpo, cuerno = toro()
    k = alto / 62.0

    def sh(g):
        return affinity.translate(affinity.scale(g, k, k, origin=(0, 0)), x, y)

    ax.add_patch(parche(sh(Polygon_sin_huecos(cuerpo)), facecolor=NEGRO, edgecolor=NEGRO, linewidth=0.4, zorder=4))
    ax.add_patch(parche(sh(cuerno), facecolor=BLANCO, edgecolor=NEGRO, linewidth=0.9, zorder=5))   # cuernos al frente
    ax.plot([x + 80 * k, x + 90 * k, x + 96 * k], [y + 19 * k, y + 15 * k, y + 14 * k], color=BLANCO, linewidth=0.9, zorder=5.5, solid_capstyle="round")   # línea de la quijada
    ax.add_patch(parche(sh(Point(89.5, 28.5).buffer(1.7)), facecolor=ROJO, edgecolor="none", zorder=6))        # ojo
    ax.add_patch(parche(sh(Point(98.5, 19.5).buffer(1.1)), facecolor=ROJO, edgecolor="none", zorder=6))        # fosa nasal
    ax.plot([x + 86 * k, x + 92 * k], [y + 31.5 * k, y + 30.5 * k], color=ROJO, linewidth=1.1, zorder=6, solid_capstyle="round")  # ceño
    for dy, ln in ((14, 20), (24, 28), (34, 18), (44, 12)):                                                   # líneas de velocidad
        ax.plot([x - 1 * k, x - (1 + ln) * k], [y + dy * k, y + dy * k], color=NEGRO, linewidth=1.0, solid_capstyle="round", zorder=2)


# Oso (erguido, garras en alto, atacando hacia la derecha) -----------------------------------------
def Polygon_sin_huecos(geom):
    """Quita los huecos internos (rellena) para que la silueta sea sólida."""
    from shapely.geometry import MultiPolygon, Polygon as SPoly
    partes = geom.geoms if hasattr(geom, "geoms") else [geom]
    return MultiPolygon([SPoly(p.exterior) for p in partes])


def oso():
    cuerpo = unary_union([
        elipse(41, 38, 17, 28, 10),            # tronco
        trazo([(48, 58), (54, 62)], 14),       # cuello
        elipse(57, 68, 12.5, 11.5, -6),        # cabeza
        elipse(70, 63.5, 9.5, 6.2, -14),       # hocico
        elipse(49.5, 78.5, 4.4, 4.4), elipse(63, 79.5, 4.4, 4.4),             # orejas
        trazo([(46, 56), (64, 59), (75, 73)], 10), elipse(77.5, 76, 5.6, 5.6),             # brazo 1 en alto
        trazo([(44, 50), (62, 45), (83, 56)], 10), elipse(86, 58.5, 5.4, 5.4),             # brazo 2 al frente
        trazo([(35, 22), (30, 7)], 12), trazo([(50, 20), (57, 7)], 12),         # patas traseras
        elipse(28, 4.5, 8.5, 3.8, 4), elipse(60, 4.5, 8.5, 3.8, -4),            # pies
    ])
    return suave(cuerpo, 2.0)


def dibujar_oso(ax, x, y, alto):
    k = alto / 88.0
    c = Polygon_sin_huecos(oso())

    def sh(g):
        return affinity.translate(affinity.scale(g, k, k, origin=(0, 0)), x, y)

    ax.add_patch(parche(sh(c), facecolor=AMARILLO, edgecolor=AZUL, linewidth=1.0, zorder=4))
    ax.add_patch(parche(sh(elipse(69.5, 59.5, 5.2, 2.4, -16)), facecolor=ROJO, edgecolor=AZUL, linewidth=0.5, zorder=6))   # boca abierta
    for tx in (64.5, 68.5, 72.5):                                                                                       # colmillos
        ax.add_patch(parche(sh(Polygon_tri(tx, 61.3)), facecolor=BLANCO, edgecolor=AZUL, linewidth=0.3, zorder=7))
    ax.add_patch(parche(sh(Point(77.5, 66.5).buffer(1.9)), facecolor=NEGRO, edgecolor="none", zorder=6))                # nariz
    ax.add_patch(parche(sh(Point(60.5, 71).buffer(1.7)), facecolor=ROJO, edgecolor="none", zorder=6))                   # ojo
    ax.plot([x + 56 * k, x + 63.5 * k], [y + 75.5 * k, y + 72.5 * k], color=AZUL, linewidth=1.3, zorder=6, solid_capstyle="round")  # ceño
    for (gx, gy, dx, dy) in ((79.5, 80.5, 2.5, 4.5), (82, 78.5, 4.5, 3.2), (76, 81.5, -0.6, 5.0),
                             (88, 62.5, 4.8, 3.0), (89.5, 59, 5.2, 0.3), (87.5, 64, 2.4, 5.0)):
        for ancho, col in ((1.9, BLANCO), (0.7, AZUL)):
            ax.plot([x + gx * k, x + (gx + dx) * k], [y + gy * k, y + (gy + dy) * k], color=col, linewidth=ancho,
                    solid_capstyle="round", zorder=7)


def Polygon_tri(cx, cy):
    from shapely.geometry import Polygon as SPoly
    return SPoly([(cx - 1.4, cy), (cx + 1.4, cy), (cx, cy - 3.0)])


if __name__ == "__main__":
    fig, ax = plt.subplots(figsize=(7, 3.4))
    ax.set_xlim(0, 13.8)
    ax.set_ylim(0, 6.2)
    ax.set_aspect("equal")
    ax.axis("off")
    dibujar_oso(ax, 0.6, 0.4, 5.2)
    dibujar_toro(ax, 6.0, 0.9, 4.4)
    salida = Path(__file__).parent / "figuras" / "mascotas_prueba.png"
    fig.savefig(salida, dpi=200, facecolor="white")
    print(salida)
