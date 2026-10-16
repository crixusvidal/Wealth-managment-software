"""Oso (amarillo) y toro (negro) para la figura de escenarios, como VECTORES editables.

Usa los íconos de Game-icons.net (CC BY 3.0) guardados en valuacion/graficos/iconos/*.svg: `bear-head` (oso gruñendo) y
`charging-bull` (toro embistiendo), ambos mirando a la derecha, hacia el precio de cierre. Cada SVG es un solo trazo: se
recolorea o se retoca en Inkscape / Illustrator / Figma y esta figura lo toma tal cual. Colores institucionales: oso
amarillo, toro negro. Los dibujos se cargan con svgpath2mpl y se agregan como parches de matplotlib (siguen siendo vectores).

Si existen `iconos/oso.png` / `iconos/toro.png` (p. ej. ilustraciones hechas en ChatGPT, PNG con fondo transparente), se usan EN
LUGAR de los íconos SVG: se recortan al contorno visible y se escalan al alto pedido. Deben mirar hacia la DERECHA (si miran a la
izquierda, agregar el sufijo `_izq`: `oso_izq.png`, `toro_izq.png`, y se voltean solos).

Uso como módulo:
    from mascotas import dibujar_oso, dibujar_toro
    ancho = dibujar_oso(ax, x_derecha, y_base, alto)   # el borde derecho del dibujo queda en x_derecha; devuelve su ancho
"""
import re
import sys
from pathlib import Path

import numpy as np
from matplotlib.patches import PathPatch
from matplotlib.path import Path as MPath
from svgpath2mpl import parse_path

sys.path.insert(0, str(Path(__file__).resolve().parents[2] / "comun"))  # paleta institucional
import estilo  # noqa: E402,F401
from estilo import AMARILLO, NEGRO

ICONOS = Path(__file__).parent / "iconos"
ARCHIVOS = {"oso": ("oso_bear-head.svg", AMARILLO), "toro": ("toro_charging-bull.svg", NEGRO)}


def cargar(nombre):
    """Lee el SVG del ícono y devuelve (Path de matplotlib normalizado, razón ancho/alto, color)."""
    archivo, color = ARCHIVOS[nombre]
    d = re.search(r'\sd="([^"]+)"', (ICONOS / archivo).read_text()).group(1)
    p = parse_path(d)
    v = p.vertices.copy()
    v[:, 1] = -v[:, 1]                                   # SVG tiene el eje y hacia abajo
    x0, y0 = v.min(axis=0)
    x1, y1 = v.max(axis=0)
    v = (v - [x0, y0]) / (y1 - y0)                       # alto = 1, base en y = 0
    return MPath(v, p.codes), (x1 - x0) / (y1 - y0), color


def png_de(nombre):
    """(ruta, voltear) del PNG de reemplazo, si existe."""
    for sufijo, voltear in (("", False), ("_izq", True)):
        r = ICONOS / f"{nombre}{sufijo}.png"
        if r.exists():
            return r, voltear
    return None, False


def dibujar_png(ax, ruta, voltear, x_der, y, alto, zorder=4):
    from PIL import Image
    im = Image.open(ruta).convert("RGBA")
    a = np.asarray(im).copy()
    if a[..., 3].min() == 255:                       # sin transparencia (fondo blanco): el blanco del borde pasa a transparente
        from scipy import ndimage
        blanco = (a[..., :3] >= 244).all(axis=-1)
        etiquetas, _ = ndimage.label(blanco)
        borde = np.unique(np.concatenate([etiquetas[0], etiquetas[-1], etiquetas[:, 0], etiquetas[:, -1]]))
        fondo = np.isin(etiquetas, borde[borde > 0])     # solo el blanco conectado con el borde; el blanco interno (ojos, cuernos) queda
        a[fondo, 3] = 0
        im = Image.fromarray(a)
    caja = im.getchannel("A").point(lambda a: 255 if a > 8 else 0).getbbox()   # contorno visible
    if caja:
        im = im.crop(caja)
    if voltear:
        im = im.transpose(Image.FLIP_LEFT_RIGHT)
    ancho = im.width / im.height * alto
    ax.imshow(np.asarray(im), extent=(x_der - ancho, x_der, y, y + alto), zorder=zorder, interpolation="lanczos", aspect="auto")
    return ancho


def dibujar(ax, nombre, x_der, y, alto, zorder=4):
    ruta, voltear = png_de(nombre)
    if ruta:
        return dibujar_png(ax, ruta, voltear, x_der, y, alto, zorder)
    p, razon, color = cargar(nombre)
    ancho = razon * alto
    v = p.vertices * alto + [x_der - ancho, y]
    ax.add_patch(PathPatch(MPath(v, p.codes), facecolor=color, edgecolor="none", linewidth=0, zorder=zorder))
    return ancho


def dibujar_oso(ax, x_der, y, alto):
    return dibujar(ax, "oso", x_der, y, alto)


def dibujar_toro(ax, x_der, y, alto):
    return dibujar(ax, "toro", x_der, y, alto)
