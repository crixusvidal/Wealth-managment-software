# Cómo reemplazar el oso y el toro con ilustraciones de ChatGPT

1. Genera las dos imágenes con el prompt de abajo (una por animal). Pide **fondo transparente**; si no lo logra, un fondo
   **blanco liso** también sirve (el código recorta el blanco del borde solo).
2. Guárdalas en esta carpeta como **`oso.png`** y **`toro.png`** (si el animal mira hacia la izquierda: `oso_izq.png` / `toro_izq.png`,
   se voltean solos). Tienen que mirar hacia el precio, o sea a la derecha.
3. Corre `python3 valuacion/graficos/escenarios.py` (desde `secciones/`): con esos PNG presentes reemplazan a los íconos SVG.
   Para volver a los íconos, borra `oso.png` y `toro.png`.

## Prompt (copiar y pegar)

Oso:
> Ilustración vectorial plana (flat vector, bordes limpios, sin degradados ni sombras) de un OSO PARDO atacando, de cuerpo
> entero, visto de perfil mirando a la DERECHA, erguido sobre las patas traseras con las garras en alto y la boca abierta con
> colmillos, expresión feroz. Relleno amarillo #ecb100, contorno grueso azul oscuro #14284B, ojos y lengua rojos #e9004b, garras
> y colmillos blancos. Fondo transparente (PNG), sin texto, sin suelo, composición centrada, 2048 px de alto.

Toro:
> Ilustración vectorial plana (flat vector, bordes limpios, sin degradados ni sombras) de un TORO NEGRO embistiendo, de cuerpo
> entero, visto de perfil mirando a la DERECHA, cabeza baja y cuernos al frente, patas en pleno galope, líneas de velocidad
> detrás, expresión furiosa. Relleno negro #000000, cuernos blancos con contorno negro, ojos y fosas nasales rojos #e9004b.
> Mismo estilo, grosor de línea y proporciones que el oso. Fondo transparente (PNG), sin texto, sin suelo, composición centrada,
> 2048 px de alto.

Consejos: pedir las dos en la misma conversación para que el estilo coincida; si salen con fondo de color o sombra, pedir
«sin sombra y con fondo transparente». Si ChatGPT entrega SVG, guardarlo igual y avísame para cargarlo como vector.
