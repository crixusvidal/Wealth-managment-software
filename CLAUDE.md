# Reglas del proyecto (informe CFA Mallplaza)

- Responder siempre en español.
- Colores, formato y cifras: ver `narrativa_y_coherencia.md` (sección 4).
- **Solo dos versiones:** de cada documento y gráfico se conservan la versión actual y la penúltima. Al crear una versión nueva, eliminar la antepenúltima y el Python que solo la generaba, y actualizar `generar_todo.sh` y la tabla "Versiones vigentes" de `narrativa_y_coherencia.md`. Las variantes de una misma versión (ES/EN, docx/pdf, normal/titulado) no cuentan como versiones; los modelos Excel quedan fuera.
- **Guardar todo ordenado:** cada archivo en su carpeta según el mapa de `narrativa_y_coherencia.md` (sección 7); intermedios en subcarpetas (`maquetacion/`, `datos/`), nunca sueltos; sin temporales en el repo (ver `.gitignore`). Al agregar una carpeta o script nuevo, actualizar ese mapa y `generar_todo.sh`. Commit y push tras cada cambio.
