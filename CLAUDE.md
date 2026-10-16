# Reglas del proyecto (informe CFA Mallplaza)

- Responder siempre en español.
- Colores, formato y cifras: ver `narrativa_y_coherencia.md` (sección 4).
- **Solo dos versiones:** de cada documento y gráfico se conservan la versión actual y la penúltima. Al crear una versión nueva, eliminar la antepenúltima y el Python que solo la generaba, y actualizar `generar_todo.sh` y la tabla "Versiones vigentes" de `narrativa_y_coherencia.md`. Las variantes de una misma versión (ES/EN, docx/pdf, normal/titulado) no cuentan como versiones; los modelos Excel quedan fuera.
- **Guardar todo ordenado:** cada archivo en su carpeta según el mapa de `narrativa_y_coherencia.md` (sección 7); intermedios en subcarpetas (`maquetacion/`, `datos/`), nunca sueltos; sin temporales en el repo (ver `.gitignore`). Al agregar una carpeta o script nuevo, actualizar ese mapa y `generar_todo.sh`. Commit y push tras cada cambio.
- **Contexto del LLM:** antes de trabajar con cifras, lee `contexto_llm/cifras_clave.md` (cifras vigentes con su fuente, generadas de los modelos; no se editan a mano). Si cambia un modelo o el informe, invoca al agente `contexto-llm` (o corre `python3 contexto_llm/generar.py` y `python3 informe_consolidado/actualizar.py`): reemplaza lo viejo por lo nuevo.
- **Entregables:** los finales viven en `informe_consolidado/vigente/` (y la penúltima en `anterior/`); no dejar copias sueltas.

