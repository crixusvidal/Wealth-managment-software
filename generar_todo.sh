#!/usr/bin/env bash
# Regenera todos los gráficos y documentos del informe a partir de los modelos de cada sección.
# Uso: bash generar_todo.sh        (requiere python3 con matplotlib/openpyxl y node con docx)
set -euo pipefail
ROOT="$(cd "$(dirname "$0")" && pwd)"
cd "$ROOT/secciones"
export NODE_PATH="${NODE_PATH:+$NODE_PATH:}/opt/node22/lib/node_modules:/opt/node-tools/node_modules"
V=valuacion/modelos; F=analisis_financiero/modelos; R=riesgos/modelos
V4=$V/Mallplaza_Valoracion_Consolidada_v4.xlsx   # = v3 + pipeline completo con fecha (170.200 m²: «Otros proyectos Chile» activados); v3 = v2 + hojas TIR
FA=$F/Mallplaza_FA_Exhibits_v7.xlsx               # v7 = v6 con «Otros proyectos Chile» activados (apertura 2028)
TOR=$R/Mallplaza_Tornado_Sensibilidad_Operacional_2026_2028.xlsx

echo "== Modelo v4 y datos de las figuras de Valuación"
python3 valuacion/graficos/modelo_v4.py          # Exhibits v7 + modelo v4 (y variante ESG) + valores pegados (tabla de deuda, TIR implícita, escenarios)
python3 valuacion/graficos/actualizar_figuras.py $V4 $FA   # datos de V1–V6 desde el modelo v4 → Mallplaza_Figuras_Valuacion_recalc.xlsx

echo "== Gráficos"
python3 negocio/graficos/negocio.py negocio/modelos/Excel_bussines_para_armar_graficos.xlsx
python3 negocio/graficos/mapa_sudamerica.py   # Figura 1 (mapa de Sudamérica; requiere geopandas<1.0)
python3 industria/graficos/industria.py industria/modelos/Excel_DATA_INDUSTRY.xlsx
python3 industria/graficos/cuota_mercado.py $V/Comparables_Mallplaza_CFA_BBG_valoracion.xlsx
python3 valuacion/graficos/valuacion.py $V/Mallplaza_Figuras_Valuacion_recalc.xlsx - $V4
python3 valuacion/graficos/anexos_valuacion.py $V4 $V/Mallplaza_Figuras_Valuacion_recalc.xlsx $FA $TOR
python3 valuacion/graficos/tir_valuacion.py $V4   # TIR firma/accionista (nominal y real) → figuras/datos_tir.json (lo lee valuacion.js)
python3 valuacion/graficos/tir_proyecto.py $V4   # TIR del proyecto nuevo (1 m² de expansión), hoja TIR_Proyecto del modelo v3 → figuras/datos_tir_proyecto.json
python3 valuacion/graficos/supuestos_cascada.py $V4 $FA $TOR   # cascada de supuestos fuertes (Figura 29) → figuras/{es,en}_V13_supuestos_cascada.png + datos_supuestos.json (escudo tributario)
python3 valuacion/graficos/sensibilidad_tasas.py $V4 $V/Mallplaza_Tasas_Historico.xlsx   # sensibilidad a tasas largas e inflación → modelos/Mallplaza_Sensibilidad_Tasas_Inflacion.xlsx + figuras V11/V12 (recalcular el Excel con LibreOffice)
python3 /root/.claude/skills/synced/508cc6bc-ad76-422f-ada9-ad0fc4279625_0b2a2ef9-bf53-46a5-b8ed-72255cce463a/xlsx/scripts/recalc.py $V/Mallplaza_Sensibilidad_Tasas_Inflacion.xlsx >/dev/null   # el script escribe fórmulas sin valores: recalcular con LibreOffice
python3 valuacion/graficos/football_field.py $V4
python3 valuacion/graficos/escenarios.py      # boceto: escenarios bear/base/bull con oso y toro
python3 valuacion/graficos/tabla_kd_caprates.py $V/Comparables_Mallplaza_CFA_BBG_valoracion.xlsx
python3 valuacion/graficos/tabla_multiplos_implicitos.py $FA
python3 analisis_financiero/graficos/analisis_financiero.py
python3 riesgos/graficos/pagina_riesgos.py $FA $V4 $TOR
python3 esg/graficos/esg.py
python3 esg/graficos/matriz_esg.py                                                 # matriz de riesgo ESG (Figura S5 del Anexo [Y])
python3 esg/graficos/esg_ahorro.py $V/Mallplaza_Valoracion_Consolidada_v4_ESG.xlsx          # cascada de ahorros ESG hasta el precio (Figura E6); carga supuestos y recalcula la hoja ESG_Ahorros
python3 macro/graficos/anexo_macro.py
python3 plantilla_informe/graficos/figuras_informe.py $FA $V4 $TOR
python3 informe_final/graficos/datos_resumen.py $V4 $V/Mallplaza_Figuras_Valuacion_recalc.xlsx

echo "== Documentos (Word)"
# Negocio, Industria y ESG (español) leen el texto de */documento/borrador/texto.json, extraído de los
# borradores del equipo con comun/borrador_a_json.py y luego corregido solo en datos; no se regenera aquí.
node negocio/documento/negocio.js
python3 industria/graficos/anexos_industria.py   # textos y tablas de los Anexos 3-7 (especificación del equipo) → figuras/datos_anexos.json
node industria/documento/industria.js
node valuacion/documento/valuacion.js      # MARCAR_NUEVO = true resalta lo nuevo (TIR y rotulado nominal/real); false lo quita
(cd valuacion/documento && for l in ES EN; do python3 /root/.claude/skills/synced/508cc6bc-ad76-422f-ada9-ad0fc4279625_0b2a2ef9-bf53-46a5-b8ed-72255cce463a/docx/scripts/office/soffice.py --headless --convert-to pdf --outdir . Mallplaza_Valuacion_$l.docx >/dev/null 2>&1; done; rm -f .~lock*)   # PDF de Valuación
node analisis_financiero/documento/analisis_financiero.js
node riesgos/documento/pagina_riesgos.js
node esg/documento/esg.js
node macro/anexo/anexo_macro.js
node plantilla_informe/documento/plantilla_informe.js
python3 informe_final/graficos/maquetar.py   # Word del informe consolidado (ES y EN)
python3 informe_final/typst/armar.py         # PDF final: cuerpo en Typst (gráficos al costado) + anexos
(cd "$ROOT" && python3 informe_consolidado/actualizar.py)     # copia los últimos PDF/Word a informe_consolidado/ (vigente + anterior)
(cd "$ROOT" && python3 contexto_llm/generar.py)             # cifras clave vigentes para el LLM (reemplaza lo anterior)
