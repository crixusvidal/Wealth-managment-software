#!/usr/bin/env bash
# Regenera todos los gráficos y documentos del informe a partir de los modelos de cada sección.
# Uso: bash generar_todo.sh        (requiere python3 con matplotlib/openpyxl y node con docx)
set -euo pipefail
ROOT="$(cd "$(dirname "$0")" && pwd)"
cd "$ROOT/secciones"
export NODE_PATH="${NODE_PATH:+$NODE_PATH:}/opt/node22/lib/node_modules:/opt/node-tools/node_modules"
V=valuacion/modelos; F=analisis_financiero/modelos; R=riesgos/modelos
V2=$V/Mallplaza_Valoracion_Consolidada_v2.xlsx
V3=$V/Mallplaza_Valoracion_Consolidada_v3.xlsx   # = v2 + hojas TIR_Implicita y TIR_Accionista (TIR nominal/real); las demás hojas son idénticas a v2
FA=$F/Mallplaza_FA_Exhibits_v6.xlsx
TOR=$R/Mallplaza_Tornado_Sensibilidad_Operacional_2026_2028.xlsx

echo "== Gráficos"
python3 negocio/graficos/negocio.py negocio/modelos/Excel_bussines_para_armar_graficos.xlsx
python3 negocio/graficos/mapa_sudamerica.py   # Figura 1 (mapa de Sudamérica; requiere geopandas<1.0)
python3 industria/graficos/industria.py industria/modelos/Excel_DATA_INDUSTRY.xlsx
python3 industria/graficos/cuota_mercado.py $V/Comparables_Mallplaza_CFA_BBG_valoracion.xlsx
python3 valuacion/graficos/valuacion.py $V/Mallplaza_Figuras_Valuacion_recalc.xlsx - $V2
python3 valuacion/graficos/anexos_valuacion.py $V2 $V/Mallplaza_Figuras_Valuacion_recalc.xlsx $FA $TOR
python3 valuacion/graficos/tir_valuacion.py $V3   # TIR firma/accionista (nominal y real) → figuras/datos_tir.json (lo lee valuacion.js)
python3 valuacion/graficos/sensibilidad_tasas.py $V3 $V/Mallplaza_Tasas_Historico.xlsx   # sensibilidad a tasas largas e inflación → modelos/Mallplaza_Sensibilidad_Tasas_Inflacion.xlsx + figuras V11/V12 (recalcular el Excel con LibreOffice)
python3 valuacion/graficos/football_field.py $V2
python3 valuacion/graficos/escenarios.py      # boceto: escenarios bear/base/bull con oso y toro
python3 valuacion/graficos/tabla_kd_caprates.py $V/Comparables_Mallplaza_CFA_BBG_valoracion.xlsx
python3 valuacion/graficos/tabla_multiplos_implicitos.py $FA
python3 analisis_financiero/graficos/analisis_financiero.py
python3 riesgos/graficos/pagina_riesgos.py $FA $V2 $TOR
python3 esg/graficos/esg.py
python3 macro/graficos/anexo_macro.py
python3 plantilla_informe/graficos/figuras_informe.py $FA $V2 $TOR
python3 informe_final/graficos/datos_resumen.py $V2 $V/Mallplaza_Figuras_Valuacion_recalc.xlsx

echo "== Documentos (Word)"
# Negocio, Industria y ESG (español) leen el texto de */documento/borrador/texto.json, extraído de los
# borradores del equipo con comun/borrador_a_json.py y luego corregido solo en datos; no se regenera aquí.
node negocio/documento/negocio.js
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
