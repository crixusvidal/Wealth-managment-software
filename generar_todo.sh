#!/usr/bin/env bash
# Regenera todos los gráficos y documentos del informe a partir de los modelos de cada sección.
# Uso: bash generar_todo.sh        (requiere python3 con matplotlib/openpyxl y node con docx)
set -euo pipefail
cd "$(dirname "$0")/secciones"
export NODE_PATH="${NODE_PATH:+$NODE_PATH:}/opt/node22/lib/node_modules:/opt/node-tools/node_modules"
V=valuacion/modelos; F=analisis_financiero/modelos; R=riesgos/modelos
V2=$V/Mallplaza_Valoracion_Consolidada_v2.xlsx
FA=$F/Mallplaza_FA_Exhibits_v6.xlsx
TOR=$R/Mallplaza_Tornado_Sensibilidad_Operacional_2026_2028.xlsx

echo "== Gráficos"
python3 negocio/graficos/negocio.py negocio/modelos/Excel_bussines_para_armar_graficos.xlsx
python3 industria/graficos/industria.py industria/modelos/Excel_DATA_INDUSTRY.xlsx
python3 industria/graficos/cuota_mercado.py $V/Comparables_Mallplaza_CFA_BBG_valoracion.xlsx
python3 valuacion/graficos/valuacion.py $V/Mallplaza_Figuras_Valuacion_recalc.xlsx - $V2
python3 valuacion/graficos/football_field.py $V2
python3 valuacion/graficos/sensibilidad_wacc.py $V/Mallplaza_Valoracion_Consolidada_VF.xlsx
python3 valuacion/graficos/tabla_kd_caprates.py $V/Comparables_Mallplaza_CFA_BBG_valoracion.xlsx
python3 valuacion/graficos/tabla_multiplos_implicitos.py $FA
python3 analisis_financiero/graficos/analisis_financiero.py
python3 riesgos/graficos/pagina_riesgos.py $FA $V2 $TOR
python3 riesgos/graficos/tornado_operacional.py $FA $V2 $TOR
python3 riesgos/graficos/heatmap_sensibilidad_wacc.py $V2
python3 riesgos/graficos/reajuste_uf.py riesgos/graficos/figuras/reajuste_uf
python3 esg/graficos/esg.py
python3 macro/graficos/anexo_macro.py
python3 plantilla_informe/graficos/figuras_informe.py $FA $V2 $TOR

echo "== Documentos (Word)"
node valuacion/documento/valuacion.js
node analisis_financiero/documento/analisis_financiero.js
node riesgos/documento/pagina_riesgos.js
node esg/documento/esg.js
node macro/anexo/anexo_macro.js
node plantilla_informe/documento/plantilla_informe.js
