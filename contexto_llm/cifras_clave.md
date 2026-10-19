# Cifras clave del informe (contexto para el LLM)

Generado 2026-10-19 19:30 por `contexto_llm/generar.py` a partir de los modelos **vigentes**. No editar a mano: se reemplaza completo en cada ejecución.
Si una cifra de un texto contradice esta tabla, está mal el texto. Cada fila indica su fuente (archivo!hoja!celda).

Reglas del informe: ver `CLAUDE.md` y `narrativa_y_coherencia.md`. Recomendación: VENDER.


## Precio y valor

| Cifra | Valor | Fuente |
|---|---|---|
| Precio de referencia (30-sep-26), CLP | CLP 3.555 | `Mallplaza_Valoracion_Consolidada_v4.xlsx!Supuestos!B8` |
| Acciones en circulación (millones) | 2.190 | `Mallplaza_Valoracion_Consolidada_v4.xlsx!Supuestos!B9` |
| Valor DCF por acción, CLP | CLP 1.859 | `Mallplaza_Valoracion_Consolidada_v4.xlsx!Precio_Objetivo!B12` |
| Valor por múltiplos (EV/EBITDA LTM ponderado) por acción, CLP | CLP 2.951 | `Mallplaza_Valoracion_Consolidada_v4.xlsx!Precio_Objetivo!B13` |
| Peso del DCF en el precio objetivo | 80% | `Mallplaza_Valoracion_Consolidada_v4.xlsx!Precio_Objetivo!B14` |
| PRECIO OBJETIVO (80% DCF / 20% múltiplos), CLP | CLP 2.077 | `Mallplaza_Valoracion_Consolidada_v4.xlsx!Precio_Objetivo!B15` |
| Potencial vs precio | −41,6% | `Mallplaza_Valoracion_Consolidada_v4.xlsx!Precio_Objetivo!B17` |
| Múltiplo EV/EBITDA LTM ponderado (70% Chile / 30% set común) | 14,4x | `Mallplaza_Valoracion_Consolidada_v4.xlsx!Precio_Objetivo!B5` |
| Valor terminal / EV (DCF) | 78,0% | `Mallplaza_Valoracion_Consolidada_v4.xlsx!DCF!C54` |
| Cap rate implícito del DCF (EBITDA LTM / EV) | 10,1% | `Mallplaza_Valoracion_Consolidada_v4.xlsx!DCF!C56` |

## Balance y operación

| Cifra | Valor | Fuente |
|---|---|---|
| Deuda financiera neta 30-jun-26 (CLP MM) | 1.285.740 | `Mallplaza_Valoracion_Consolidada_v4.xlsx!Supuestos!B10` |
| EBITDA LTM 2T26 (CLP MM) | 544.935 | `Mallplaza_Valoracion_Consolidada_v4.xlsx!Supuestos!B13` |
| DFN / EBITDA LTM | 2,36x | `calculado: Supuestos!B10 / B13` |
| Tasa de impuesto operacional de largo plazo | 27% | `Mallplaza_Valoracion_Consolidada_v4.xlsx!Supuestos!B59` |
| Crecimiento terminal nominal (g) | 3,0% | `Mallplaza_Valoracion_Consolidada_v4.xlsx!Supuestos!B88` |

## Costo de capital

| Cifra | Valor | Fuente |
|---|---|---|
| WACC nominal CLP 2S26E | 10,54% | `Mallplaza_Valoracion_Consolidada_v4.xlsx!WACC!C43` |
| WACC nominal CLP terminal | 10,41% | `Mallplaza_Valoracion_Consolidada_v4.xlsx!WACC!I43` |
| WACC real terminal | 7,20% | `Mallplaza_Valoracion_Consolidada_v4.xlsx!WACC!I44` |
| Ke nominal CLP terminal | 11,93% | `Mallplaza_Valoracion_Consolidada_v4.xlsx!WACC!I39` |
| Kd nominal después de impuestos terminal | 5,11% | `Mallplaza_Valoracion_Consolidada_v4.xlsx!WACC!I42` |
| Beta re-apalancada terminal | 1,14 | `Mallplaza_Valoracion_Consolidada_v4.xlsx!WACC!I37` |
| CRP ponderado por ingresos | 1,41% | `Mallplaza_Valoracion_Consolidada_v4.xlsx!WACC!C21` |
| Tasa libre de riesgo usada (UST 10 años) | 4,98% | `Mallplaza_Valoracion_Consolidada_v4.xlsx!WACC!C22` |
| WACC real implícita en la tasación (Nota 11, Chile) | 5,55% | `Mallplaza_Valoracion_Consolidada_v4.xlsx!WACC!C47` |

## Expansión

| Cifra | Valor | Fuente |
|---|---|---|
| Yield on cost de la expansión | 7,27% | `Mallplaza_Valoracion_Consolidada_v4.xlsx!Expansion!C37` |
| RONIC real después de impuestos (yield × (1 − t)) | 5,30% | `Mallplaza_Valoracion_Consolidada_v4.xlsx!Expansion!C38` |

## Retornos

| Cifra | Valor | Fuente |
|---|---|---|
| ROIC 2025 (NOPAT / capital promedio a valor razonable) | 7,34% | `Mallplaza_FA_Exhibits_v7.xlsx!E8_ROIC_RONIC!C21` |
| ROIC 2025 − WACC real (pp) | 0,15 | `Mallplaza_FA_Exhibits_v7.xlsx!E8_ROIC_RONIC!C33` |
| RONIC − WACC real (pp) | −2,16 | `Mallplaza_FA_Exhibits_v7.xlsx!E8_ROIC_RONIC!C41` |
| ROE normalizado 2025 | 9,7% | `Mallplaza_FA_Exhibits_v7.xlsx!E1_DuPont!H28` |
| ROE reportado 2025 | 37,5% | `Mallplaza_FA_Exhibits_v7.xlsx!E1_DuPont!H27` |

## Escenarios

| Cifra | Valor | Fuente |
|---|---|---|
| Precio objetivo escenario Bear, CLP | CLP 1.542 | `datos_anexos.json (réplica del DCF v4, Anexo 18)` |
| Bear: variación vs precio | −56,6% | `datos_anexos.json` |
| Precio objetivo escenario Base, CLP | CLP 2.077 | `datos_anexos.json (réplica del DCF v4, Anexo 18)` |
| Base: variación vs precio | −41,6% | `datos_anexos.json` |
| Precio objetivo escenario Bull, CLP | CLP 2.915 | `datos_anexos.json (réplica del DCF v4, Anexo 18)` |
| Bull: variación vs precio | −18,0% | `datos_anexos.json` |
| Precio objetivo escenario Tasadores, CLP | CLP 2.734 | `datos_anexos.json (réplica del DCF v4, Anexo 18)` |
| Tasadores: variación vs precio | −23,1% | `datos_anexos.json` |

## Mercado

| Cifra | Valor | Fuente |
|---|---|---|
| Mínimo de 52 semanas, CLP | CLP 2.470 | `datos_resumen.json (Bloomberg)` |
| Máximo de 52 semanas, CLP | CLP 4.700 | `datos_resumen.json (Bloomberg)` |
| Dividendo esperado 2027E, CLP por acción | 69,4 | `datos_resumen.json (modelo vigente, hoja Deuda)` |
| Capitalización bursátil (CLP MM) | 7.785.450 | `datos_resumen.json` |

## Monte Carlo

| Cifra | Valor | Fuente |
|---|---|---|
| Precio objetivo simulado (10.000 iteraciones), mediana, CLP | CLP 2.047 | `informe_final/graficos/cifras.json (cifras_informe.py; réplica del DCF v4)` |
| Simulaciones con precio objetivo bajo el precio (3.555) | 100,0% | `informe_final/graficos/cifras.json (cifras_informe.py; réplica del DCF v4)` |

## Pipeline

| Cifra | Valor | Fuente |
|---|---|---|
| Pipeline con fecha 2026–2028 (m²) | 170.200 | `informe_final/graficos/cifras.json (cifras_informe.py; réplica del DCF v4)` |
| Pipeline con fecha, inversión (USD MM) | 613 | `informe_final/graficos/cifras.json (cifras_informe.py; réplica del DCF v4)` |

## Mercado

| Cifra | Valor | Fuente |
|---|---|---|
| Precio objetivo de consenso a 12 meses (Bloomberg, 14 fuentes), CLP | CLP 4.111 | `informe_final/graficos/cifras.json (cifras_informe.py; réplica del DCF v4)` |

## Sensibilidades

| Cifra | Valor | Fuente |
|---|---|---|
| Δ valor DCF por acción con WACC −100 pb (CLP) | 430 | `informe_final/graficos/cifras.json (cifras_informe.py; réplica del DCF v4)` |
| Δ valor DCF por acción con WACC +100 pb (CLP) | -326 | `informe_final/graficos/cifras.json (cifras_informe.py; réplica del DCF v4)` |
| Δ valor DCF por acción con +1 pp de arriendo mismas tiendas 2027–28 (CLP) | 45,2 | `informe_final/graficos/cifras.json (cifras_informe.py; réplica del DCF v4)` |
| Δ valor DCF por acción con +1 pp de ventas de locatarios 2027–28 (CLP) | 2,6 | `informe_final/graficos/cifras.json (cifras_informe.py; réplica del DCF v4)` |

## Industria

| Cifra | Valor | Fuente |
|---|---|---|
| GLA Mallplaza (m²) | 2.360.135 | `Excel_DATA_INDUSTRY.xlsx!EF. OPERATIVA (CLPM2)!C6` |
| GLA Cenco Malls (m²) | 1.538.391 | `Excel_DATA_INDUSTRY.xlsx!EF. OPERATIVA (CLPM2)!C7` |
| GLA Parque Arauco (m²) | 1.340.500 | `Excel_DATA_INDUSTRY.xlsx!EF. OPERATIVA (CLPM2)!C8` |
| EBITDA por m² trimestral 2T26 Mallplaza (CLP) | 59.451 | `Eficiencia_EBITDA_GLA_2T26V2.xlsx!Eficiencia_EBITDA_GLA_2T26!F6` |
| EBITDA por m² trimestral 2T26 Cenco Malls (CLP) | 58.384 | `Eficiencia_EBITDA_GLA_2T26V2.xlsx!Eficiencia_EBITDA_GLA_2T26!F7` |
| EBITDA por m² trimestral 2T26 Parque Arauco (CLP) | 55.363 | `Eficiencia_EBITDA_GLA_2T26V2.xlsx!Eficiencia_EBITDA_GLA_2T26!F8` |

## Negocio

| Cifra | Valor | Fuente |
|---|---|---|
| A: GLA por país — Chile (%) | 61,9 | `Excel_bussines_para_armar_graficos.xlsx!Grafico 1, 2 tortas!B:D` |
| A: GLA por país — Perú (%) | 26,2 | `Excel_bussines_para_armar_graficos.xlsx!Grafico 1, 2 tortas!B:D` |
| A: GLA por país — Colombia (%) | 11,9 | `Excel_bussines_para_armar_graficos.xlsx!Grafico 1, 2 tortas!B:D` |
| B: Ingresos por fuente — Arriendo (%) | 81,4 | `Excel_bussines_para_armar_graficos.xlsx!Grafico 1, 2 tortas!B:D` |
| B: Ingresos por fuente — Parking (%) | 10,8 | `Excel_bussines_para_armar_graficos.xlsx!Grafico 1, 2 tortas!B:D` |
| B: Ingresos por fuente — Servicios comerciales y otros (%) | 5,9 | `Excel_bussines_para_armar_graficos.xlsx!Grafico 1, 2 tortas!B:D` |
| B: Ingresos por fuente — Publicidad (%) | 1,9 | `Excel_bussines_para_armar_graficos.xlsx!Grafico 1, 2 tortas!B:D` |
