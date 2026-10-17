# Narrativa y coherencia del informe: Mallplaza (Plaza S.A.)

CFA Institute Research Challenge 2026 · Recomendación: **VENDER** · Precio objetivo **CLP 2.077** · Precio al 30-sep-26 **CLP 3.555** · Retorno esperado **−41,6%** (−39,6% con dividendos)

Este archivo es la guía del equipo para que todas las secciones cuenten la misma historia con las mismas cifras. Cada cifra ancla indica de dónde sale; si una sección usa un número distinto, está mal la sección, no esta tabla.

---

## 1. La tesis en una frase

Mallplaza es un excelente negocio de arriendos (ingresos casi como un bono, margen EBITDA de 80% y balance holgado), pero a CLP 3.555 el mercado lo paga como si sus flujos fueran casi libres de riesgo. **El desacuerdo con el mercado no está en los flujos, está en la tasa**: el precio exige una WACC nominal de 7,60% (4,5% real), 281 pb menor que la nuestra, y ni siquiera crecer lo salva, porque cada m² nuevo rinde 7,3% real antes de impuestos y 5,3% después, menos que la WACC real de 7,2%.

## 2. El arco narrativo (qué aporta cada sección a la historia)

| # | Sección | Rol en la historia | Mensaje clave | Carpeta |
|---|---|---|---|---|
| 1 | Business Description | Presentar el activo | Operador líder de centros comerciales en Chile, Perú y Colombia; el 81,4% de los ingresos son arriendos (93% renta fija). GLA: 61,9% Chile, 26,2% Perú, 11,9% Colombia. Las tiendas por departamento bajaron de 26,7% a 19,0% del GLA (2020–2T26) | `secciones/negocio` |
| 2 | Industry | El negocio es bueno y defendible | Trío (Mallplaza, Cenco Malls, Parque Arauco) con 76% del GLA en Chile; ocupación alta; Porter favorable; Mallplaza es el más eficiente, pero por poco (EBITDA/m² trimestral 59,5k vs 58,4k Cenco y 55,7k PA, **brecha de solo 6,7%**) | `secciones/industria` |
| 3 | Financial Analysis | Separar el motor real del ruido contable | La utilidad reportada sigue al tasador (NIC 40); el FFO sigue al negocio. Crecimiento orgánico real ~3%; dos tercios del crecimiento 2025 fueron comprados (Perú). ROE normalizado 8–10%. Caja cubre el dividendo 2,6x. El AAA lo justifican los flujos, no el LTV | `secciones/analisis_financiero` |
| 4 | Valuation | Cuantificar la brecha | DCF CLP 1.859 (80%) y múltiplos CLP 2.951 (20%) → CLP 2.077. Ningún método alcanza el precio; el precio objetivo queda bajo el mínimo de 52 semanas (CLP 2.470). El mercado paga cap rate de 6,0%; nuestro DCF implica 10,1%. Incluso el escenario bull (todo sale bien) da CLP 2.915 | `secciones/valuacion` |
| 5 | Investment Risks | Qué podría romper la tesis | La tasa es el riesgo dominante (FR1: alta/alta). La operación mueve poco comparada con la tasa (WACC ±100 pb: −326/+430 CLP por acción; margen en su mínimo histórico: −190). La UF casi no mueve el valor (+1,1% con UF +3%). Riesgos no cuantificables (RR1, RR2) → análisis cualitativo | `secciones/riesgos` |
| 6 | ESG | Calidad del gobierno y sostenibilidad | Perfil "verde" local fuerte pero puntaje Bloomberg bajo (3,39); 1 de 9 directores independiente; interlocking con Falabella (53,05%) | `secciones/esg` |
| A | Anexos | Respaldo | WACC por período (Valuación), supuestos y estados resumidos (FA), macro 2022–2026, ESG | `*/anexo` |

**Abstract.** Va al inicio del informe consolidado (ES/EN, en `informe_final/documento/informe_final.js`): tesis en un párrafo con las cifras ancla y una franja con recomendación, precio objetivo, precio y potencial. El Resumen de inversión va al final.

**Cómo se encadenan:** Business e Industry dicen "es un gran negocio" → Financial Analysis dice "pero crece ~3% real y su utilidad está inflada por la tasación" → Valuation dice "y a este precio el mercado ya paga más de lo que esos flujos valen" → Risks dice "lo que puede cambiar la conclusión es la tasa, no la operación" → ESG agrega el matiz de gobierno (interlocking).

## 3. Cifras ancla (fuente única)

Las cifras vigentes con su celda de origen están en `contexto_llm/cifras_clave.md` (generado de los modelos) y las que calcula el informe (réplica del DCF, sensibilidades, escenarios, TIR) en `secciones/informe_final/graficos/cifras.json` (`cifras_informe.py`); los textos las leen con `secciones/comun/cifras.js`. Modelos vigentes: valoración **v4**, Exhibits de FA **v7**.

| Concepto | Valor | Fuente |
|---|---|---|
| Precio de referencia | CLP 3.555 (30-sep-26) | v4 Supuestos B8 (Bloomberg) |
| Rango 52 semanas | CLP 2.470 (10-oct-25) – 4.700 (27-abr-26) | Figuras_Valuacion V1 (Bloomberg) |
| Valor DCF por acción | CLP 1.859 | v4 DCF |
| Valor por múltiplos (EV/EBITDA UDM 14,4x) | CLP 2.951 | v4 Precio_Objetivo |
| Precio objetivo (80/20) | CLP 2.077 | v4 Precio_Objetivo |
| APV (chequeo) | CLP 1.895 (empresa sin deuda 1.630 + escudos 265) | v4 APV |
| WACC nominal 2S26E → terminal | 10,54% → 10,41% (real 7,32% → 7,20%) | v4 WACC |
| Ke terminal / Kd después de impuestos | 11,93% / 5,11% (cupón real 3,6% + spread) | v4 WACC |
| g terminal | 3,0% nominal (0% real) | v4 Supuestos |
| Retorno de la expansión (RONIC real, después de impuestos) vs WACC real | 5,3% vs 7,2% (antes de impuestos: 7,3%) | v4 Expansion; WACC |
| Valor terminal / EV | 78% | v4 DCF |
| EBITDA UDM 2T26 / margen | CLP 544.935 MM / 80,5% | v4 Supuestos |
| Margen EBITDA 2026E (año calendario) | 79,7% | v4 Forecast_Base |
| DFN 30-jun-26 / DFN/EBITDA | CLP 1.285.740 MM / 2,36x → 3,5x (2031) | v4 Supuestos; Deuda |
| Pipeline con fecha 2026–2028 | 170.200 m² (114.400 Chile + 55.800 Perú; incluye 46.500 m² de «otros proyectos Chile», apertura 2028), USD 613 MM | Exhibits v7 (D43 = 1, D44 = 2028) |
| GLA real de la industria en Chile | ≈ 4,3 millones de m² (GLA del trío sin oficinas 3,29 M ÷ 76%) | Industria v7 |
| Cap rate implícito: mercado / DCF / pares Chile / libro / tasadores | 6,0% / 10,1% / 6,4% / 7,5% / 5,55% real | v4 Precio_Objetivo; DCF |
| WACC que justifica el precio (nominal; desplaza toda la curva) | 7,60% (4,5% real), −281 pb | Figuras_Valuacion V6 (réplica) |
| TIR de la firma al precio (tasa única, FCFF): nominal / real | 7,61% / 4,48% (−280 / −272 pb vs WACC 10,41% / 7,20%) | v4 TIR_Implicita; TIR_Accionista |
| TIR del accionista a CLP 3.555: nominal / real | −1,8% / −4,7% (vs Ke 11,93% / 8,67%: −1.374 / −1.334 pb) | v4 TIR_Accionista |
| TIR del accionista al precio objetivo CLP 2.077 (control): nominal / real | 10,6% / 7,4% (−131 / −128 pb vs Ke) | v4 TIR_Accionista |
| TIR del proyecto nuevo (1 m² de expansión): nominal / real | 7,98% / 4,84% (−243 / −236 pb vs WACC) | v4 TIR_Proyecto |
| Escenarios (precio objetivo; bear = todo sale mal, base = central, bull = todo sale bien) | Bear 1.542 · Base 2.077 · Bull 2.915 · Tasadores 2.734; ponderado 30/60/10 (pesos de la matriz de riesgo): 2.001 (−43,7%) | réplica del DCF v4 (`escenarios_dcf.py`; Anexo 18) |
| Pesos de los escenarios | Bear 30% · Base 60% · Bull 10%: FR1 (tasas) probabilidad alta y OR1 (deterioro operacional) media en la matriz de riesgo; baja 10%, media 30%, alta 50% a 12 meses (supuesto del equipo); bear = FR1 y OR1 juntas (el menos probable de los dos: 30%), bull = sus contrarios (10%), base = el resto | `riesgos/graficos/matriz_datos.py` |
| Definición de escenarios | Bear: WACC +100 pb y cada palanca del tornado en su mínimo histórico; Bull: WACC −100 pb, cada palanca en su máximo e impuesto 23%; g 3,0% en todos | `escenarios_dcf.py` |
| Sensibilidades del valor DCF por acción (CLP) | WACC −100 / +100 pb: +430 / −326 · arriendo mismas tiendas +1 pp: +45 · margen +1 pp: +39 · ocupación +1 pp: +25 · ventas de locatarios +1 pp: +2,6 (renta variable 5,7% del ingreso) | réplica del DCF v4 |
| Consenso Bloomberg (12 meses) | CLP 4.111 (3.377–4.600; 14 fuentes: 1 comprar, 10 mantener, 2 vender) | Bloomberg |
| Dividendo esperado | ~CLP 69 por acción (2,0%) | v4 Deuda (2027E ordinario) |
| Mezcla de ingresos para el CRP | 67,2% Chile (CRP ponderado 1,41%) | v4 WACC |
| Deuda en UF | 79% | Release 2T26 |

## 4. Reglas de coherencia (aplican a todas las secciones)

1. **Moneda y unidades.** En español: "CLP MM" (millones) y "CLP mil MM" (miles de millones); "billones" solo para 10¹². En inglés: "CLP million / bn / trillion".
2. **UDM = LTM.** Unificar una sola sigla por idioma (español: UDM; inglés: LTM).
3. **Fechas.** Valoración al 30-jun-26; precio al 30-sep-26. Decirlo cuando se compare valor con precio.
4. **Colores.** Solo la paleta institucional de `secciones/comun/estilo.py` (rojo, celeste, amarillo, negro, azul; verde y morado solo para Parque Arauco y Cencosud). Sin grises ni tonos pastel en series.
5. **Figuras.** Sin título ni fuente dentro de la imagen del informe (van en el Word). Las versiones "_titulado" son solo para compartir fuera del informe.
6. **Margen 2026E.** 79,7% es el año calendario; 79,8% es solo el semestre 2S26E. No mezclarlos.
8. **Nominal y real, siempre.** Toda tasa (WACC, Ke, Kd, g, cap rate, TIR) dice si es nominal CLP o real (UF); unidas por Fisher (1 + nominal) = (1 + real)(1 + π), con π = 3,0% (Supuestos B34). Real y nominal nunca se mezclan en una misma frase sin rotular. La TIR de la firma (FCFF) se compara con la WACC; la del accionista (distribuciones + salida), con el Ke. El 7,60% (V6, desplaza toda la curva de WACC) y el 7,61% (TIR de tasa única) son dos lecturas distintas de la misma idea y ambas dan CLP 3.555: decir cuál se usa.
7. **Solo dos versiones.** De cada documento y cada gráfico se conservan solo la versión actual y la penúltima; las anteriores se eliminan junto con el código Python que solo servía para generarlas (siguen en el historial de git). Las variantes de una misma versión (ES/EN, docx/pdf, normal/"_titulado") no cuentan como versiones. Los modelos Excel quedan fuera de esta regla.

### Versiones vigentes (actual → penúltima)

| Elemento | Actual | Penúltima |
|---|---|---|
| Página de Riesgos | `riesgos/documento/Mallplaza_Investment_Risks_*.docx` | `plantilla_informe/documento/Investment_Risks_template.docx` |
| Figuras de Riesgos (matriz, tornado, mapa de calor, UF) | `riesgos/.../es_/en_fig23–26` | `plantilla_informe/.../fig1–fig4` |
| Football field | `valuacion/.../es_/en_V1_football.png` | `valuacion/.../football_field.png` |
| Supuestos fuertes de la valoración (Figura 29, cascada + tabla) | `valuacion/.../es_/en_V13_supuestos_cascada.png` | `valuacion/.../es_/en_V4_capital_trabajo.png` |
| ESG: scorecard (E1) y energía/residuos (E3) con pares de Bloomberg y desglose de lo valorizado (CFA beta 4.1) | `esg/graficos/figuras/es_/en_E1_scorecard.png`, `..._E3_energia_residuos.png` | `esg/graficos/figuras/anterior/` |
| ESG: ahorros ESG y menor costo de capital hasta el precio (E6, escenario) y matriz de riesgo ESG (S5) | `esg/graficos/figuras/es_/en_E6_ahorro_esg.png`, `..._AE5_matriz_esg.png`; modelo `valuacion/modelos/Mallplaza_Valoracion_Consolidada_v4_ESG.xlsx` | (primera versión) |
| Borrador ESG del equipo | `esg/documento/borrador/CFA_ESG_PLAZA_beta_4_1.docx` (+ texto.json / texto_en.json) | `CFA_ESG_PLAZA_beta_3.docx` |
| Industria: Porter (ind6) y EBITDA/m² (ind8) con los datos nuevos del equipo | `industria/graficos/figuras/es_/en_ind6_porter.png`, `..._ind8_ebitda_m2.png` (datos: `industria/modelos/Porter_Nuevo.xlsx`, `Eficiencia_EBITDA_GLA_2T26V2.xlsx`) | `industria/graficos/figuras/anterior/` |
| Anexos de Industria 3 a 7 (formato del informe: tablas y gráficos ind2, ind1, ind7) | `industria/anexo/Mallplaza_Industria_*_Anexo.docx/.pdf` (especificación: `industria/modelos/Anexos_Industry_especificacion.xlsx`; referencia: `industria/documento/borrador/Industria_Mallplaza_v7_equipo.docx`) | (primera versión) |
| Sensibilidad a tasas largas e inflación | `valuacion/.../es_/en_V11_sens_tasas_inflacion.png` (grilla) y `..._V12_mix_tasas_inflacion.png` (mix, heat map); modelo `valuacion/modelos/Mallplaza_Sensibilidad_Tasas_Inflacion.xlsx` | `..._V12_mix_lineas_tasas_inflacion.png` (mix en líneas, versión anterior) · V12 (heat map) está en Riesgos, Figura 26, que reemplaza a `riesgos/graficos/figuras/es_/en_fig26_uf.png` (queda como versión anterior) |
| Página de Valuación (Word y PDF) | `valuacion/documento/Mallplaza_Valuacion_*` (con la TIR y lo nuevo marcado) | `valuacion/documento/anterior/Mallplaza_Valuacion_*` |
| Modelos de valoración y de FA | `valuacion/modelos/Mallplaza_Valoracion_Consolidada_v4.xlsx` (+ `_v4_ESG.xlsx`) y `analisis_financiero/modelos/Mallplaza_FA_Exhibits_v7.xlsx` (los construye `valuacion/graficos/modelo_v4.py`: pipeline con fecha 170.200 m², «Otros proyectos Chile» activados) | v3 (+ `_v3_ESG`) y Exhibits v6 (modelos Excel fuera de la regla de dos versiones) |
| Resumen de inversión: trayectoria de la acción (Figura R1 = 32), escenarios con el oso y el toro (R2 = 33) y escalera DCF → precio → consenso (R3 = 34) | `valuacion/.../es_/en_V14_trayectoria.png`, `..._V10_escenarios.png` (ilustración `iconos/oso.png`, `toro.png`), `..._V15_escalera.png`; scripts `resumen_figuras.py`, `escenarios.py` | V10 con íconos vectoriales (boceto no publicado) |
| Texto del cuerpo de Valuación, FA, Resumen y Abstract | `valuacion/documento/texto_valuacion.js`, `analisis_financiero/documento/texto_fa.js`, `informe_final/documento/texto_resumen.js` (cifras desde `cifras.json`) | texto anterior (en el historial de git) |
| Anexos de Valuación 15–19 y de FA 9–14 | `valuacion/documento/anexo_valuacion.js`, `analisis_financiero/documento/anexo_fa.js` | (primera versión) |
| Industria: cuota de mercado por país (Figura 2, horizontal) | `industria/.../es_/en_ind5_market_share_paises.png` | `..._ind5b_market_share_paises_vertical.png` |
| FA: perfil de vencimientos (Figura 18) y capex incremental (Figura 19) a ancho completo | `analisis_financiero/graficos/figuras/es_/en_FA6_vencimientos_ancha.png`, `..._FA7_capex_ancha.png` | `..._FA6_vencimientos.png`, `..._FA7_capex.png` |

Eliminadas por esta regla: figuras sueltas de Riesgos (`matriz_riesgo`, `tornado_operacional`, `heatmap_sensibilidad_wacc`, `reajuste_uf`) y la grilla `sensibilidad_wacc` de Valuación (modelo VF), con `matriz_riesgo.py` y `sensibilidad_wacc.py`. `tornado_operacional.py`, `heatmap_sensibilidad_wacc.py` y `reajuste_uf.py` quedan solo como bibliotecas de cálculo de `pagina_riesgos.py`.

## 5. Inconsistencias conocidas y decisiones pendientes

| # | Tema | Dónde | Estado |
|---|---|---|---|
| 1 | Numeración de figuras y anexos | Todas | Resuelto en el informe consolidado (`informe_final/`): figuras 1–34 y anexos 1–20 de corrido; cada sección conserva su numeración propia y `informe_final.js` la traduce con un mapa por sección |
| 2 | Kd: cupón (3,6%) vs tasa efectiva de la última emisión (4,33%) | Valuación vs FA | Con 4,33% la WACC sube 12 pb y el precio objetivo baja a ~CLP 2.042 (el texto usa el valor recalculado en cifras.json). Ideal: YTM de Bloomberg (YAS) |
| 3 | IRSA dentro del set de pares | Valuación | Incluida (modelo v4). Si se excluye: múltiplo 14,0x, precio objetivo ~CLP 2.059 (por recalcular) |
| 4 | Fecha de valoración (30-jun) vs precio (30-sep) | Valuación | Llevar el DCF al 30-sep → ~CLP 2.119. No aplicado |
| 5 | Dividendo 2S26 (~CLP 27/acción) y fecha ex-dividendo | Valuación | Confirmar con Bloomberg (DVD) |
| 6 | Gran Plaza (9,4% en la Figura 26; −CLP 10,7 por acción en el DCF) | Valuación | Nadie del equipo pudo verificar el dato |
| 7 | Brecha de cap rate: 490 pb (Estructura_Valuacion) vs 413 pb (V5) | Valuación | La figura usa 413 pb |
| 8 | Tramos de inflación y moneda de la Figura 3 de FA | FA | Derivados de los totales del borrador; reemplazar si hay valores exactos |
| 9 | Consenso de precio objetivo (ANR) | Valuación / Resumen | Resuelto: CLP 4.111 (rango 3.377–4.600; 14 fuentes) |
| 10 | Fuentes políticas de RR2 (Chile, Perú, Colombia) | Riesgos | Tomadas de resúmenes de búsqueda; verificar cada URL antes de entregar |
| 12 | Escenarios: bear = todo sale mal, bull = todo sale bien | Valuación / Resumen / Riesgos | Resuelto: bear 1.542, base 2.077, bull 2.915 con la réplica del DCF v4 (todas las palancas del tornado; WACC ±100 pb; Anexo 18) |
| 13 | Pipeline USD 613 MM / 170.200 m² (texto) vs 123.700 m² / USD 445 MM con fecha (modelo) | Industria / Valuación / Riesgos | Resuelto en el modelo v4: se activan «Otros proyectos Chile» (46.500 m², apertura 2028) y el modelo coincide con el texto |
| 14 | WACC implícita 7,60% (V6) vs TIR de tasa única 7,61% (v4) | Valuación | Resuelto en el texto con una nota bajo la tabla de la TIR (−281 pb vs −280 pb) |
| 15 | Control del modelo: comprar al valor DCF (CLP 1.859) da una TIR de accionista de 13,5% nominal, 157 pb sobre el Ke terminal (11,93%) | Valuación | Sin resolver. Probables causas: costo de la deuda del flujo (cupón + reajuste, ~5,4% antes de impuestos) menor que el Kd de la WACC (6,7%); minoritarios a libro; dividendos con rezago |
| 16 | Rótulo nominal/real en anexos y figuras de Valuación | Valuación | Resuelto en los anexos 15–19 y en el texto; revisar nuevas figuras al agregarlas |
| 11 | Rebaja del impuesto corporativo a 23% (Ley de Reconstrucción) | Riesgos / Valuación | No incluida en el modelo (+5,2% al precio objetivo si se aplica; es el impuesto de 23% del escenario bull) |

## 6. Preguntas que probablemente hará el jurado

- ¿Por qué su WACC (10,4%) es tan superior a la tasa de los tasadores (5,55% real) y a la del mercado (cap rate 6,0%)? → Ke en USD + CRP; con la tasa de los tasadores el precio objetivo (CLP 2.734) igual queda 23% bajo el precio.
- ¿Por qué no suponer más crecimiento? → RONIC 5,3% < WACC real 7,2%: crecer resta valor.
- ¿No están siendo pesimistas? → El consenso espera un EBITDA NTM solo ~3% sobre el nuestro.
- ¿Por qué pesan 20% los múltiplos y 70% Chile? → Dos pares chilenos; la brecha con Brasil es de tasas (spreads de 385 vs 302 pb).
- ¿Qué tendría que pasar para equivocarse? → Compresión adicional de cap rates o una prima estructural de los activos Tier A.

## 7. Mapa de archivos

```
CLAUDE.md                     reglas del proyecto (idioma, versiones, orden)
narrativa_y_coherencia.md     esta guía
pendientes.md                  lista única de pendientes del proyecto (abrir y cerrar ahí; no dejar pendientes sueltos en otros archivos)
generar_todo.sh               regenera gráficos, documentos por sección y el informe consolidado
informe_consolidado/          ENTREGABLES FINALES: vigente/ (ES/EN, PDF y Word actuales) y anterior/ (penúltima); actualizar.py
contexto_llm/                 contexto del LLM: cifras_clave.md/.json y archivos_vigentes.md, regenerados de los modelos vigentes (generar.py)
.claude/agents/contexto-llm.md  agente que sincroniza contexto_llm/ e informe_consolidado/ con los últimos archivos
data/                         datos de origen y diccionario (raw, processed, sources, logs)
secciones/
  comun/                      estilo.py (paleta), formato.js (formato Word), cifras.js (lee cifras.json; formatos ES/EN), borrador_a_json.py, logos/
  <sección>/                  negocio, industria, esg, analisis_financiero, valuacion, riesgos, macro, plantilla_informe
    documento/                constructor .js + Word/PDF de la sección (ES/EN)
    documento/anterior/       (valuacion) penúltima versión del Word/PDF, sin la TIR
    documento/borrador/       borrador del equipo (.docx) y su texto en JSON (texto.json / texto_en.json)
    anexo/                    anexo de la sección (ES/EN)
    graficos/                 scripts de Python de los gráficos
    graficos/figuras/         cada gráfico (es_/en_, y _titulado para compartir fuera del informe); datos_*.json (valuacion: datos_wacc, datos_anexos, datos_tir)
    valuacion/graficos/       modelo_v4.py, escenarios_dcf.py (réplica del DCF por escenario), actualizar_figuras.py, resumen_figuras.py (V14, V15), iconos/ (oso, toro)
    modelos/                  Excel de entrada de la sección (valuacion: v4 vigente y v3; los construye graficos/modelo_v4.py; analisis_financiero: Exhibits v7 y v6)
  informe_final/              informe consolidado
    documento/                informe_final.js + cuerpo.js → Word final (ES/EN); PDF final (lo escribe typst/armar.py)
    documento/maquetacion/    intermedios de la maquetación del Word (plan, posiciones, escalas, saltos)
    graficos/                 cifras_informe.py (cifras.json: única fuente de números de los textos), datos_resumen.py (datos clave), maquetar.py (Word), aplicar_correcciones.py
    documento/texto_resumen.js  Abstract y Resumen de inversión (ES/EN) calculados desde cifras.json
    typst/                    informe.typ (diseño del PDF A4: 10 páginas de cuerpo + 10 de anexos), exportar.js (textos → JSON), armar.py (PDF final)
    typst/datos/              cuerpo_<lang>.json exportado para Typst
```

**Entregables finales:** `informe_consolidado/vigente/Mallplaza_Informe_CFA_{ES,EN}.pdf` (para entregar; cuerpo en Typst + anexos) y `.docx` (para editar). `secciones/informe_final/` es la carpeta de trabajo que los arma.

**Escenarios con el oso y el toro:** `valuacion/graficos/escenarios.py` genera `es_/en_V10_escenarios.png` (y `.svg` editable) con la ilustración del equipo (`valuacion/graficos/iconos/oso.png` y `toro.png`, recortadas de la imagen suministrada): bear 1.542 (oso), base 2.077 y bull 2.915 (toro) frente al precio de cierre (3.555). Es la Figura R2 (33) del Resumen de inversión; cifras de `datos_anexos.json` (Anexo 18). Los íconos vectoriales anteriores (Game-icons.net, CC BY 3.0) quedan en `iconos/` como respaldo del script; si se usan, citar «Íconos: Game-icons.net (CC BY 3.0)».
