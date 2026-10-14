# Revisión de narrativa y coherencia (agente revisor)

Informe generado por un agente revisor a partir de `narrativa_y_coherencia.md`, el texto de todas las secciones y el modelo v2. No modificó archivos. Las acciones quedan pendientes de decisión del equipo.

## Hallazgos principales

### 1. Riesgos contradice la tesis (el más grave)
Riesgos y el anexo macro presentan el alza de tasas como el principal riesgo a la baja, pero para Valuación es el catalizador del VENDER. Para un VENDER, los riesgos relevantes son los que podrían romperlo:
- baja de tasas o compresión de cap rates;
- prima estructural de los activos Tier A;
- OPA o reorganización de Falabella (53%);
- recompras con los CLP 1,8 billones;
- free float acotado y dificultad para vender corto.

Valuación ya remite a Riesgos para la prima Tier A y los cap rates, pero Riesgos no los trata.

### 2. Cifras a unificar entre secciones
| Tema | Dónde | Corregir a |
|---|---|---|
| Rating | Riesgos (FR1) dice "AA+/BBB" | AAA (Feller Rate, 25-sep-26), como dice FA |
| Duración de la deuda | Riesgos (FR1) dice ~6,7 años | 6,6 años (FA, Nota 14) |
| Pipeline | Valuación y RR2 dicen USD 613 MM y 170.200 m² | El modelo usa 123.700 m² con fecha y USD 445 MM restantes; 613/170.200 solo para el costo por m² |
| Margen "mínimo/máximo histórico" (75,8% / 81,3%) | Valuación y Riesgos (OR1) | Rotularlos como trimestrales (4T23–2T26); la serie anual es 76,2–80,5% |
| Macro de Chile | Valuación: inflación 4,1% y PIB 1,6% | Una sola fuente con el anexo macro: IPC 3,9% y PIB 0,25–0,75% (IPoM jun-26) |
| ROIC 7,3% vs WACC | FA lo compara con la WACC nominal (10,4%) | Compararlo con la WACC real (7,2%), contra la que queda encima; corregir la frase sobre el valor en libros |
| "Ningún método justifica el precio" | Valuación | "Ninguno de los métodos que usamos": el múltiplo de Parque Arauco (3.967) y el P/E (8.644) superan el precio |
| Capitalización y deuda bruta | Hoja Supuestos y anexo WACC: 7.905.900 (sin fecha) y 1.579 (dic-25) | Fechar o actualizar (precio 3.555 → 7.785; deuda bruta jun-26: 1.633) |
| Apalancamiento objetivo | El D/E objetivo usa 3,5x el EBITDA UDM, no el de 2031 | Revisar; efecto ≈ −15 pb en la WACC (por verificar) |
| Nombre del par | "Cencosud Shopping" en FA, Valuación y WACC; "Cenco Malls" en Industria | Un solo nombre |
| Figuras de Industria | Solo en inglés y con título y fuente dentro del PNG | Hacerlas es_/en_ sin título ni fuente |
| UST 10Y en 5,22% "máximo desde 2002" | Valuación, Riesgos (FR1), anexo macro | Por verificar: en jun-2007 tocó ~5,26%, sería máximo desde 2007 |
| ESG | Puntaje Bloomberg 3,39 "bajo la media" vs "líder de la industria global" (HCFC) | Aclarar a qué se refiere cada dato |

### 3. Huecos que atacaría el jurado
- **Beta de 1,14 para un negocio "tipo bono"**, apoyada en pares brasileños. Defensa disponible: los dos pares chilenos desapalancan a 0,94. Mostrar la beta propia de Mallplaza.
- **Rf UST + CRP (6,80%) vs bono en pesos (5,43%).** La diferencia equivale a ~107 pb de WACC, más de un tercio de la brecha de 285 pb. Justificarla con el EMBI o los CDS de Chile.
- **El valor terminal implica 8,9x EBITDA**, bajo el par más barato. Contrastarlo con transacciones privadas, por ejemplo el múltiplo de la compra de Open Plaza Perú (2024).
- **"Crecer destruye valor" solo vale con nuestra WACC.** El yield on cost (6,9%) supera el cap rate de mercado. Mostrar el DCF sin expansión.
- **Catalizador y plazo.** Explicar por qué la acción subió 90% en 12 meses y qué cerraría la brecha.
- **El ESG no conecta con el valor.** Falta el porcentaje de rentas que viene de Falabella y cómo pesa el gobierno en la tesis.
- **"Por qué Mallplaza y no el sector"** (29% de prima sobre Cencosud con solo 1,9% más de EBITDA/m²) debería ser un pilar del Resumen de inversión y de Industria.

## Las 5 mejoras de mayor impacto
1. Reescribir Riesgos como riesgos para el VENDER y alinear el anexo macro con eso.
2. Redactar Negocio, Industria y el Resumen de inversión, con figuras es_/en_ sin título interno.
3. Blindar la tasa: beta chilena y beta propia, CRP frente al spread de mercado, y múltiplo de salida frente a transacciones.
4. Unificar las cifras y términos de la tabla anterior.
5. Un puente final del precio objetivo en un solo cuadro (DCF al 30-sep, Kd 4,33%, IRSA, impuesto 23%, consenso) y un catalizador fechado.

## Cifras verificadas sin diferencias
Coinciden con el modelo v2:
- valores por acción: precio 3.555, DCF 1.857, múltiplos 2.951, PO 2.076 y APV 1.894;
- retornos: −41,6% y −39,7% con dividendo;
- puente del DCF 583 + 1.884 − 610;
- valor terminal: 76% del EV;
- cap rates 10,1% (DCF) y 6,0% (mercado);
- WACC: 10,54% → 10,41% nominal (7,32% → 7,20% real); beta 0,94 / 1,14; CRP 1,41%; Ke 11,93%;
- RONIC 5,0%;
- distribuciones adicionales 1,79 billones;
- acciones 2.190 MM;
- P/BV 1,61x y FFO/capitalización 5,5%;
- ponderación por país 1.916–2.144 y spread Chile de 385 pb;
- sensibilidad: −6,5% / −12,3%.
