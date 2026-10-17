# Lista de pendientes del proyecto (informe CFA Mallplaza)

Una sola lista para todo lo que sigue abierto. Al cerrar un pendiente, pasarlo a «Cerrados» con la fecha. Detalle de las inconsistencias de cifras: `narrativa_y_coherencia.md` (sección 5).

## Datos que faltan (Bloomberg y fuentes)
- [ ] Consenso de precio objetivo (ANR): dato de alto impacto para el texto de Valuación (inconsistencia 9).
- [ ] YTM del bono UF de Mallplaza (YAS): reemplaza el Kd por el cupón de 3,6% y la tasa efectiva de 4,33% (inconsistencia 2; con 4,33% el precio objetivo baja a CLP 2.042).
- [ ] Fechas y monto del dividendo 2S26 (DVD, ~CLP 27 por acción) (inconsistencia 5).
- [ ] Verificar el dato de Gran Plaza (cap rate 9,4% en la Figura 33) (inconsistencia 6).
- [ ] Verificar una por una las URL de las fuentes políticas de RR2 en Riesgos (inconsistencia 10).
- [ ] Valores exactos de los tramos de inflación y moneda de la Figura 3 de FA, hoy derivados de los totales del borrador (inconsistencia 8).

## Valuación
- [ ] Rotular nominal/real en los Anexos [B], [C], [E] y [F] y en las Figuras 27 y 34 (WACC ±100 pb, g 2,5%, WACC implícita) (inconsistencia 16).
- [ ] Comparar los dos cap rates sobre la misma base (6,0% del mercado y 5,55% real de los tasadores) con una conversión explícita; hoy el texto solo los compara en real con la TIR. Falta espacio en las 2 páginas.
- [ ] Resolver la inconsistencia del modelo: al valor DCF de CLP 1.857 el accionista gana 13,5% nominal, 160 pb sobre el Ke de 11,93% (inconsistencia 15). Revisar costo de la deuda en los flujos frente al Kd de la WACC, minoritarios a libro y rezago de los dividendos.
- [ ] TIR del accionista por escenario (bear y bull): exige rehacer la trayectoria de deuda de cada caso.
- [ ] Llevar el DCF al 30-sep-26 (~CLP 2.118) o dejar explícito que no se aplica (inconsistencia 4).
- [ ] Rebaja del impuesto corporativo a 23%: no está en el modelo (+5,1% al precio objetivo) (inconsistencia 11).
- [ ] Pipeline: texto con USD 613 MM y 170.200 m² frente al modelo con 123.700 m² y USD 445 MM (inconsistencia 13); agregar el pipeline completo después, por decisión del equipo.
- [ ] Modelo de sensibilidad a tasas largas e inflación: hecho en `Mallplaza_Sensibilidad_Tasas_Inflacion.xlsx`; falta decidir dónde entra en el informe y validar los supuestos de traspaso (ver su hoja Supuestos).

- [ ] Escudo tributario (goodwill tributario y pérdidas): el modelo lo agota en 2029 (impuesto en caja 16,8% → 27%), un ahorro total de CLP 134.173 MM (CLP 61 por acción sin descontar). Verificar contra la Nota 13 de los EEFF (saldo del impuesto diferido activo por pérdidas tributarias y goodwill tributario por amortizar) que ese saldo respalde esa cifra; hoy es un supuesto [S] del Forecast («verificar Nota 13»). Los EEFF no están en el repo y esta sesión no pudo abrirlos.
- [ ] Pasar al Excel (hoja TIR_Implicita o TIR_Accionista) las TIR del activo al precio objetivo y al valor DCF, hoy calculadas solo con `tir_valuacion.py`.

- [ ] TIR del proyecto nuevo (4,62% real, 7,76% nominal): validar los supuestos del flujo de un m² (impuesto de largo plazo de 27%, mantención de 5,5% y capital de trabajo de 15,6% sobre los ingresos del proyecto); para igualar la WACC real los ingresos por m² tendrían que ser 1,61x los del modelo (yield on cost de 6,9% a 11,1%).

- [ ] ESG, ahorros (Figura E6): los precios unitarios (energía CLP 110/kWh, agua CLP 1.500/m³, residuos CLP 10.000/t) y el 25% del ahorro que retiene Mallplaza son supuestos ilustrativos cargados en la hoja ESG_Ahorros; reemplazarlos con los datos reales (precio de la energía, tarifa de agua, costo por tonelada y % no traspasado al gasto común, Nota 2.20). Faltan además el monto elegible de la Ley I+D y el spread de un crédito sostenible.
- [ ] ESG, Falabella: la matriz de riesgo del borrador decía 53,05% y la tabla y el texto 53,10%; se usó 53,10% en todo. Confirmar con la Memoria.
- [ ] ESG, track record: se usó la versión neutral de la beta 4.1; el Punteo del equipo trae «referencias» más elogiosas (p. ej. «arquitecto» de Mallplaza, «libre de multas y brechas») que chocan con las sanciones de 2025; decidir si se incorporan.
- [ ] ESG, idioma: la beta 4.1 trae los títulos y subtítulos del texto en español en inglés (ENVIRONMENTAL, SOCIAL, GOVERNANCE) y etiquetas en inglés dentro del párrafo; se dejó tal cual.

## Texto del informe
- [ ] Reescribir Riesgos de inversión (hoy con +3,9% por impuesto y 9 notas pendientes) y conectar la narrativa entre secciones.
- [ ] Resumen de inversión: redactar el texto (hoy solo está la tabla de datos clave).
- [ ] ESG: diferencias de texto frente a la versión anterior (Cadem 31/385, Ley I+D 55%, mujeres 53,2%).
- [ ] Comparación del ROIC (7,3%) con la WACC real, no con la nominal, en todos los textos.

## Gráficos y diseño
- [ ] Ilustraciones del oso (amarillo) y del toro (negro) hechas en ChatGPT: guardarlas en `secciones/valuacion/graficos/iconos/` como `oso.png` y `toro.png` y regenerar `escenarios.py`.
- [ ] Decidir dónde va la figura de escenarios (Anexo 18 o junto al texto de escenarios en Valuación); citar «Íconos: Game-icons.net (CC BY 3.0)» si se usan los SVG.
- [ ] Word consolidado: sin la portada de Mallplaza ni el mismo ajuste de gráficos que el PDF de Typst; al armarlo avisó que las Figuras 32 y 33 quedaron sin ubicar. El PDF es el entregable.
- [ ] Regenerar el informe consolidado con el texto nuevo de Valuación (TIR) y los cambios de Valuación posteriores (hoy el consolidado importa `valuacion.js`; el PDF Typst no lleva la tabla de TIR).

## Mantenimiento
- [ ] `contexto_llm/generar.py` pasará a leer el modelo v3 en su próxima corrida: revisar que las 56 cifras sigan igual.
- [ ] Conservar solo dos versiones: revisar `secciones/valuacion/documento/anterior/` cuando salga la próxima versión de Valuación.

## Cerrados
- [x] 17-oct-26: portada con colores de Mallplaza, márgenes casi nulos, gráficos al costado en el PDF.
- [x] 17-oct-26: TIR (firma y accionista, nominal y real) en Valuación y modelo v3.
