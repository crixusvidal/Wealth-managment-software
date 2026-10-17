# Lista de pendientes del proyecto (informe CFA Mallplaza)

Una sola lista para todo lo que sigue abierto. Al cerrar un pendiente, pasarlo a «Cerrados» con la fecha. Detalle de las inconsistencias de cifras: `narrativa_y_coherencia.md` (sección 5).

## Datos que faltan (Bloomberg y fuentes)
- [ ] YTM del bono UF de Mallplaza (YAS): reemplaza el Kd por el cupón de 3,6% y la tasa efectiva de 4,33% (inconsistencia 2; con 4,33% el precio objetivo baja a ~CLP 2.042).
- [ ] Fechas y monto del dividendo 2S26 (DVD, ~CLP 27 por acción) (inconsistencia 5).
- [ ] Verificar el dato de Gran Plaza (cap rate 9,4% en la Figura 26; −CLP 10,7 por acción) (inconsistencia 6).
- [ ] Verificar una por una las URL de las fuentes políticas de RR2 en Riesgos (inconsistencia 10).
- [ ] Valores exactos de los tramos de inflación y moneda de la Figura 15 de FA, hoy derivados de los totales del borrador (inconsistencia 8).
- [ ] Precios de Bloomberg de los hitos de la Figura R1 (trayectoria): solo el mínimo (10-oct-25), el máximo (27-abr-26) y el cierre (30-sep-26) están confirmados; 9-oct-25, 6-nov-25 y 25-may-26 son precios implícitos en notas de prensa (objetivo / (1 + potencial)); el 15-dic-25 no tiene precio.
- [ ] GLA total de la industria en Chile (≈4,3 millones de m², estimado como GLA del trío sin oficinas ÷ 76%): confirmar con la CChC u otra fuente; de él depende la presión de oferta (5,1%).
- [ ] Renta variable: 5,1% del ingreso (1S26) frente a 5,7% (UDM); el texto rotula cada cifra con su período; unificar cuando el equipo defina cuál usar.

## Valuación
- [ ] Modelo v4 propio: el equipo no entregó su v4; `valuacion/graficos/modelo_v4.py` construye la v4 (v3 + pipeline con fecha desde Exhibits v7). Reemplazar por el modelo del equipo si difiere y correr `generar_todo.sh`.
- [ ] Ponderación de los escenarios 25/50/25 (supuesto del equipo, por validar).
- [ ] Comparar los dos cap rates sobre la misma base (6,0% del mercado y 5,55% real de los tasadores) con una conversión explícita.
- [ ] TIR del accionista por escenario (bear y bull): exige rehacer la trayectoria de deuda de cada caso (hoy solo TIR de la firma).
- [ ] Llevar el DCF al 30-sep-26 (~CLP 2.119) o dejar explícito que no se aplica (inconsistencia 4; el texto lo dice).
- [ ] Rebaja del impuesto corporativo a 23%: no está en el caso base (+5,2% al precio objetivo; sí está en el bull) (inconsistencia 11).
- [ ] Escudo tributario (goodwill tributario y pérdidas): el modelo lo agota en 2029 (impuesto en caja 16,8% → 27%); verificar contra la Nota 13 de los EEFF.
- [ ] Pasar al Excel (hoja TIR_Implicita o TIR_Accionista) las TIR del activo al precio objetivo y al valor DCF, hoy calculadas solo con `tir_valuacion.py`.
- [ ] TIR del proyecto nuevo (4,84% real, 7,98% nominal): validar los supuestos del flujo de un m² (impuesto de largo plazo de 27%, mantención de 5,5% y capital de trabajo de 15,6%).
- [ ] Modelo de sensibilidad a tasas largas e inflación (`Mallplaza_Sensibilidad_Tasas_Inflacion.xlsx`): validar los supuestos de traspaso (θ y λ).
- [ ] Hipótesis de por qué la administración se expande (contabilidad, revalorización, controlador, compensación): verificar con la memoria, las notas de partes relacionadas y la política de bonos.

## ESG
- [ ] Ahorros ESG (Figura E6): los precios unitarios (energía CLP 110/kWh, agua CLP 1.500/m³, residuos CLP 10.000/t) y el 25% del ahorro que retiene Mallplaza son supuestos ilustrativos de la hoja ESG_Ahorros; reemplazarlos con datos reales.
- [ ] Falabella: la matriz de riesgo del borrador decía 53,05% y la tabla y el texto 53,10%; se usó 53,10%. Confirmar con la Memoria.
- [ ] Track record: se usó la versión neutral de la beta 4.1; decidir si se incorporan las «referencias» más elogiosas del equipo (chocan con las sanciones de 2025).
- [ ] Idioma: la beta 4.1 trae títulos y etiquetas en inglés dentro del texto en español; se dejó tal cual.

## Gráficos y diseño
- [ ] Word consolidado: sin la portada de Mallplaza ni el ajuste de gráficos del PDF de Typst, y sin las tablas del Resumen; el PDF es el entregable.
- [ ] Los números de figura citados en los textos con letra («figura 12») están en minúscula en ESG; unificar el estilo de citas.

## Mantenimiento
- [ ] Conservar solo dos versiones: al sacar la próxima versión de Valuación, borrar `secciones/valuacion/documento/anterior/` y mover la actual allí.

## Cerrados
- [x] 17-oct-26: portada con colores de Mallplaza, márgenes casi nulos, gráficos al costado en el PDF.
- [x] 17-oct-26: TIR (firma y accionista, nominal y real) en Valuación y modelo v3.
- [x] 17-oct-26: informe final y definitivo ES/EN (A4, 20 páginas: 10 de cuerpo + 10 de anexos): modelo v4 (pipeline con fecha de 170.200 m²), GLA real de la industria (≈4,3 millones de m²), DCF y figuras dependientes regenerados, relación precio → ventas, nominal/real rotulado, Abstract sin «rinde un 5,0%», escenarios bear/base/bull con el oso y el toro (1.542 / 2.077 / 2.915), Resumen de inversión con tres figuras nuevas (R1–R3), tabla de escenarios y de sensibilidades, Figura 2 horizontal, Figuras 18 y 19 a ancho completo, Anexos 3, 7, 9–19 sincronizados y todas las figuras y anexos citados.
- [x] 17-oct-26: consenso de Bloomberg (14 fuentes, CLP 4.111) incorporado; inconsistencias 9, 12, 13, 14, 15 y 16 resueltas; textos de Industria y Riesgos con cifras vigentes (se cerró el +3,9% por impuesto, ahora +5,2%).

