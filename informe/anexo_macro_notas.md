# Anexo macro: notas para el equipo (no va en el informe)

## Datos por verificar con Bloomberg antes de enviar
Todos los datos de los gráficos están en `data/macro/macro_2022_2026.xlsx`, con fuente y estado por cifra:

- **Verificado:** cifra citada por EIA, FRED, Bundesbank, FMI o BCCh (búsqueda de oct-2026).
- **Aproximado:** serie pública según el equipo. Cubre la mayor parte de 2022–2025 y Chile 10Y. Hay que reemplazarla con Bloomberg.
- **Interpolado:** meses sin dato publicado (por ejemplo, Brent ene-26 y jul-26; algunos meses de 2026 de Bund, OAT y Chile).

Después de reemplazar los datos, regenerar todo con estos comandos:
```
python data/macro/crear_macro_xlsx.py   # solo si se edita el script; si no, editar el Excel directamente
python graficos/anexo_macro.py
NODE_PATH=$(npm root -g) node informe/anexo_macro.js
```

## Cifras del texto con fuente
| Cifra | Fuente |
|---|---|
| Brent USD 103 (mar-26), 117 (abr), 107 (may), 85 (jun), 91 (ago), 114 (sep) | EIA STEO abr/sep/oct-26; Statista |
| Ormuz: más de un cuarto del petróleo transportado por mar | EIA, Today in Energy |
| UST 10Y 5,22% el 8-oct-26; máximo desde 2002 (no 2007) | FRED DGS10; CNBC 5 y 7-oct-26 |
| OAT–Bund 117 pb al 30-sep-26; ~150 pb a inicios de octubre, máximo desde 2011–12 | Banque de France / Bundesbank (vía ideal-investisseur); Trading Economics |
| FMI: PIB mundial 2026 3,0%; inflación 4,1% (2025) → 4,7% (2026) | FMI, WEO, actualización jul-2026 |
| PIB Chile 2026 0,25–0,75% (desde 2–3% en IPoM dic-25); IPC 3,9%; TPM 4,5% | BCCh, IPoM dic-25, jun-26 y sep-26 |

## Corrección aplicada también en Investment Risks
"máximo desde 2007" se cambió a "máximo en 24 años (desde 2002)", que es lo que reporta CNBC (5-oct-26).
