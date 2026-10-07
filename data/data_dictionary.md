# Data dictionary: peer comparables database for Mallplaza

Base for the relative valuation by comparable multiples (CFA Institute Research Challenge).
Every figure must be traceable to an official document: URL, document date, fiscal period,
page or section, and currency.

## Cutoff date

- Default: latest financial information published up to **30-Sep-2026**, unless the team sets the official CFA Challenge cutoff.
- Documents published after the cutoff go to `data/post_cutoff/` and are never mixed into the main base.

## Folders

| Folder | Content |
|---|---|
| `raw/<Company>/` | Official documents downloaded as published (PDF/XLSX), unmodified |
| `processed/` | The three CSVs below |
| `post_cutoff/` | Documents dated after the cutoff (reference only) |
| `logs/` | Access, download and extraction log |
| `sources/` | Source lists and domains to allow |

## `processed/sources_master.csv` (one row per document)

| Column | Allowed values / format |
|---|---|
| company, country, ticker | Text |
| entity_scope | `consolidado`, `segmento malls`, `subsidiaria inmobiliaria`, `activo/portfolio`, `holding` |
| document_type | e.g. audited FS, interim FS, earnings release, annual report, investor presentation, regulatory filing |
| document_title | Title as published |
| period_covered | e.g. `2Q26`, `FY2025`, `9M26` |
| publication_date, filing_date | `YYYY-MM-DD` |
| language | `es`, `pt`, `en` |
| source_priority | 1–8 per the source hierarchy (1 = official IR site) |
| official_source | `yes` / `no` |
| url | Exact document URL |
| local_file_path | Path under `data/raw/` |
| cutoff_status | `usable_pre_cutoff`, `post_cutoff`, `date_uncertain` |
| audited_status | `audited`, `interim_unaudited`, `earnings_release`, `investor_presentation`, `regulatory_filing` |
| consolidated_or_segment | `consolidated`, `segment`, `subsidiary`, `property_level` |
| notes | Free text |

## `processed/financial_data_raw.csv` (one row per company × period × metric)

| Column | Format |
|---|---|
| fiscal_year, fiscal_quarter | `2026`, `Q2` (`FY` for annual; `LTM` only if built, with formula in `notes`) |
| period_end_date | `YYYY-MM-DD` |
| currency | Reporting currency as published (CLP, BRL, ARS, MXN, PEN, USD). No conversion in this file |
| units | `units`, `thousands`, `millions` |
| metric | Name from the metric list below |
| metric_value | Number exactly as reported (no rounding) |
| accounting_standard | `IFRS`, `IFRS (IAS 29 hyperinflation)`, `MX FIBRA IFRS`, etc. |
| reported_or_adjusted | `reported` / `adjusted` (adjustment described in `notes` with page) |
| recurring_or_nonrecurring | `recurring` / `nonrecurring` |
| source_document, source_url, source_page_or_section | Document title, URL, page and table |
| extraction_method | `manual_pdf_table`, `xlsx_cell`, `calculated` (formula in `notes`) |

**Minimum metrics:** Revenue, Rental Revenue, Net Rental Income, Service Charge Revenue, Parking Revenue, Other Property Income, EBITDA (reported), Adjusted EBITDA, EBIT, Net Income, FFO, AFFO, NOI · Cash, Short-term debt, Long-term debt, Total debt, Net debt, Lease liabilities (IFRS 16, separate), Investment property, Total assets, Non-controlling interests, Equity, NAV (if reported) · CFO, Maintenance capex, Expansion capex, Total capex, Acquisitions/disposals of investment property, FCF (if reported) · Fair value gains/losses on investment property, Gains/losses on asset sales, M&A and integration effects, one-offs, FX gains/losses, Hyperinflation effects, Impairments.

## `processed/operating_kpis_raw.csv` (one row per company × period × KPI)

| Column | Format |
|---|---|
| kpi | Name from the KPI list below |
| kpi_value, unit | Value and unit (`m2`, `%`, `count`, `CLP/m2`, …) |
| scope | `total`, `owned`, `managed`, `country:<X>`, `segment:<X>` |

**KPIs:** Number of malls, GLA total, GLA owned vs managed, Occupancy, Same-store sales, Tenant sales, Visitor traffic, NOI or EBITDA per m², Rent per m², Occupancy cost, Revenue mix, EBITDA margin, NOI margin, Pipeline / land bank, Malls under construction or expansion, Capex expansion vs maintenance, Revenue/EBITDA by country and segment, Net debt/EBITDA, LTV, Weighted average lease term.

## Rules

1. Do not mix periods, currencies or metrics (EBITDA ≠ NOI; FFO ≠ free cash flow; net income ≠ cash).
2. Never use the consolidated EBITDA of diversified holdings (InRetail Perú, IRSA) as a pure mall multiple.
3. A missing metric is recorded as `not disclosed`, `not available` or `no segment-level data found`, listing the official sources checked. Never estimated.
4. Every calculated figure carries its formula in `notes`.
