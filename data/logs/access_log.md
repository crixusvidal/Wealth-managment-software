# Access log

| Date | Tool | Host | Result |
|---|---|---|---|
| 2026-10-07 | curl (shell) | www.parquearauco.com, ri.multiplan.com.br, www.cmfchile.cl, www.sec.gov, www.mallplaza.com, ri.iguatemi.com.br, example.com, www.google.com | `CONNECT tunnel failed, response 403` (proxy blocks all egress) |
| 2026-10-07 | WebFetch | www.sec.gov | `EGRESS_BLOCKED` |
| 2026-10-07 | WebSearch | — | Works, but returns mostly secondary sources; used only to identify official document domains (see `sources/domain_allowlist.md`) |
