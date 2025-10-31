# Résultats des tests manuels

- `browser_container.run_playwright_script` (import XLSX simple) → `2 lignes affichées sur 2`, 2 lignes rendues.
- `browser_container.run_playwright_script` (import XLSX volumineux) → `400 lignes affichées sur 400`.
- `browser_container.run_playwright_script` (export XLSX) → archive ZIP contenant `[Content_Types].xml`, `_rels/.rels`, `docProps/app.xml`, `docProps/core.xml`, `xl/_rels/workbook.xml.rels`, `xl/workbook.xml`, `xl/worksheets/sheet1.xml` (taille ≈ 4 Ko).
- `browser_container.run_playwright_script` (import CSV via __dataExplorer API) → `2 lignes affichées sur 2`.
- `browser_container.run_playwright_script` (import XLSX via __dataExplorer API) → `2 lignes affichées sur 2`.
- `browser_container.run_playwright_script` (import CSV sans worker) → `1 ligne affichée sur 1`.
