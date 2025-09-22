## What this script does
Merges **manual keywords** into an existing Excel bibliometrics workbook, normalizes entries (UPPER or Title), recomputes keyword frequencies, flags DOIs still without keywords, appends diagnostics — and preserves the other sheets.

## Expected input (Excel, .xlsx)
Required sheets:
- `keywords_por_artigo` — columns: doi, pmid, revista, publisher_domain, keyword, fonte_keyword
- `artigos_sem_keywords` — columns: doi, pmid, revista, publisher_domain, **KW - busca manual** (free text)
- `revistas_por_artigo` — columns: doi, pmid, revista, publisher_domain
Optional:
- `diagnostico` — columns: metric, value

## Installation
Python 3.10+  
```bash
pip install -r requirements.txt
# python_bibliometrics
