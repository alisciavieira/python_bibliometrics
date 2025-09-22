## What this script does
Merges **manual keywords** into an existing Excel bibliometrics workbook, normalizes entries (UPPER or Title), recomputes keyword frequencies, flags DOIs still without keywords, appends diagnostics — and preserves the other sheets.

## Expected input (Excel, .xlsx)
Required sheets:
- `keywords_per_article` — columns: doi, pmid, journal, publisher_domain, keyword, fonte_keyword
- `articles_without_keywords` — columns: doi, pmid, journal, publisher_domain, **KW - manual search** (free text)
- `journals_per_paper` — columns: doi, pmid, journal, publisher_domain
Optional:
- `diagnosis` — columns: metric, value

## Installation
Python 3.10+  
```bash
pip install -r requirements.txt
# python_bibliometrics
