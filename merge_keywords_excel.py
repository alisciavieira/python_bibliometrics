import pandas as pd
import re

# ===== CONFIG =====
IN_XLSX  = "RP11(final antes de unir KWs).xlsx"   # sua planilha já pronta
OUT_XLSX = "RP12[KWs unidas com manuais].xlsx"    # saída
SHEET_KW        = "keywords_por_artigo"
SHEET_KW_FREQ   = "keywords_frequencia"
SHEET_NO_KW     = "artigos_sem_keywords"
SHEET_REV       = "revistas_por_artigo"
SHEET_DIAG      = "diagnostico"
MANUAL_COL      = "KW - busca manual"            # nome exato da coluna que você preencheu
KEYWORD_NORMALIZE = "upper"                      # "upper" (recomendado) ou "title"
# ===================

def normalize_keyword(s: str) -> str:
    s = " ".join(str(s).split())
    return s.title() if KEYWORD_NORMALIZE == "title" else s.upper()

split_pat = re.compile(r"[;,\|/·•]")

# ---- 1) Ler todas as abas ----
sheets = pd.read_excel(IN_XLSX, sheet_name=None)

if SHEET_KW not in sheets or SHEET_NO_KW not in sheets or SHEET_REV not in sheets:
    missing = [n for n in [SHEET_KW, SHEET_NO_KW, SHEET_REV] if n not in sheets]
    raise RuntimeError(f"Aba(s) ausente(s) na planilha: {missing}")

df_kw_old   = sheets[SHEET_KW].copy()
df_no_kw    = sheets[SHEET_NO_KW].copy()
df_rev      = sheets[SHEET_REV].copy()
df_diag_old = sheets.get(SHEET_DIAG, pd.DataFrame(columns=["metric","value"])).copy()

# ---- 2) Expandir keywords manuais (respeita duplicatas do input) ----
df_no_kw[MANUAL_COL] = df_no_kw.get(MANUAL_COL, "").astype(str)

df_manual_rows = []
for _, row in df_no_kw.iterrows():
    raw = (row.get(MANUAL_COL) or "").strip()
    if not raw:
        continue
    parts = [p.strip() for p in split_pat.split(raw) if p.strip()]
    parts = [normalize_keyword(p) for p in parts]
    for kw in parts:
        df_manual_rows.append({
            "doi": row.get("doi"),
            "pmid": row.get("pmid"),
            "revista": row.get("revista"),
            "publisher_domain": row.get("publisher_domain"),
            "keyword": kw,
            "fonte_keyword": "manual"
        })

df_manual = pd.DataFrame(df_manual_rows, columns=[
    "doi","pmid","revista","publisher_domain","keyword","fonte_keyword"
])

# ---- 3) Unir com as keywords existentes e recalcular frequência ----
df_kw_new = pd.concat([df_kw_old, df_manual], ignore_index=True)

if not df_kw_new.empty:
    df_kw_freq_new = (df_kw_new.groupby("keyword", as_index=False)
                                .size()
                                .rename(columns={"size":"frequencia"})
                                .sort_values(by=["frequencia","keyword"],
                                             ascending=[False, True], kind="stable"))
else:
    df_kw_freq_new = pd.DataFrame(columns=["keyword","frequencia"])

# ---- 4) Quais DOIs (não-conferência) ainda ficaram sem keywords após a união? ----
# Base de linhas não-conferência (com duplicatas) = a própria 'revistas_por_artigo'
base_rev_dois_all = df_rev["doi"].dropna().astype(str).tolist()
dois_com_kw_pos = set(df_kw_new["doi"].dropna().astype(str).unique())

rows_rest = []
for _, r in df_rev.iterrows():
    d = str(r.get("doi"))
    if d not in dois_com_kw_pos:
        rows_rest.append({
            "doi": d,
            "pmid": r.get("pmid"),
            "revista": r.get("revista"),
            "publisher_domain": r.get("publisher_domain"),
            "observacao": "continua_sem_keywords_apos_uniao"
        })
df_no_kw_rest = pd.DataFrame(rows_rest)

# ---- 5) Diagnóstico pós-união ----
n_kw_lines_pos   = len(df_kw_new)
n_kw_unique_pos  = df_kw_freq_new.shape[0]

n_no_kw_lines_pos   = len(df_no_kw_rest)
n_no_kw_unique_pos  = df_no_kw_rest["doi"].nunique() if not df_no_kw_rest.empty else 0
n_nonconf_unique    = df_rev["doi"].nunique() if not df_rev.empty else 0
pct_no_kw_unique_pos = round(100.0 * n_no_kw_unique_pos / n_nonconf_unique, 2) if n_nonconf_unique else 0.0

df_diag_add = pd.DataFrame([
    {"metric":"linhas_keywords_por_artigo_pos_merge", "value": n_kw_lines_pos},
    {"metric":"keywords_unicas_pos_merge",            "value": n_kw_unique_pos},
    {"metric":"artigos_sem_keywords_dois_unicos_pos_merge",             "value": n_no_kw_unique_pos},
    {"metric":"artigos_sem_keywords_pct_dos_nao_conferencia_pos_merge", "value": pct_no_kw_unique_pos},
])

df_diag_new = pd.concat([df_diag_old, df_diag_add], ignore_index=True)

# ---- 6) Gravar novo arquivo mantendo as outras abas inalteradas ----
with pd.ExcelWriter(OUT_XLSX, engine="openpyxl") as xlw:
    for name, df in sheets.items():
        # vamos sobrescrever as que atualizamos abaixo
        if name in {SHEET_KW, SHEET_KW_FREQ, SHEET_NO_KW, SHEET_DIAG}:
            continue
        df.to_excel(xlw, index=False, sheet_name=name)

    # sobrescritas/novas
    df_kw_new.to_excel(xlw, index=False, sheet_name=SHEET_KW)
    df_kw_freq_new.to_excel(xlw, index=False, sheet_name=SHEET_KW_FREQ)
    # mantemos a aba original de 'artigos_sem_keywords' como estava
    sheets[SHEET_NO_KW].to_excel(xlw, index=False, sheet_name=SHEET_NO_KW)
    # e criamos a aba com os que AINDA ficaram sem KWs
    df_no_kw_rest.to_excel(xlw, index=False, sheet_name="artigos_sem_keywords_restantes")
    # diagnóstico acrescido
    df_diag_new.to_excel(xlw, index=False, sheet_name=SHEET_DIAG)

print("✅ Unificação concluída!")
print("Arquivo gerado:", OUT_XLSX)
print("Resumo pós-união:")
print(" - linhas_keywords_por_artigo_pos_merge:", n_kw_lines_pos)
print(" - keywords_unicas_pos_merge:", n_kw_unique_pos)
print(" - artigos_sem_keywords_dois_unicos_pos_merge:", n_no_kw_unique_pos)
print(" - artigos_sem_keywords_pct_dos_nao_conferencia_pos_merge:", pct_no_kw_unique_pos, "%")
