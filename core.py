"""
Lógica de dados do comparador de vacinação (dengue – Taubaté/SP).

Tudo aqui é pandas "puro" (sem Streamlit), para poder ser testado separadamente.
Princípios:
  * nunca inventar dados nem presumir que uma coluna existe: tudo é detectado a partir do conteúdo;
  * nunca descartar registros em silêncio: tudo o que é ignorado/corrigido é contado e devolvido num "log";
  * valor ausente (vazio) é diferente de zero.
"""
from __future__ import annotations

import difflib
import io
import re
import unicodedata
from dataclasses import dataclass, field

import numpy as np
import pandas as pd

# --------------------------------------------------------------------------- #
# Utilitários de texto e números
# --------------------------------------------------------------------------- #
MISSING_TOKENS = {"", "-", "–", "—", "nan", "none", "null", "n/a", "na", "nd", "n/d", "s/i", "#n/a", "#n/d"}


def strip_accents(s: str) -> str:
    return "".join(c for c in unicodedata.normalize("NFKD", s) if not unicodedata.combining(c))


def norm_text(value) -> str:
    """Texto normalizado para COMPARAR (minúsculas, sem acento, sem pontuação, espaços colapsados)."""
    if value is None or (isinstance(value, float) and np.isnan(value)):
        return ""
    s = strip_accents(str(value)).lower()
    s = re.sub(r"[^a-z0-9<>]+", " ", s)
    return re.sub(r"\s+", " ", s).strip()


def clean_label(value):
    """Texto para EXIBIR: tira espaços extras (inclusive espaço não separável). Vazio vira NaN."""
    if value is None or (isinstance(value, float) and np.isnan(value)):
        return np.nan
    if isinstance(value, (pd.Timestamp,)) or pd.isna(value):
        return value if not pd.isna(value) else np.nan
    s = re.sub(r"\s+", " ", str(value).replace("\xa0", " ")).strip()
    return np.nan if s.lower() in MISSING_TOKENS else s


def label_str(value) -> str:
    """Converte um valor de categoria em texto (1.0 -> '1')."""
    if isinstance(value, (float, np.floating)) and float(value).is_integer():
        return str(int(value))
    return str(value)


def parse_number(value) -> float:
    """Converte número em formato brasileiro ou internacional (ou já numérico) para float. Não convertível -> NaN."""
    if value is None:
        return np.nan
    if isinstance(value, (bool, np.bool_)):
        return np.nan
    if isinstance(value, (int, float, np.integer, np.floating)):
        return float(value)
    s = str(value).replace("\xa0", " ").strip()
    if s.lower() in MISSING_TOKENS:
        return np.nan
    s = re.sub(r"(?i)r\$|%|\s", "", s)
    if not re.fullmatch(r"[-+]?[\d.,]+", s):
        return np.nan
    has_dot, has_comma = "." in s, "," in s
    if has_dot and has_comma:
        dec = "," if s.rfind(",") > s.rfind(".") else "."
        thou = "." if dec == "," else ","
        s = s.replace(thou, "").replace(dec, ".")
    elif has_comma:
        s = s.replace(",", ".") if s.count(",") == 1 else s.replace(",", "")
    elif has_dot:
        if re.fullmatch(r"[-+]?\d{1,3}(\.\d{3})+", s):  # 1.234 ou 1.234.567 -> milhar
            s = s.replace(".", "")
        elif s.count(".") > 1:
            return np.nan
    try:
        return float(s)
    except ValueError:
        return np.nan


# Faixa etária: "< 30 dias", "1 ano", "2 a 4 anos", "100 anos ou mais", "12 a 17 anos"...
AGE_RE = re.compile(
    r"^\s*(<=|>=|<|>|menos de|mais de)?\s*(\d+)\s*(?:(?:a|à|-|e)\s*(\d+)\s*)?"
    r"(dias?|semanas?|meses|mes|mês|anos?)\b(?:\s*ou\s*(mais|menos))?\s*$",
    re.IGNORECASE,
)
_UNIT_DAYS = {"dia": 1, "semana": 7, "mes": 30, "mês": 30, "meses": 30, "ano": 365}


def age_sort_key(label: str):
    """Chave numérica (em dias) para ordenar faixas etárias; None se o texto não parece uma faixa etária."""
    m = AGE_RE.match(str(label))
    if not m:
        return None
    prefix, n1, _n2, unit, suffix = m.groups()
    unit = unit.lower()
    unit = unit[:-1] if unit.endswith("s") and unit not in ("meses", "mes") else unit
    key = int(n1) * _UNIT_DAYS.get(unit, 365)
    if prefix and prefix.strip() in ("<", "<=", "menos de"):
        key -= 0.5  # "< 1 ano" vem antes de "1 ano"
    if suffix and suffix.lower() == "mais":
        key += 0.25
    return float(key)


# --------------------------------------------------------------------------- #
# Leitura
# --------------------------------------------------------------------------- #
def list_sheets(data: bytes) -> list[str]:
    return pd.ExcelFile(io.BytesIO(data)).sheet_names


def read_raw(data: bytes, sheet: str) -> pd.DataFrame:
    """Lê a planilha crua (sem cabeçalho, sem conversão de tipos). O índice vira o nº da linha no Excel."""
    raw = pd.read_excel(io.BytesIO(data), sheet_name=sheet, header=None, dtype=object)
    raw.index = range(1, len(raw) + 1)
    return raw


def _is_number_like(v) -> bool:
    return not pd.isna(parse_number(v))


def guess_header_row(raw: pd.DataFrame, max_scan: int = 20) -> int:
    """Posição (0 = primeira linha da planilha) da linha de cabeçalho mais provável."""
    if raw.empty:
        return 0
    width = int(raw.head(max_scan).notna().sum(axis=1).max())
    for pos in range(min(max_scan, len(raw))):
        row = raw.iloc[pos]
        cells = row.dropna()
        if len(cells) == 0 or len(cells) < 0.6 * width:
            continue
        if all(isinstance(c, str) and not _is_number_like(c) for c in cells):
            return pos
    return 0


@dataclass
class TableInfo:
    header_row_excel: int = 1
    original_headers: dict = field(default_factory=dict)  # nome final -> cabeçalho original
    empty_rows_dropped: int = 0
    empty_cols_dropped: int = 0
    notes: list = field(default_factory=list)


def build_table(raw: pd.DataFrame, header_pos: int) -> tuple[pd.DataFrame, TableInfo]:
    """Monta o DataFrame a partir da linha de cabeçalho escolhida. Só remove linhas/colunas 100% vazias (e informa)."""
    info = TableInfo(header_row_excel=int(raw.index[header_pos]) if len(raw) else 1)
    header_vals = list(raw.iloc[header_pos]) if len(raw) else []
    body = raw.iloc[header_pos + 1:].copy()

    all_empty_rows = body.isna().all(axis=1) | body.apply(lambda r: all(pd.isna(clean_label(x)) for x in r), axis=1)
    info.empty_rows_dropped = int(all_empty_rows.sum())
    body = body[~all_empty_rows]

    names, seen = [], {}
    for i, h in enumerate(header_vals):
        h = clean_label(h)
        original = h
        if isinstance(h, str) and h.startswith("="):  # expressão do BI/Qlik no lugar do nome da coluna
            h = None
        base = h if isinstance(h, str) and h else f"Coluna {i + 1}"
        seen[base] = seen.get(base, 0) + 1
        final = base if seen[base] == 1 else f"{base} ({seen[base]})"
        names.append((final, original, h is None and isinstance(original, str)))
    body.columns = [n[0] for n in names]

    # colunas totalmente vazias
    empty_cols = [c for c in body.columns if body[c].map(clean_label).isna().all()]
    info.empty_cols_dropped = len(empty_cols)
    body = body.drop(columns=empty_cols)

    # colunas cujo cabeçalho era uma fórmula do sistema de origem: nomeia pelo conteúdo
    formula_cols = {final for final, _o, was_formula in names if was_formula and final in body.columns}
    for final, original, _wf in names:
        if final in body.columns:
            info.original_headers[final] = original if isinstance(original, str) else ""
    rename = {}
    for c in formula_cols:
        if classify_column(body[c])["kind"] == "faixa_etaria" and "Faixa etária" not in body.columns and "Faixa etária" not in rename.values():
            rename[c] = "Faixa etária"
            info.notes.append(f"O cabeçalho de '{c}' era uma fórmula do sistema de origem; a coluna foi identificada "
                              f"pelo conteúdo (faixas de idade) e nomeada 'Faixa etária'.")
        else:
            info.notes.append(f"O cabeçalho de '{c}' era uma fórmula do sistema de origem e não pôde ser identificado "
                              f"pelo conteúdo; mantive um nome genérico.")
    if rename:
        body = body.rename(columns=rename)
        info.original_headers = {rename.get(k, k): v for k, v in info.original_headers.items()}
    return body, info


# --------------------------------------------------------------------------- #
# Classificação de colunas
# --------------------------------------------------------------------------- #
KIND_LABELS = {
    "faixa_etaria": "Faixa etária",
    "numerico": "Número",
    "data": "Data",
    "texto": "Texto / categoria",
    "vazia": "Vazia",
}


def classify_column(s: pd.Series) -> dict:
    """Detecta o tipo da coluna pelo conteúdo (e não pelo nome)."""
    cleaned = s.map(clean_label)
    nn = cleaned.dropna()
    out = {"kind": "vazia", "n": 0, "text_numbers": 0, "unparsed": []}
    if nn.empty:
        return out
    out["n"] = len(nn)
    uniq = pd.Series(nn.unique())

    # 1) faixa etária
    is_age = uniq.map(lambda v: isinstance(v, str) and bool(AGE_RE.match(v)))
    if is_age.mean() >= 0.7:
        out["kind"] = "faixa_etaria"
        return out

    # 2) data (datas nativas do Excel ou texto dd/mm/aaaa)
    native_dt = nn.map(lambda v: isinstance(v, (pd.Timestamp, np.datetime64)) or hasattr(v, "year") and hasattr(v, "month"))
    if native_dt.mean() >= 0.9:
        out["kind"] = "data"
        return out

    # 3) número (inclusive número guardado como texto)
    parsed = uniq.map(parse_number)
    if parsed.notna().mean() >= 0.9:
        out["kind"] = "numerico"
        out["text_numbers"] = int(nn.map(lambda v: isinstance(v, str)).sum())
        out["unparsed"] = [str(v) for v in uniq[parsed.isna()].tolist()][:20]
        return out

    # 4) data escrita como texto
    str_vals = uniq[uniq.map(lambda v: isinstance(v, str))]
    if len(str_vals) and str_vals.str.contains(r"\d{1,4}[/\-.]\d{1,2}[/\-.]\d{1,4}").mean() >= 0.9:
        dt = pd.to_datetime(str_vals, dayfirst=True, errors="coerce", format="mixed")
        if dt.notna().mean() >= 0.9:
            out["kind"] = "data"
            return out

    out["kind"] = "texto"
    return out


# --------------------------------------------------------------------------- #
# Mapeamento automático de colunas entre os dois arquivos
# --------------------------------------------------------------------------- #
ROLE_HINTS = {
    "Faixa etária": ["faixa etaria", "faixa etária", "idade"],
    "Sexo": ["sexo", "genero"],
    "Bairro": ["bairro", "localidade", "distrito"],
    "Unidade de saúde": ["unidade", "estabelecimento", "ubs", "cnes", "local de aplicacao", "local aplicacao"],
    "Dose": ["dose"],
    "Período": ["data", "periodo", "mes", "ano", "competencia", "semana"],
}
MEASURE_HINTS = ["total", "quantidade", "qtd", "qtde", "n de doses", "numero de doses", "doses aplicadas", "doses"]


def role_of(name: str) -> str | None:
    n = norm_text(name)
    for role, words in ROLE_HINTS.items():
        if any(norm_text(w) in n.split() or norm_text(w) in n for w in words):
            return role
    return None


def name_similarity(a: str, b: str) -> float:
    na, nb = norm_text(a), norm_text(b)
    if not na or not nb:
        return 0.0
    if na == nb:
        return 1.0
    seq = difflib.SequenceMatcher(None, na, nb).ratio()
    ta, tb = set(na.split()), set(nb.split())
    tok = len(ta & tb) / max(1, len(ta | tb))
    contain = 0.85 if (na in nb or nb in na) else 0.0
    return max(seq, tok, contain)


def kinds_compatible(ka: str, kb: str) -> float:
    if ka == kb:
        return 1.0
    if {ka, kb} <= {"texto", "faixa_etaria"}:
        return 0.7
    return 0.0


def suggest_mapping(kinds_a: dict, kinds_b: dict, threshold: float = 0.55) -> dict:
    """Devolve {coluna_A: coluna_B ou None}. Um-para-um, pelos melhores pares (nome + tipo de conteúdo)."""
    cands = []
    for a, ka in kinds_a.items():
        for b, kb in kinds_b.items():
            compat = kinds_compatible(ka, kb)
            if compat == 0:
                continue
            sim = name_similarity(a, b)
            if role_of(a) and role_of(a) == role_of(b):
                sim = max(sim, 0.75)
            score = 0.7 * sim + 0.3 * compat
            if score >= threshold:
                cands.append((score, a, b))
    mapping, used_b = {a: None for a in kinds_a}, set()
    for score, a, b in sorted(cands, reverse=True):
        if mapping[a] is None and b not in used_b:
            mapping[a] = b
            used_b.add(b)
    return mapping


def pair_kind(ka: str, kb: str) -> tuple[str, str | None]:
    """Tipo final de um par de colunas e, se houver incompatibilidade, um aviso."""
    if ka == kb:
        return ka, None
    if {ka, kb} <= {"texto", "faixa_etaria"}:
        return "faixa_etaria", None
    return "texto", (f"Os tipos de conteúdo são diferentes ({KIND_LABELS.get(ka, ka)} × {KIND_LABELS.get(kb, kb)}); "
                     "a coluna será comparada como texto.")


# --------------------------------------------------------------------------- #
# Padronização (limpeza) dos dois arquivos com nomes comuns de coluna
# --------------------------------------------------------------------------- #
@dataclass
class Standardized:
    df: pd.DataFrame                 # colunas com nomes comuns; índice = nº da linha no Excel
    kinds: dict                      # nome comum -> tipo final
    log: dict                        # estatísticas de limpeza por coluna
    label_variants: dict             # coluna -> {rótulo_unificado: [variações]}


def standardize(df: pd.DataFrame, pairs: list[tuple[str, str]], side: str, final_kinds: dict) -> Standardized:
    """
    pairs: [(coluna_A, coluna_B)]; side = 'a' ou 'b'. O nome comum é o nome da coluna em A.
    Limpa textos, converte números/datas e cria '<coluna> (mês)' para colunas de data.
    """
    idx = 0 if side == "a" else 1
    out, kinds, log, variants = {}, {}, {}, {}
    for pair in pairs:
        common, src = pair[0], pair[idx]
        kind = final_kinds[common]
        s = df[src]
        entry = {"coluna_original": src, "tipo": kind}
        if kind == "numerico":
            parsed = s.map(parse_number)
            was_text = s.map(lambda v: isinstance(v, str) and not pd.isna(parse_number(v)))
            blank = s.map(clean_label).isna()
            entry.update(
                texto_convertido_em_numero=int(was_text.sum()),
                vazios=int(blank.sum()),
                nao_convertidos=int((parsed.isna() & ~blank).sum()),
                zeros=int((parsed == 0).sum()),
            )
            out[common] = parsed.astype("float64")
        elif kind == "data":
            cleaned = s.map(clean_label)
            dt = pd.to_datetime(cleaned, dayfirst=True, errors="coerce", format="mixed")
            entry.update(vazios=int(cleaned.isna().sum()), nao_convertidos=int((dt.isna() & cleaned.notna()).sum()))
            out[common] = dt
            out[f"{common} (mês)"] = dt.dt.strftime("%Y-%m")
            kinds[f"{common} (mês)"] = "texto"
            log[f"{common} (mês)"] = {"coluna_original": src, "tipo": "derivada de data"}
        else:
            raw_txt = s.map(lambda v: v if (isinstance(v, str) or pd.isna(v)) else label_str(v))
            cleaned = raw_txt.map(clean_label)
            changed = int(((raw_txt.notna()) & (raw_txt.astype(str) != cleaned.astype(str)) & cleaned.notna()).sum())
            entry.update(vazios=int(cleaned.isna().sum()), textos_com_espacos_corrigidos=changed)
            out[common] = cleaned
            # variações de escrita que serão tratadas como o mesmo rótulo ("Centro" / "CENTRO ")
            v = cleaned.dropna()
            if not v.empty:
                groups = v.groupby(v.map(norm_text)).agg(lambda x: sorted(set(x)))
                multi = {g[0]: g for g in groups if len(g) > 1}
                if multi:
                    variants[common] = multi
        kinds[common] = kind
        log[common] = entry
    std = pd.DataFrame(out, index=df.index)
    return Standardized(std, kinds, log, variants)


# --------------------------------------------------------------------------- #
# Linhas de "total", duplicatas, qualidade
# --------------------------------------------------------------------------- #
TOTAL_WORDS = {"total", "total geral", "geral", "soma", "subtotal", "totais"}


def find_total_rows(df: pd.DataFrame, text_cols: list[str]) -> pd.Series:
    """True nas linhas em que alguma coluna de texto é 'Total', 'Total geral', 'Soma'..."""
    mask = pd.Series(False, index=df.index)
    for c in text_cols:
        if c in df.columns and not pd.api.types.is_numeric_dtype(df[c]) and not pd.api.types.is_datetime64_any_dtype(df[c]):
            mask |= df[c].map(lambda v: norm_text(v) in TOTAL_WORDS)
    return mask


def exact_duplicates(df: pd.DataFrame) -> pd.DataFrame:
    """Linhas totalmente idênticas (todas as colunas). Mostra todas as ocorrências (inclusive a primeira)."""
    if df.empty:
        return df
    return df[df.duplicated(keep=False)]


def _key_series(df: pd.DataFrame, key_cols: list[str]) -> pd.Series:
    parts = []
    for c in key_cols:
        col = df[c]
        if pd.api.types.is_numeric_dtype(col):
            parts.append(col.map(lambda v: "" if pd.isna(v) else label_str(v)))
        else:
            parts.append(col.map(norm_text))
    key = parts[0]
    for p in parts[1:]:
        key = key + "␟" + p
    return key


def duplicated_keys(df: pd.DataFrame, key_cols: list[str]) -> pd.DataFrame:
    """Registros cuja chave (combinação das colunas escolhidas, ignorando maiúsculas/acentos) aparece mais de uma vez."""
    if df.empty or not key_cols:
        return df.iloc[0:0]
    k = _key_series(df, key_cols)
    return df[k.duplicated(keep=False)].assign(**{"_chave": k[k.duplicated(keep=False)]}).sort_values("_chave").drop(columns="_chave")


def key_quality(df: pd.DataFrame, key_cols: list[str]) -> dict:
    if not key_cols:
        return {"reliable": False, "reason": "Nenhuma coluna de chave escolhida.", "n": len(df), "missing": 0, "dups": 0}
    k = _key_series(df, key_cols)
    missing = int(df[key_cols].isna().any(axis=1).sum())
    dups = int(k.duplicated(keep=False).sum())
    ok = missing == 0 and dups == 0 and len(df) > 0
    reason = ""
    if missing:
        reason = f"{missing} registro(s) com chave vazia. "
    if dups:
        reason += f"{dups} registro(s) com chave repetida."
    return {"reliable": ok, "reason": reason.strip(), "n": len(df), "missing": missing, "dups": dups}


def profile_columns(df: pd.DataFrame, log: dict) -> pd.DataFrame:
    rows = []
    for c in df.columns:
        entry = log.get(c, {})
        s = df[c]
        row = {
            "Coluna": c,
            "Tipo detectado": KIND_LABELS.get(entry.get("tipo"), entry.get("tipo", "")),
            "Preenchidos": int(s.notna().sum()),
            "Vazios (ausentes)": int(s.isna().sum()),
            "Valores diferentes": int(s.nunique(dropna=True)),
        }
        if entry.get("tipo") == "numerico":
            row["Zeros (valor 0 real)"] = int((s == 0).sum())
            row["Números guardados como texto (convertidos)"] = entry.get("texto_convertido_em_numero", 0)
            row["Não convertidos (viraram vazio)"] = entry.get("nao_convertidos", 0)
        elif entry.get("tipo") in ("texto", "faixa_etaria"):
            row["Textos com espaços corrigidos"] = entry.get("textos_com_espacos_corrigidos", 0)
        rows.append(row)
    return pd.DataFrame(rows)


# --------------------------------------------------------------------------- #
# Comparação agregada por categoria
# --------------------------------------------------------------------------- #
def _aggregate(df: pd.DataFrame, dim: str, measure: str | None) -> pd.DataFrame:
    d = df[[dim] + ([measure] if measure else [])].copy()
    d["_k"] = d[dim].map(lambda v: "" if pd.isna(v) else norm_text(label_str(v)))
    d["_lbl"] = d[dim].map(lambda v: np.nan if pd.isna(v) else label_str(v))
    # registros sem categoria ficam em um grupo próprio, visível
    d.loc[d["_k"] == "", "_lbl"] = "(sem informação)"
    d.loc[d["_k"] == "", "_k"] = "(sem informação)"
    g = d.groupby("_k", sort=False)
    res = pd.DataFrame({"label": g["_lbl"].first(), "registros": g.size()})
    if measure:
        res["valor"] = g[measure].sum(min_count=1)   # todos vazios -> NaN (e não 0)
    else:
        res["valor"] = res["registros"].astype(float)
    return res


def _order_labels(keys: list, labels: dict, kind: str) -> list:
    if kind == "faixa_etaria" and all(age_sort_key(labels[k]) is not None for k in keys):
        return sorted(keys, key=lambda k: age_sort_key(labels[k]))
    if keys and all(re.fullmatch(r"\d{4}-\d{2}", labels[k]) for k in keys):
        return sorted(keys, key=lambda k: labels[k])
    return keys


def compare_by_dimension(a: pd.DataFrame, b: pd.DataFrame, dim: str, measure: str | None, kind: str = "texto") -> pd.DataFrame:
    """Tabela categoria a categoria. Ausente ≠ zero: categoria que não existe num arquivo fica em branco."""
    ga, gb = _aggregate(a, dim, measure), _aggregate(b, dim, measure)
    keys = list(ga.index) + [k for k in gb.index if k not in ga.index]
    labels = {k: (ga.loc[k, "label"] if k in ga.index else gb.loc[k, "label"]) for k in keys}
    keys = _order_labels(keys, labels, kind)

    res = pd.DataFrame({"Categoria": [labels[k] for k in keys]})
    res["Registros A"] = [int(ga.loc[k, "registros"]) if k in ga.index else np.nan for k in keys]
    res["Registros B"] = [int(gb.loc[k, "registros"]) if k in gb.index else np.nan for k in keys]
    res[["Registros A", "Registros B"]] = res[["Registros A", "Registros B"]].astype("Int64")
    res["Valor A"] = [ga.loc[k, "valor"] if k in ga.index else np.nan for k in keys]
    res["Valor B"] = [gb.loc[k, "valor"] if k in gb.index else np.nan for k in keys]
    res["Diferença (B − A)"] = res["Valor B"] - res["Valor A"]      # NaN se faltar algum dos lados
    with np.errstate(divide="ignore", invalid="ignore"):
        pct = (res["Valor B"] - res["Valor A"]) / res["Valor A"] * 100
    res["Variação %"] = pct.where(res["Valor A"] != 0)               # base zero -> sem %
    tot_a, tot_b = res["Valor A"].sum(min_count=1), res["Valor B"].sum(min_count=1)
    res["% do total A"] = res["Valor A"] / tot_a * 100 if tot_a and tot_a != 0 else np.nan
    res["% do total B"] = res["Valor B"] / tot_b * 100 if tot_b and tot_b != 0 else np.nan
    res["Mudança de participação (p.p.)"] = res["% do total B"] - res["% do total A"]

    def status(i):
        in_a, in_b = bool(pd.notna(res.at[i, "Registros A"])), bool(pd.notna(res.at[i, "Registros B"]))
        if in_a and in_b:
            if pd.isna(res.at[i, "Valor A"]) or pd.isna(res.at[i, "Valor B"]):
                return "Sem valor em um dos arquivos"
            return "Nos dois arquivos"
        return "Só no arquivo A" if in_a else "Só no arquivo B"

    res["Situação"] = [status(i) for i in res.index]
    return res


def summarize_totals(a: pd.DataFrame, b: pd.DataFrame, measure: str | None, table: pd.DataFrame | None = None) -> dict:
    """Totais gerais, diferença absoluta/percentual e conciliação com as categorias exclusivas."""
    out = {"registros_a": len(a), "registros_b": len(b), "tem_medida": bool(measure)}
    if measure:
        ta, tb = a[measure].sum(min_count=1), b[measure].sum(min_count=1)
    else:
        ta, tb = float(len(a)), float(len(b))
    out["total_a"], out["total_b"] = ta, tb
    out["dif"] = (tb - ta) if not (pd.isna(ta) or pd.isna(tb)) else np.nan
    out["pct"] = (out["dif"] / ta * 100) if (not pd.isna(ta) and ta != 0 and not pd.isna(out["dif"])) else np.nan
    if table is not None and not table.empty:
        comparable = table[table["Situação"] == "Nos dois arquivos"]
        out["dif_comparaveis"] = comparable["Diferença (B − A)"].sum()
        out["so_a"] = table.loc[table["Situação"] == "Só no arquivo A", "Valor A"].sum(min_count=1)
        out["so_b"] = table.loc[table["Situação"] == "Só no arquivo B", "Valor B"].sum(min_count=1)
        out["n_so_a"] = int((table["Situação"] == "Só no arquivo A").sum())
        out["n_so_b"] = int((table["Situação"] == "Só no arquivo B").sum())
        out["n_maior"] = int((comparable["Diferença (B − A)"] > 0).sum())
        out["n_menor"] = int((comparable["Diferença (B − A)"] < 0).sum())
        out["n_igual"] = int((comparable["Diferença (B − A)"] == 0).sum())
    return out


def fmt_num(v, decimals: int = 0, signed: bool = False) -> str:
    """Formata número no padrão brasileiro (1.234,5)."""
    if v is None or pd.isna(v):
        return "—"
    s = f"{v:+,.{decimals}f}" if signed else f"{v:,.{decimals}f}"
    return s.replace(",", "X").replace(".", ",").replace("X", ".")


def plain_summary(table: pd.DataFrame, totals: dict, label_a: str, label_b: str, dim: str, unit: str = "doses") -> list[str]:
    """Frases em português simples descrevendo as principais diferenças."""
    lines = []
    d = totals.get("dif")
    if d is not None and not pd.isna(d):
        if d == 0:
            lines.append(f"Os dois arquivos têm o **mesmo total** de {unit} ({fmt_num(totals['total_a'])}).")
        else:
            mais = "mais" if d > 0 else "menos"
            pct = "" if pd.isna(totals.get("pct")) else f" ({fmt_num(abs(totals['pct']), 1)}%)"
            lines.append(f"**{label_b}** tem **{fmt_num(abs(d))} {unit} a {mais}** que **{label_a}**{pct}: "
                         f"{fmt_num(totals['total_a'])} → {fmt_num(totals['total_b'])}.")
    comp = table[table["Situação"] == "Nos dois arquivos"]
    if not comp.empty:
        lines.append(f"Em **{dim}**: {totals['n_maior']} categoria(s) com valor maior em {label_b}, "
                     f"{totals['n_menor']} com valor menor e {totals['n_igual']} sem mudança.")
        top_up = comp.nlargest(1, "Diferença (B − A)").iloc[0]
        top_dn = comp.nsmallest(1, "Diferença (B − A)").iloc[0]
        if top_up["Diferença (B − A)"] > 0:
            lines.append(f"Maior aumento em números absolutos: **{top_up['Categoria']}** "
                         f"({fmt_num(top_up['Diferença (B − A)'], signed=True)}).")
        if top_dn["Diferença (B − A)"] < 0:
            lines.append(f"Maior redução em números absolutos: **{top_dn['Categoria']}** "
                         f"({fmt_num(top_dn['Diferença (B − A)'], signed=True)}).")
        shift = comp.reindex(comp["Mudança de participação (p.p.)"].abs().sort_values(ascending=False).index).head(1)
        if not shift.empty and not pd.isna(shift.iloc[0]["Mudança de participação (p.p.)"]):
            r = shift.iloc[0]
            lines.append(f"Quem mais mudou de **peso** no total: **{r['Categoria']}** "
                         f"({fmt_num(r['% do total A'], 1)}% → {fmt_num(r['% do total B'], 1)}%).")
    if totals.get("n_so_a"):
        lines.append(f"{totals['n_so_a']} categoria(s) existem **só em {label_a}**.")
    if totals.get("n_so_b"):
        lines.append(f"{totals['n_so_b']} categoria(s) existem **só em {label_b}**.")
    return lines


# --------------------------------------------------------------------------- #
# Comparação registro a registro (exige chave confiável)
# --------------------------------------------------------------------------- #
def compare_records(a: pd.DataFrame, b: pd.DataFrame, key_cols: list[str], compare_cols: list[str]) -> dict:
    """
    Registros adicionados (só em B), removidos (só em A) e alterados (mesma chave, valores diferentes).
    Só deve ser chamada quando key_quality(...)['reliable'] for True nos dois arquivos.
    """
    ka, kb = _key_series(a, key_cols), _key_series(b, key_cols)
    A = a.assign(_k=ka.values, _linha_A=a.index)
    B = b.assign(_k=kb.values, _linha_B=b.index)
    removed = A[~A["_k"].isin(B["_k"])].drop(columns="_k")
    added = B[~B["_k"].isin(A["_k"])].drop(columns="_k")
    both = A.merge(B, on="_k", suffixes=(" [A]", " [B]"))

    changes = []
    for c in compare_cols:
        if c in key_cols:
            continue
        ca = both[f"{c} [A]"] if f"{c} [A]" in both else both[c]
        cb = both[f"{c} [B]"] if f"{c} [B]" in both else both[c]
        if pd.api.types.is_numeric_dtype(ca) and pd.api.types.is_numeric_dtype(cb):
            same = (ca.isna() & cb.isna()) | pd.Series(np.isclose(ca.fillna(0), cb.fillna(0)) & ca.notna() & cb.notna(), index=both.index)
        else:
            na_, nb_ = ca.map(lambda v: norm_text(label_str(v)) if not pd.isna(v) else None), cb.map(lambda v: norm_text(label_str(v)) if not pd.isna(v) else None)
            same = (na_ == nb_) | (na_.isna() & nb_.isna())
        for i in both.index[~same]:
            row = {k: both.at[i, f"{k} [A]"] if f"{k} [A]" in both else both.at[i, k] for k in key_cols}
            row.update({"Coluna alterada": c, "Valor A": ca.at[i], "Valor B": cb.at[i]})
            if pd.api.types.is_numeric_dtype(ca) and not (pd.isna(ca.at[i]) or pd.isna(cb.at[i])):
                row["Diferença (B − A)"] = cb.at[i] - ca.at[i]
            changes.append(row)
    changed = pd.DataFrame(changes)
    n_changed_records = changed[key_cols].drop_duplicates().shape[0] if not changed.empty else 0
    return {
        "added": added, "removed": removed, "changed": changed,
        "n_added": len(added), "n_removed": len(removed),
        "n_changed": n_changed_records, "n_same": len(both) - n_changed_records,
    }


# --------------------------------------------------------------------------- #
# Sugestões automáticas (chave, critério do arquivo)
# --------------------------------------------------------------------------- #
ID_WORDS = {"id", "codigo", "cod", "cns", "cpf", "protocolo", "registro", "identificador", "numero"}


def suggest_key(a: pd.DataFrame, b: pd.DataFrame, columns: list[str], dims: list[str]) -> list[str]:
    """
    Sugere uma chave só quando ela é confiável nos dois arquivos (sem vazios e sem repetições).
    1) uma coluna com nome de identificador; 2) a combinação das dimensões escolhidas; senão, nenhuma.
    """
    for c in columns:
        if c.endswith("(mês)"):
            continue
        if set(norm_text(c).split()) & ID_WORDS:
            if key_quality(a, [c])["reliable"] and key_quality(b, [c])["reliable"]:
                return [c]
    if dims and key_quality(a, dims)["reliable"] and key_quality(b, dims)["reliable"]:
        return list(dims)
    return []


def detect_criterion(filename: str) -> str | None:
    """Pelo NOME do arquivo, tenta reconhecer se é 'Residência' ou 'Ocorrência' (apenas uma dica, nunca uma certeza)."""
    n = norm_text(filename)
    if "residenc" in n:
        return "Residência"
    if "ocorrenc" in n:
        return "Ocorrência"
    return None