"""
Comparador de arquivos de vacinação contra a dengue – Taubaté/SP.

Como rodar:
    pip install -r requirements.txt
    streamlit run app.py
"""
from __future__ import annotations

import io
import re

import numpy as np
import pandas as pd
import plotly.graph_objects as go
import streamlit as st

import core

st.set_page_config(page_title="Comparador de vacinação – dengue (Taubaté)", page_icon="💉", layout="wide")

COLOR_A, COLOR_B = "#3b78d8", "#e8833a"      # arquivo A / arquivo B (distinguíveis por daltônicos)
COLOR_UP, COLOR_DOWN = "#1b9e8f", "#c8553d"  # B maior / B menor (sem verde-vermelho puro)
NONE_OPT = "— sem correspondência —"
NO_MEASURE = "(nenhuma — contar registros)"


# --------------------------------------------------------------------------- #
# Funções de apoio da interface
# --------------------------------------------------------------------------- #
_VER = tuple(int(x) for x in re.findall(r"\d+", st.__version__)[:2])
_NEW_WIDTH_API = _VER >= (1, 50)   # a partir daqui existe width="stretch" (e use_container_width fica obsoleto)


def show_df(df: pd.DataFrame, **kwargs):
    """st.dataframe ocupando a largura toda, compatível com versões novas e antigas do Streamlit."""
    if _NEW_WIDTH_API:
        st.dataframe(df, width="stretch", **kwargs)
    else:
        st.dataframe(df, use_container_width=True, **kwargs)


def show_chart(fig: go.Figure):
    if _NEW_WIDTH_API:
        st.plotly_chart(fig, width="stretch")
    else:
        st.plotly_chart(fig, use_container_width=True)


def pretty(df: pd.DataFrame) -> pd.DataFrame:
    """Para exibição: colunas numéricas com valores inteiros não mostram '.0'."""
    out = df.copy()
    for c in out.columns:
        if pd.api.types.is_float_dtype(out[c]):
            nn = out[c].dropna()
            if len(nn) and (nn % 1 == 0).all():
                out[c] = out[c].astype("Int64")
    return out


@st.cache_data(show_spinner=False)
def cached_sheets(data: bytes):
    return core.list_sheets(data)


@st.cache_data(show_spinner=False)
def cached_raw(data: bytes, sheet: str):
    return core.read_raw(data, sheet)


@st.cache_data(show_spinner=False)
def cached_table(data: bytes, sheet: str, header_pos: int):
    raw = core.read_raw(data, sheet)
    table, info = core.build_table(raw, header_pos)
    kinds = {c: core.classify_column(table[c])["kind"] for c in table.columns}
    return table, info, kinds


def to_excel_bytes(sheets: dict[str, pd.DataFrame]) -> bytes:
    buf, used = io.BytesIO(), set()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        for name, df in sheets.items():
            base = re.sub(r"[\[\]:*?/\\]", "", name)[:28] or "Planilha"
            final, n = base, 2
            while final in used:
                final, n = f"{base[:26]}_{n}", n + 1
            used.add(final)
            df.to_excel(writer, sheet_name=final, index=False)
    return buf.getvalue()


def load_side(tag: str, title: str, uploaded) -> dict:
    """Escolha de planilha + cabeçalho + prévia para um dos arquivos."""
    data = uploaded.getvalue()
    skey = f"{tag}_{uploaded.name}_{len(data)}"
    st.markdown(f"##### {title}: `{uploaded.name}`")
    try:
        sheets = cached_sheets(data)
    except Exception as e:  # arquivo corrompido, protegido por senha etc.
        st.error(f"Não consegui abrir este arquivo ({e}). Verifique se é um Excel válido e sem senha.")
        st.stop()
    if not sheets:
        st.error("Este arquivo não tem nenhuma planilha.")
        st.stop()
    sheet = st.selectbox("Planilha a analisar", sheets, key=f"sheet_{skey}")
    raw = cached_raw(data, sheet)
    if raw.empty:
        st.error("A planilha escolhida está vazia.")
        st.stop()
    guessed = core.guess_header_row(raw)
    with st.expander("Opções avançadas da planilha"):
        header_row = st.number_input(
            "Linha do Excel onde estão os nomes das colunas", min_value=1, max_value=len(raw),
            value=int(raw.index[guessed]), step=1, key=f"hdr_{skey}_{sheet}",
            help="Detectada automaticamente. Mude se o título das colunas estiver em outra linha.")
    header_pos = int(header_row) - 1
    table, info, kinds = cached_table(data, sheet, header_pos)
    if table.empty or table.shape[1] == 0:
        st.error("Não encontrei dados abaixo da linha de cabeçalho escolhida. Ajuste a linha nas opções avançadas.")
        st.stop()

    crit = core.detect_criterion(uploaded.name)
    default_label = crit or re.sub(r"[_\-]+", " ", uploaded.name.rsplit(".", 1)[0])
    label = st.text_input("Nome deste arquivo nos resultados", value=default_label, key=f"lbl_{skey}")
    st.caption(f"{len(table):,} linhas de dados × {table.shape[1]} colunas · cabeçalho na linha {info.header_row_excel} "
               f"do Excel".replace(",", "."))
    for n in info.notes:
        st.info(n)
    if info.empty_rows_dropped or info.empty_cols_dropped:
        st.warning(f"Ignorei {info.empty_rows_dropped} linha(s) e {info.empty_cols_dropped} coluna(s) "
                   f"totalmente vazias (nenhum dado foi perdido).")
    show_df(pretty(table.head(100)))
    if len(table) > 100:
        st.caption("Mostrando as 100 primeiras linhas.")
    return {"name": uploaded.name, "label": label.strip() or title, "table": table, "info": info,
            "kinds": kinds, "sig": f"{skey}_{sheet}_{header_pos}", "criterion": crit}


# --------------------------------------------------------------------------- #
# Cabeçalho
# --------------------------------------------------------------------------- #
st.title("💉 Comparador de arquivos – vacinação contra a dengue em Taubaté")
st.markdown("Envie **dois arquivos Excel** e veja, de forma visual, **o que mudou de um para o outro**.")
with st.expander("Como usar (passo a passo)"):
    st.markdown(
        "1. **Envie os dois arquivos** (.xlsx ou .xls) e escolha a planilha de cada um.\n"
        "2. **Confira as colunas correspondentes** – o app sugere automaticamente; você pode corrigir.\n"
        "3. **Defina o que comparar** (qual coluna tem o número de doses e por quais categorias quer separar).\n"
        "4. Veja os resultados nas abas: resumo, por categoria, registro a registro, qualidade dos dados e exportação.\n\n"
        "Nada é apagado dos seus dados sem aviso: tudo o que o app corrige ou ignora é informado."
    )

# --------------------------------------------------------------------------- #
# 1. Upload
# --------------------------------------------------------------------------- #
st.header("1. Enviar os arquivos")
col_up_a, col_up_b = st.columns(2)
with col_up_a:
    up_a = st.file_uploader("Arquivo A (primeiro)", type=["xlsx", "xls"], key="up_a")
with col_up_b:
    up_b = st.file_uploader("Arquivo B (segundo)", type=["xlsx", "xls"], key="up_b")
if not (up_a and up_b):
    st.info("Envie os dois arquivos para começar.")
    st.stop()

col_a, col_b = st.columns(2)
with col_a:
    side_a = load_side("a", "Arquivo A", up_a)
with col_b:
    side_b = load_side("b", "Arquivo B", up_b)

label_a, label_b = side_a["label"], side_b["label"]
if label_a == label_b:
    label_a, label_b = f"{label_a} (A)", f"{label_b} (B)"

# Diferença não é, por si só, erro
crit_a, crit_b = side_a["criterion"], side_b["criterion"]
with st.container(border=True):
    st.markdown("**Antes de comparar: diferença não é o mesmo que erro.** Dois arquivos podem divergir por motivos legítimos, "
                "como o **critério de contagem** (doses por *residência* do paciente × doses por *local de ocorrência/aplicação*), "
                "a **data da extração**, **filtros diferentes** ou **correções posteriores** no sistema. "
                "Este app mostra **onde** e **quanto** os valores diferem – ele não decide qual arquivo está certo.")
    if crit_a and crit_b and crit_a != crit_b:
        st.markdown(f"Pelo **nome** dos arquivos, parece que **A é por {crit_a.lower()}** e **B é por {crit_b.lower()}** "
                    "(confira se é isso mesmo). Nesse caso, é **esperado** que os números sejam diferentes.")

# --------------------------------------------------------------------------- #
# 2. Mapeamento de colunas
# --------------------------------------------------------------------------- #
st.header("2. Conferir as colunas correspondentes")
st.caption("O app procurou colunas parecidas (pelo nome e pelo conteúdo). Corrija abaixo se alguma estiver errada. "
           "Os arquivos **não precisam** ter a mesma estrutura.")
kinds_a, kinds_b = side_a["kinds"], side_b["kinds"]
auto_map = core.suggest_mapping(kinds_a, kinds_b)
options_b = [NONE_OPT] + list(kinds_b)
sig = f"{side_a['sig']}|{side_b['sig']}"

h1, h2, h3 = st.columns([3, 3, 3])
h1.markdown("**Coluna do arquivo A**")
h2.markdown("**Tipo de conteúdo detectado**")
h3.markdown("**Coluna correspondente no arquivo B**")
chosen = {}
for i, (ca, ka) in enumerate(kinds_a.items()):
    c1, c2, c3 = st.columns([3, 3, 3])
    c1.write(ca)
    c2.write(core.KIND_LABELS.get(ka, ka))
    default = auto_map.get(ca) or NONE_OPT
    chosen[ca] = c3.selectbox(f"Correspondente de {ca}", options_b, index=options_b.index(default),
                              key=f"map_{sig}_{i}", label_visibility="collapsed")

used = [b for b in chosen.values() if b != NONE_OPT]
dups_b = sorted({b for b in used if used.count(b) > 1})
if dups_b:
    st.error("A mesma coluna do arquivo B foi associada a mais de uma coluna do arquivo A: " + ", ".join(dups_b) +
             ". Ajuste para continuar.")
    st.stop()
pairs = [(a, b) for a, b in chosen.items() if b != NONE_OPT]
if not pairs:
    st.error("Nenhuma coluna foi associada entre os arquivos. Associe ao menos uma coluna para comparar.")
    st.stop()
only_a = [a for a, b in chosen.items() if b == NONE_OPT]
only_b = [b for b in kinds_b if b not in used]
if only_a or only_b:
    st.info("**Colunas sem correspondência** (não entram na comparação, mas não foram apagadas) – "
            f"só em A: {', '.join(only_a) or 'nenhuma'}; só em B: {', '.join(only_b) or 'nenhuma'}.")

final_kinds = {}
for a, b in pairs:
    k, warn = core.pair_kind(kinds_a[a], kinds_b[b])
    final_kinds[a] = k
    if warn:
        st.warning(f"**{a} ↔ {b}**: {warn}")

std_a = core.standardize(side_a["table"], pairs, "a", final_kinds)
std_b = core.standardize(side_b["table"], pairs, "b", final_kinds)
df_a_all, df_b_all = std_a.df, std_b.df
kinds_std = std_a.kinds
common_cols = list(df_a_all.columns)

# --------------------------------------------------------------------------- #
# 3. O que comparar
# --------------------------------------------------------------------------- #
st.header("3. Definir o que comparar")
numeric_cols = [c for c in common_cols if kinds_std.get(c) == "numerico"]
hinted = [c for c in numeric_cols if any(core.norm_text(h) in core.norm_text(c) for h in core.MEASURE_HINTS)]
measure_opts = [NO_MEASURE] + numeric_cols
measure_default = hinted[0] if hinted else NO_MEASURE
c1, c2 = st.columns(2)
with c1:
    measure_sel = st.selectbox("Coluna com o **número de doses** (o valor a somar)", measure_opts,
                               index=measure_opts.index(measure_default), key=f"measure_{sig}",
                               help="Se os arquivos têm uma linha por dose/pessoa, escolha '(nenhuma)': o app contará as linhas.")
measure = None if measure_sel == NO_MEASURE else measure_sel
unit = "doses" if measure else "registros"

dim_candidates = [c for c in common_cols if c != measure and kinds_std.get(c) in ("texto", "faixa_etaria")]
n_rows_ref = max(len(df_a_all), len(df_b_all))
dim_default = [c for c in dim_candidates
               if not (n_rows_ref > 50 and df_a_all[c].nunique() > 0.5 * len(df_a_all))]
with c2:
    dims = st.multiselect("Separar a comparação por (categorias)", dim_candidates, default=dim_default, key=f"dims_{sig}",
                          help="Faixa etária, sexo, bairro, unidade de saúde, dose, período (mês)… "
                               "Só aparecem as colunas que existem nos dois arquivos.")

# linhas de "Total" (evitar contar duas vezes)
text_cols = [c for c in common_cols if kinds_std.get(c) in ("texto", "faixa_etaria")]
tot_a, tot_b = core.find_total_rows(df_a_all, text_cols), core.find_total_rows(df_b_all, text_cols)
ignore_totals = False
if tot_a.any() or tot_b.any():
    ignore_totals = st.checkbox(
        f"Ignorar linhas de total ({int(tot_a.sum())} em A, {int(tot_b.sum())} em B) para não contar em dobro",
        value=True, key=f"tot_{sig}")
mask_a = tot_a if ignore_totals else pd.Series(False, index=df_a_all.index)
mask_b = tot_b if ignore_totals else pd.Series(False, index=df_b_all.index)
ignored_rows = {"A": df_a_all[mask_a], "B": df_b_all[mask_b]}
df_a, df_b = df_a_all[~mask_a], df_b_all[~mask_b]
if ignore_totals:
    st.caption("Linhas ignoradas ficam listadas na aba *Qualidade dos dados*.")

key_default = core.suggest_key(df_a, df_b, [c for c in common_cols if c != measure], dims)
key_opts = [c for c in common_cols if not c.endswith("(mês)")]
key_cols = st.multiselect(
    "Chave para identificar o **mesmo registro** nos dois arquivos (opcional)", key_opts,
    default=[c for c in key_default if c in key_opts], key=f"key_{sig}_{'|'.join(key_default)}",
    help="Ex.: um número de identificação. Só use se ela for única e preenchida em todos os registros; "
         "caso contrário, o app não compara registro a registro (para não gerar conclusões erradas).")

if not dims and not measure:
    st.warning("Escolha ao menos uma coluna de categorias ou a coluna de doses para ver comparações.")

# Cálculos principais
dim_kinds = {d: kinds_std.get(d, "texto") for d in dims}
tables = {d: core.compare_by_dimension(df_a, df_b, d, measure, dim_kinds[d]) for d in dims}
first_dim = dims[0] if dims else None
totals = core.summarize_totals(df_a, df_b, measure, tables.get(first_dim))


def view_table(t: pd.DataFrame) -> pd.DataFrame:
    """Renomeia colunas genéricas (A/B) para os nomes escolhidos pelo usuário."""
    ren = {"Valor A": f"{unit.capitalize()} – {label_a}", "Valor B": f"{unit.capitalize()} – {label_b}",
           "Registros A": f"Linhas – {label_a}", "Registros B": f"Linhas – {label_b}",
           "Diferença (B − A)": f"Diferença ({label_b} − {label_a})",
           "% do total A": f"% do total – {label_a}", "% do total B": f"% do total – {label_b}",
           "Só no arquivo A": f"Só em {label_a}", "Só no arquivo B": f"Só em {label_b}"}
    out = t.rename(columns=ren)
    if "Situação" in out:
        out["Situação"] = out["Situação"].replace({"Só no arquivo A": f"Só em {label_a}", "Só no arquivo B": f"Só em {label_b}"})
    return out


def pct_config(df: pd.DataFrame) -> dict:
    cfg = {}
    for c in df.columns:
        if c.startswith("Variação %") or c.startswith("% do total") or "p.p." in c:
            cfg[c] = st.column_config.NumberColumn(c, format="%.1f")
    return cfg


# --------------------------------------------------------------------------- #
# Resultados
# --------------------------------------------------------------------------- #
st.header("4. Resultados")
tab_res, tab_cat, tab_rec, tab_q, tab_exp = st.tabs(
    ["📋 Resumo", "🔎 Por categoria", "🧾 Registro a registro", "🧼 Qualidade dos dados", "⬇️ Exportar"])

# ---- Resumo ----
with tab_res:
    m1, m2, m3, m4, m5 = st.columns(5)
    m1.metric(f"Linhas – {label_a}", core.fmt_num(totals["registros_a"]))
    m2.metric(f"Linhas – {label_b}", core.fmt_num(totals["registros_b"]))
    m3.metric(f"Total de {unit} – {label_a}", core.fmt_num(totals["total_a"]))
    m4.metric(f"Total de {unit} – {label_b}", core.fmt_num(totals["total_b"]))
    delta = None if pd.isna(totals["pct"]) else f"{core.fmt_num(totals['pct'], 1, signed=True)}%"
    m5.metric("Diferença (B − A)", core.fmt_num(totals["dif"], signed=True), delta=delta, delta_color="off")
    if not measure:
        st.caption("Sem coluna de doses: o total acima é a **contagem de linhas**.")
    if pd.isna(totals["pct"]) and not pd.isna(totals["dif"]):
        st.caption("Variação percentual indisponível porque o total do arquivo A é zero.")

    st.subheader("Em palavras simples")
    if first_dim:
        for line in core.plain_summary(tables[first_dim], totals, label_a, label_b, first_dim, unit):
            st.markdown(f"- {line}")
        if dims[1:]:
            st.caption(f"Resumo feito com base em **{first_dim}**. Veja as demais categorias na aba *Por categoria*.")
    else:
        for line in core.plain_summary(pd.DataFrame(columns=["Situação"]), totals, label_a, label_b, "—", unit)[:1]:
            st.markdown(f"- {line}")

    if first_dim and totals.get("n_so_a", 0) + totals.get("n_so_b", 0) > 0:
        st.warning(
            f"Parte da diferença total vem de categorias que existem em só um dos arquivos "
            f"({core.fmt_num(totals.get('so_a'))} só em A; {core.fmt_num(totals.get('so_b'))} só em B). "
            f"Nas categorias presentes nos dois, a diferença é de {core.fmt_num(totals.get('dif_comparaveis'), signed=True)}.")
    if first_dim:
        st.subheader(f"Visão rápida – por {first_dim}")
        t = tables[first_dim]
        dcol = "Diferença (B − A)"
        fig = go.Figure(go.Bar(x=t["Categoria"], y=t[dcol],
                               marker_color=[COLOR_UP if (v or 0) >= 0 else COLOR_DOWN for v in t[dcol].fillna(0)],
                               hovertemplate="%{x}<br>%{y:+,.0f}<extra></extra>"))
        fig.update_layout(title=f"Diferença ({label_b} − {label_a}) por {first_dim}", yaxis_title=f"Diferença em {unit}",
                          xaxis_title=None, height=380, margin=dict(t=60, b=40))
        show_chart(fig)
        st.caption(f"Barras para cima: {label_b} tem mais {unit}. Barras para baixo: {label_b} tem menos.")

# ---- Por categoria ----
with tab_cat:
    if not dims:
        st.info("Escolha ao menos uma coluna de categorias no passo 3.")
    else:
        d_sel = st.selectbox("Comparar por", dims, key="dim_view")
        sort_opt = st.radio("Ordenar por", ["Ordem original", "Maior diferença (valor absoluto)", "Maior variação %",
                                            "Ordem alfabética"], horizontal=True, key="sort_view")
        t = tables[d_sel].copy()
        if sort_opt.startswith("Maior diferença"):
            t = t.reindex(t["Diferença (B − A)"].abs().sort_values(ascending=False, na_position="last").index)
        elif sort_opt.startswith("Maior variação"):
            t = t.reindex(t["Variação %"].abs().sort_values(ascending=False, na_position="last").index)
        elif sort_opt.startswith("Ordem alfa"):
            t = t.sort_values("Categoria", key=lambda s: s.map(core.norm_text))
        tot_d = core.summarize_totals(df_a, df_b, measure, tables[d_sel])
        for line in core.plain_summary(tables[d_sel], tot_d, label_a, label_b, d_sel, unit):
            st.markdown(f"- {line}")

        view = st.radio("Gráfico", ["Valores lado a lado", "Diferença", "Participação no total (%)"], horizontal=True, key="chart_view")
        n = len(t)
        horizontal = n > 12
        height = max(420, 24 * n + 120) if horizontal else 460
        fig = go.Figure()
        if view == "Valores lado a lado":
            series = [(label_a, "Valor A", COLOR_A), (label_b, "Valor B", COLOR_B)]
        elif view == "Participação no total (%)":
            series = [(label_a, "% do total A", COLOR_A), (label_b, "% do total B", COLOR_B)]
        else:
            series = None
        if series:
            for name, col, color in series:
                if horizontal:
                    fig.add_bar(y=t["Categoria"], x=t[col], name=name, marker_color=color, orientation="h")
                else:
                    fig.add_bar(x=t["Categoria"], y=t[col], name=name, marker_color=color)
            fig.update_layout(barmode="group")
        else:
            colors = [COLOR_UP if (v or 0) >= 0 else COLOR_DOWN for v in t["Diferença (B − A)"].fillna(0)]
            if horizontal:
                fig.add_bar(y=t["Categoria"], x=t["Diferença (B − A)"], marker_color=colors, orientation="h", name="Diferença")
            else:
                fig.add_bar(x=t["Categoria"], y=t["Diferença (B − A)"], marker_color=colors, name="Diferença")
            fig.update_layout(showlegend=False)
        if horizontal:
            fig.update_yaxes(autorange="reversed", title=None)
        fig.update_layout(height=height, margin=dict(t=30, b=40), legend=dict(orientation="h", y=1.08))
        show_chart(fig)
        if view == "Participação no total (%)":
            st.caption("Participação mostra o **peso** de cada categoria dentro do seu próprio arquivo – útil quando os totais são diferentes.")
        if view == "Diferença":
            st.caption(f"Positivo: {label_b} tem mais {unit}; negativo: tem menos. Categorias que existem só em um arquivo não têm barra.")

        tv = view_table(t)
        show_df(pretty(tv), hide_index=True, column_config=pct_config(tv))

        only_a_t = t[t["Situação"] == "Só no arquivo A"]
        only_b_t = t[t["Situação"] == "Só no arquivo B"]
        nov_t = t[t["Situação"] == "Sem valor em um dos arquivos"]
        ca_, cb_ = st.columns(2)
        with ca_:
            st.markdown(f"**Categorias só em {label_a}:** {len(only_a_t)}")
            if len(only_a_t):
                st.write(", ".join(only_a_t["Categoria"]))
        with cb_:
            st.markdown(f"**Categorias só em {label_b}:** {len(only_b_t)}")
            if len(only_b_t):
                st.write(", ".join(only_b_t["Categoria"]))
        if len(nov_t):
            st.warning("Categorias que existem nos dois arquivos, mas **sem valor preenchido** em um deles "
                       "(ausente ≠ zero): " + ", ".join(nov_t["Categoria"]))

# ---- Registro a registro ----
records = None
with tab_rec:
    st.markdown("Compara registro por registro usando a **chave** escolhida no passo 3. "
                f"**Adicionado** = existe em {label_b} e não em {label_a}; **removido** = existe em {label_a} e não em {label_b}; "
                "**alterado** = mesma chave, valores diferentes.")
    if not key_cols:
        st.info("Nenhuma chave escolhida (ou nenhuma chave confiável foi encontrada). Sem uma chave, não dá para dizer com "
                "segurança se um registro foi adicionado, removido ou alterado – por isso essa comparação fica desativada. "
                "Use a aba *Por categoria* para ver as diferenças por grupo.")
    else:
        qa, qb = core.key_quality(df_a, key_cols), core.key_quality(df_b, key_cols)
        if not (qa["reliable"] and qb["reliable"]):
            st.error("A chave escolhida **não é confiável**, então a comparação registro a registro foi bloqueada.")
            st.write(f"Arquivo A: {qa['reason'] or 'ok'}  \nArquivo B: {qb['reason'] or 'ok'}")
            for lbl, d, q in ((label_a, df_a, qa), (label_b, df_b, qb)):
                if q["dups"] or q["missing"]:
                    with st.expander(f"Ver registros problemáticos em {lbl}"):
                        if q["missing"]:
                            st.markdown("Chave vazia:")
                            show_df(pretty(d[d[key_cols].isna().any(axis=1)]))
                        if q["dups"]:
                            st.markdown("Chave repetida:")
                            show_df(pretty(core.duplicated_keys(d, key_cols)))
        else:
            compare_cols = [c for c in common_cols if not c.endswith("(mês)")]
            records = core.compare_records(df_a, df_b, key_cols, compare_cols)
            r1, r2, r3, r4 = st.columns(4)
            r1.metric("Adicionados", core.fmt_num(records["n_added"]))
            r2.metric("Removidos", core.fmt_num(records["n_removed"]))
            r3.metric("Alterados", core.fmt_num(records["n_changed"]))
            r4.metric("Iguais", core.fmt_num(records["n_same"]))
            for title, key_, text in (("Alterados", "changed", "Mesmo registro, valor diferente"),
                                      ("Adicionados", "added", f"Só em {label_b}"),
                                      ("Removidos", "removed", f"Só em {label_a}")):
                with st.expander(f"{title} ({len(records[key_])} linhas) – {text}"):
                    if records[key_].empty:
                        st.write("Nenhum.")
                    else:
                        show_df(pretty(records[key_]), hide_index=True)

# ---- Qualidade ----
with tab_q:
    st.markdown("Aqui está tudo o que o app **corrigiu, converteu ou ignorou** – e os registros que merecem atenção.")
    key_for_dups = key_cols or dims
    for lbl, side, std, df_all, df_used, ign in (
            (label_a, side_a, std_a, df_a_all, df_a, ignored_rows["A"]),
            (label_b, side_b, std_b, df_b_all, df_b, ignored_rows["B"])):
        st.subheader(f"Arquivo: {lbl}")
        with st.container(border=True):
            info = side["info"]
            st.markdown(f"- Linhas de dados lidas: **{len(side['table'])}** (cabeçalho na linha {info.header_row_excel}); "
                        f"linhas totalmente vazias ignoradas: **{info.empty_rows_dropped}**; "
                        f"colunas totalmente vazias ignoradas: **{info.empty_cols_dropped}**.")
            st.markdown(f"- Linhas de total ignoradas: **{len(ign)}**.")
            if len(ign):
                show_df(pretty(ign))
            for col_name, orig in info.original_headers.items():
                if orig.startswith("="):
                    st.markdown(f"- Cabeçalho original de **{col_name}** era uma fórmula do sistema de origem: `{orig}`")
            show_df(pretty(core.profile_columns(std.df, std.log)), hide_index=True)
            if measure and measure in df_used:
                ms = df_used[measure]
                st.markdown(f"- Coluna **{measure}**: **{int(ms.isna().sum())}** vazio(s) (ausente) e **{int((ms == 0).sum())}** zero(s) "
                            "(valor 0 informado) – são tratados de forma diferente.")
                if ms.isna().any():
                    with st.expander("Ver linhas com valor ausente"):
                        show_df(pretty(df_used[ms.isna()]))
            if std.label_variants:
                st.markdown("- Rótulos que só diferiam por maiúsculas/minúsculas, acentos ou espaços foram tratados como o **mesmo** valor:")
                for col_name, groups in std.label_variants.items():
                    for _k, variants in groups.items():
                        st.markdown(f"    - {col_name}: " + " / ".join(f"`{v}`" for v in variants))
            dup_full = core.exact_duplicates(df_used)
            st.markdown(f"- Linhas **totalmente repetidas**: **{len(dup_full)}**")
            if len(dup_full):
                with st.expander("Ver linhas totalmente repetidas (nada foi apagado)"):
                    show_df(pretty(dup_full))
            if key_for_dups:
                dk = core.duplicated_keys(df_used, key_for_dups)
                what = "chave" if key_cols else "combinação de categorias"
                st.markdown(f"- Linhas com **{what} repetida** ({', '.join(key_for_dups)}): **{len(dk)}**"
                            + (" – nesses casos o app **soma** os valores dentro da mesma categoria." if len(dk) and measure else ""))
                if len(dk):
                    with st.expander(f"Ver linhas com {what} repetida"):
                        show_df(pretty(dk))
            st.caption("Repetição não significa, necessariamente, erro: em tabelas já resumidas por categoria ela pode ser legítima.")

# ---- Exportar ----
with tab_exp:
    st.markdown("Baixe todos os resultados em uma planilha Excel (resumo, tabelas por categoria, registros adicionados/removidos/alterados, "
                "duplicatas e o mapeamento de colunas usado).")
    resumo = pd.DataFrame({
        "Indicador": ["Linhas", "Linhas", f"Total de {unit}", f"Total de {unit}", "Diferença (B − A)", "Variação %"],
        "Arquivo": [label_a, label_b, label_a, label_b, f"{label_b} − {label_a}", f"{label_b} − {label_a}"],
        "Valor": [totals["registros_a"], totals["registros_b"], totals["total_a"], totals["total_b"], totals["dif"], totals["pct"]],
    })
    sheets: dict[str, pd.DataFrame] = {"Resumo": resumo}
    for d, t in tables.items():
        sheets[f"Por {d}"] = pretty(view_table(t))
    if records:
        for title, key_ in (("Alterados", "changed"), ("Adicionados", "added"), ("Removidos", "removed")):
            if not records[key_].empty:
                sheets[title] = pretty(records[key_])
    for tag, d in (("A", df_a), ("B", df_b)):
        dup = core.exact_duplicates(d)
        if not dup.empty:
            sheets[f"Repetidas {tag}"] = pretty(dup)
    for tag, ign in ignored_rows.items():
        if not ign.empty:
            sheets[f"Total ignorado {tag}"] = pretty(ign)
    sheets["Mapeamento de colunas"] = pd.DataFrame(
        [{"Coluna em A": a, "Coluna em B": b, "Tipo usado": core.KIND_LABELS.get(final_kinds[a], final_kinds[a])} for a, b in pairs])
    st.download_button("⬇️ Baixar resultados (Excel)", data=to_excel_bytes(sheets),
                       file_name="comparacao_vacinacao_dengue_taubate.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")