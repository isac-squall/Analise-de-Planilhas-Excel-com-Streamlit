# -*- coding: utf-8 -*-
"""Aplicacao Streamlit para analise de planilhas Excel e CSV.

Funcionalidades:
- Carregar arquivos .xlsx ou .csv (com deteccao de encoding e delimitador)
- Visualizar dados brutos e escolher a linha de cabecalho
- Limpar e equalizar dados (espacos, duplicatas, linhas/colunas vazias)
- Converter automaticamente colunas numericas
- Apurar carteiras, totais e pessoas por unidade de operacao
- Gerar dashboards dinamicos (barras, pizza, linha, dispersao)
- Relatorio de qualidade de dados e download dos resultados
"""

import csv
import io
import os
import tempfile
from pathlib import Path

import pandas as pd
import plotly.express as px
import streamlit as st

# --------------------------------------------------------------------------- #
# Helpers puros (sem dependencia do Streamlit) - testaveis em isolamento
# --------------------------------------------------------------------------- #

ENCODINGS = ["utf-8-sig", "utf-8", "cp1252", "latin-1"]
DELIMITERS = ";,\t|"


def _br_to_float(value):
    """Converte um valor isolado no formato brasileiro para float.

    Regra: o ultimo separador (',' ou '.') e o decimal; os demais sao
    separadores de milhar. So converte quando ha pelo menos uma virgula.
    """
    if not isinstance(value, str):
        return None
    v = value.strip()
    if "," not in v:
        return None
    last_sep = max(v.rfind(","), v.rfind("."))
    if last_sep < 0:
        return None
    integer_part = v[:last_sep]
    decimal_part = v[last_sep + 1 :]
    integer_clean = integer_part.replace(".", "").replace(",", "")
    try:
        return float(f"{integer_clean}.{decimal_part}")
    except ValueError:
        return None


def detect_encoding(raw: bytes) -> str:
    """Retorna o encoding mais provavel para os bytes fornecidos."""
    for enc in ENCODINGS:
        try:
            raw.decode(enc)
            return enc
        except UnicodeDecodeError:
            continue
    return "latin-1"


def detect_delimiter(text: str) -> str:
    """Detecta o delimitador mais provavel em um CSV."""
    try:
        return csv.Sniffer().sniff(text[:8192], delimiters=DELIMITERS).delimiter
    except Exception:
        return ";"


def make_unique_columns(cols) -> list:
    """Garante nomes de coluna unicos, preservando os nomes originais."""
    seen = {}
    result = []
    for i, col in enumerate(cols):
        col = str(col).strip()
        if col == "":
            col = f"Coluna_sem_nome_{i}"
        count = seen.get(col, 0)
        seen[col] = count + 1
        result.append(col if count == 0 else f"{col}_{count}")
    return result


def clean_columns(raw: pd.DataFrame, header_idx: int) -> pd.DataFrame:
    """Usa a linha escolhida como cabecalho e limpa os nomes das colunas."""
    if header_idx >= len(raw):
        header_idx = 0
    df = raw.iloc[header_idx:].copy()
    df.columns = [str(c) if pd.notna(c) else "" for c in df.iloc[0]]
    df = df.iloc[1:].reset_index(drop=True)
    df.columns = make_unique_columns(df.columns)
    return df


def _to_float(value):
    """Tenta converter um valor isolado para float (padrao ou formato BR)."""
    try:
        return float(value)
    except (TypeError, ValueError):
        return _br_to_float(value)


def try_convert_numeric(df: pd.DataFrame, threshold: float = 0.8) -> pd.DataFrame:
    """Converte colunas de texto que sejam majoritariamente numericas.

    Para cada valor tenta o formato padrao e, na falha, o formato brasileiro
    (virgula decimal / ponto de milhar), comum em planilhas nacionais.
    """
    out = df.copy()
    for col in df.columns:
        dtype = out[col].dtype
        if not (pd.api.types.is_object_dtype(dtype) or pd.api.types.is_string_dtype(dtype)):
            continue
        non_null = out[col].dropna()
        if len(non_null) == 0:
            continue
        parsed = non_null.map(_to_float)
        if parsed.notna().mean() >= threshold:
            out[col] = out[col].map(_to_float)
    return out


def clean_dataframe(
    df: pd.DataFrame,
    strip_ws: bool = True,
    drop_dup: bool = True,
    drop_empty_rows: bool = True,
    drop_empty_cols: bool = True,
    convert_numeric: bool = True,
) -> pd.DataFrame:
    """Aplica a equalizacao configurada na base de dados."""
    df = df.copy()
    if strip_ws:
        df = df.map(lambda x: x.strip() if isinstance(x, str) else x)
    if drop_dup:
        df = df.drop_duplicates()
    if drop_empty_rows:
        df = df.dropna(how="all").reset_index(drop=True)
    if drop_empty_cols:
        df = df.dropna(axis=1, how="all")
    if convert_numeric:
        df = try_convert_numeric(df)
    return df.reset_index(drop=True)


def numeric_columns(df: pd.DataFrame) -> list:
    """Colunas com tipo numerico (ou category numerica)."""
    return df.select_dtypes(include=["number"]).columns.tolist()


def categorical_columns(df: pd.DataFrame) -> list:
    """Colunas nao numericas (texto, datas, booleanos)."""
    return [
        c
        for c in df.columns
        if c not in set(numeric_columns(df)) and c not in {"", "index"}
    ]


def guess_group_col(df: pd.DataFrame) -> str:
    """Sugere uma coluna de agrupamento (baixa cardinalidade)."""
    cols = categorical_columns(df) + numeric_columns(df)
    best, best_score = cols[0] if cols else df.columns[0], None
    for col in cols:
        s = df[col].dropna()
        if len(s) == 0:
            continue
        score = s.nunique() / max(len(s), 1)
        if best_score is None or score < best_score:
            best, best_score = col, score
    return best


def guess_ident_col(df: pd.DataFrame) -> str:
    """Sugere uma coluna de identificador (alta cardinalidade)."""
    cols = categorical_columns(df) + numeric_columns(df)
    if not cols:
        return df.columns[0]
    return max(cols, key=lambda c: df[c].nunique(dropna=True))


def apurar(
    df: pd.DataFrame,
    col_unidade: str,
    col_carteira: str,
    col_pessoa: str,
    modo: str = "Contagem",
) -> pd.DataFrame:
    """Conta pessoas por (unidade, carteira)."""
    if col_unidade == col_carteira:
        col_carteira = col_pessoa
    if modo == "Valores unicos":
        res = (
            df.groupby([col_unidade, col_carteira], dropna=False)[col_pessoa]
            .nunique()
            .reset_index(name="Total")
        )
    else:
        res = (
            df.groupby([col_unidade, col_carteira], dropna=False)[col_pessoa]
            .count()
            .reset_index(name="Total")
        )
    return res.sort_values("Total", ascending=False).reset_index(drop=True)


def build_chart(df: pd.DataFrame, chart_type: str, col_cat: str, col_num: str, col_color: str = None):
    """Monta o grafico solicitado a partir de colunas selecionadas."""
    dfn = df.copy()
    dfn[col_num] = pd.to_numeric(dfn[col_num], errors="coerce")
    dfn = dfn.dropna(subset=[col_num])
    if dfn.empty:
        return None
    if col_color and col_color == col_cat:
        col_color = None

    if chart_type == "Pizza":
        agg = dfn.groupby(col_cat, dropna=False)[col_num].sum().reset_index()
        agg = agg.sort_values(col_num, ascending=False)
        return px.pie(
            agg, names=col_cat, values=col_num, title="Participacao por Categoria"
        )

    if chart_type == "Linha":
        agg = dfn.groupby(col_cat, dropna=False)[col_num].sum().reset_index()
        return px.line(
            agg, x=col_cat, y=col_num, markers=True, title="Evolucao por Categoria"
        )

    if chart_type == "Dispersao":
        return px.scatter(
            dfn,
            x=dfn[col_num],
            y=dfn.index,
            color=col_cat if col_cat in dfn.columns else None,
            title=f"Distribuicao de {col_num}",
        )

    if col_color is not None:
        agg = (
            dfn.groupby([col_cat, col_color], dropna=False)[col_num]
            .sum()
            .reset_index()
        )
        return px.bar(
            agg,
            x=col_cat,
            y=col_num,
            color=col_color,
            title=f"Desempenho por {col_cat}",
            barmode="group",
            text_auto=True,
        )

    agg = dfn.groupby(col_cat, dropna=False)[col_num].sum().reset_index()
    return px.bar(
        agg,
        x=col_cat,
        y=col_num,
        color=col_cat,
        title=f"Desempenho por {col_cat}",
        barmode="group",
        text_auto=True,
    )


def to_csv_bytes(df: pd.DataFrame) -> bytes:
    return df.to_csv(index=False).encode("utf-8-sig")


def to_excel_bytes(df: pd.DataFrame) -> bytes:
    buf = io.BytesIO()
    df.to_excel(buf, index=False, engine="openpyxl")
    return buf.getvalue()


# --------------------------------------------------------------------------- #
# Carregamento com cache
# --------------------------------------------------------------------------- #

@st.cache_data(show_spinner=False)
def load_excel_sheets(data: bytes):
    """Le todas as abas de um arquivo Excel como dados brutos (sem cabecalho)."""
    with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as tmp:
        tmp.write(data)
        tmp_path = tmp.name
    try:
        xl = pd.ExcelFile(tmp_path)
        raw_by_sheet = {
            name: pd.read_excel(tmp_path, sheet_name=name, header=None)
            for name in xl.sheet_names
        }
        return xl.sheet_names, raw_by_sheet
    finally:
        os.unlink(tmp_path)


@st.cache_data(show_spinner=False)
def load_csv_raw(data: bytes):
    """Le um CSV como dados brutos, detectando encoding e delimitador."""
    encoding = detect_encoding(data)
    text = data.decode(encoding, errors="replace")
    sep = detect_delimiter(text)
    df = pd.read_csv(io.StringIO(text), sep=sep, header=None, engine="python")
    return df, sep, encoding


# --------------------------------------------------------------------------- #
# Interface
# --------------------------------------------------------------------------- #

def main():
    st.set_page_config(page_title="Analisador de Planilhas", layout="wide")

    st.title("Analisador de Planilhas Excel e CSV")
    st.caption(
        "Carregue um arquivo, escolha a aba e a linha de cabecalho, limpe os dados "
        "e explore apuracoes, dashboards e qualidade de dados."
    )

    uploaded = st.file_uploader("Escolha um arquivo Excel (.xlsx) ou CSV", type=["xlsx", "csv"])
    if uploaded is None:
        st.info("Envie um arquivo para comecar.")
        return

    data = uploaded.getvalue()
    ext = Path(uploaded.name).suffix.lower()

    # --------------------------- Leitura do arquivo ------------------------ #
    try:
        if ext == ".csv":
            raw, sep, enc = load_csv_raw(data)
            st.caption(f"CSV detectado: delimitador {sep!r}, encoding {enc}")
            raw_by_sheet = None
        else:
            sheet_names, raw_by_sheet = load_excel_sheets(data)
            sep, enc = None, None
    except Exception as exc:  # noqa: BLE001
        st.error(f"Falha ao ler o arquivo: {exc}")
        return

    # ------------------------ Selecao de aba (Excel) ----------------------- #
    if raw_by_sheet is not None:
        usable = [
            name
            for name, frame in raw_by_sheet.items()
            if not frame.dropna(how="all").empty
        ]
        if not usable:
            st.error("Nenhuma aba com dados encontrada no arquivo.")
            return
        largest = max(usable, key=lambda n: len(raw_by_sheet[n]))
        options = ["Todas as abas (unificar)"] + usable
        choice = st.selectbox(
            "Selecione a aba para analise",
            options,
            index=options.index(largest),
            key="sheet_choice",
        )
        if choice == "Todas as abas (unificar)":
            frames = [
                raw_by_sheet[n] for n in usable if not raw_by_sheet[n].empty
            ]
            if frames:
                max_cols = max(f.shape[1] for f in frames)
                frames = [f.reindex(columns=range(max_cols)) for f in frames]
                raw = pd.concat(frames, ignore_index=True)
            else:
                raw = pd.DataFrame()
        else:
            raw = raw_by_sheet[choice]
    else:
        choice = uploaded.name

    if raw.empty or len(raw) == 0:
        st.error("O arquivo selecionado nao contem linhas de dados.")
        return

    # --------------------- Linha de cabecalho e pre-visualizacao ----------- #
    max_header = max(0, len(raw) - 1)
    header_idx = st.number_input(
        "Linha que contem os nomes das colunas (0 = primeira linha)",
        min_value=0,
        max_value=min(50, max_header),
        value=0,
        step=1,
        key="header_row",
    )

    st.write("**Previa dos dados brutos** - escolha a linha do cabecalho acima:")
    st.dataframe(raw.head(15), width="stretch")

    # ---------------------------- Limpeza de dados ------------------------- #
    with st.sidebar:
        st.header("Opcoes de limpeza")
        opt_strip = st.checkbox("Remover espacos em branco", value=True, key="opt_strip")
        opt_dedup = st.checkbox("Remover linhas duplicadas", value=True, key="opt_dedup")
        opt_rows = st.checkbox("Remover linhas totalmente vazias", value=True, key="opt_rows")
        opt_cols = st.checkbox("Remover colunas totalmente vazias", value=True, key="opt_cols")
        opt_num = st.checkbox("Converter colunas numericas", value=True, key="opt_num")

    try:
        base = clean_columns(raw, header_idx)
        df = clean_dataframe(
            base,
            strip_ws=opt_strip,
            drop_dup=opt_dedup,
            drop_empty_rows=opt_rows,
            drop_empty_cols=opt_cols,
            convert_numeric=opt_num,
        )
    except Exception as exc:  # noqa: BLE001
        st.error(f"Erro na limpeza dos dados: {exc}")
        return

    if df.empty:
        st.warning("A base ficou vazia apos a limpeza. Ajuste as opcoes.")
        return

    num_cols = numeric_columns(df)
    cat_cols = categorical_columns(df)

    def _default_index(col):
        if col in df.columns:
            return df.columns.get_loc(col)
        return 0

    default_unidade = guess_group_col(df)
    default_carteira = next(
        (c for c in cat_cols + num_cols if c != default_unidade),
        default_unidade,
    )
    default_pessoa = guess_ident_col(df)

    col_unidade = st.selectbox(
        "Coluna de Unidade", df.columns, key="col_unidade",
        index=_default_index(default_unidade),
    )
    col_carteira = st.selectbox(
        "Coluna de Carteira", df.columns, key="col_carteira",
        index=_default_index(default_carteira),
    )
    col_pessoa = st.selectbox(
        "Coluna de Pessoa / Identificador", df.columns, key="col_pessoa",
        index=_default_index(default_pessoa),
    )

    tab_dados, tab_apuracao, tab_dash, tab_qualidade = st.tabs(
        ["Dados e Limpeza", "Apuracao", "Dashboards", "Qualidade de Dados"]
    )

    # ------------------------------ Aba Dados ------------------------------ #
    with tab_dados:
        c1, c2, c3, c4 = st.columns(4)
        c1.metric("Linhas", f"{len(df):,}")
        c2.metric("Colunas", len(df.columns))
        c3.metric("Valores ausentes", int(df.isna().sum().sum()))
        c4.metric("Linhas duplicadas", int(df.duplicated().sum()))

        st.write(f"**Base de dados limpa** - origem: `{choice}`")
        st.dataframe(df, width="stretch")

        dl1, dl2 = st.columns(2)
        with dl1:
            st.download_button(
                "Baixar dados como CSV",
                to_csv_bytes(df),
                file_name=f"dados_limpos_{choice.replace(' ', '_')}.csv",
                mime="text/csv",
                key="dl_csv",
            )
        with dl2:
            st.download_button(
                "Baixar dados como Excel",
                to_excel_bytes(df),
                file_name=f"dados_limpos_{choice.replace(' ', '_')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="dl_xlsx",
            )

    # ----------------------------- Aba Apuracao ---------------------------- #
    with tab_apuracao:
        st.subheader("Apuracao por Unidade de Operacao e Carteira")
        modo = st.radio("Metodo de apuracao", ["Contagem", "Valores unicos"], horizontal=True,
                        key="modo_apuracao")
        resumo = apurar(df, col_unidade, col_carteira, col_pessoa, modo)
        st.dataframe(resumo, width="stretch")
        st.download_button(
            "Baixar apuracao",
            to_csv_bytes(resumo),
            file_name="apuracao.csv",
            mime="text/csv",
            key="dl_apuracao",
        )

        st.subheader("Total por Unidade de Operacao")
        resumo_unidade = (
            df.groupby(col_unidade, dropna=False)[col_pessoa]
            .count()
            .reset_index(name="Total")
            .sort_values("Total", ascending=False)
        )
        st.dataframe(resumo_unidade, width="stretch")
        st.download_button(
            "Baixar total por unidade",
            to_csv_bytes(resumo_unidade),
            file_name="total_por_unidade.csv",
            mime="text/csv",
            key="dl_total_unidade",
        )

    # ------------------------------ Aba Dashboard -------------------------- #
    with tab_dash:
        st.subheader("Dashboard dinamico")
        if num_cols:
            col_valor = st.selectbox("Coluna de valor (numerica)", num_cols, key="col_valor")
            col_cat = st.selectbox("Coluna de categoria", cat_cols + num_cols, key="col_cat",
                                   index=0)
            col_color = st.selectbox(
                "Coluna de cor (opcional)", ["(nenhuma)"] + df.columns.tolist(),
                key="col_color",
            )
            chart_type = st.selectbox(
                "Tipo de grafico", ["Barras", "Pizza", "Linha", "Dispersao"], key="chart_type"
            )
            color_col = None if col_color == "(nenhuma)" else col_color
            fig = build_chart(df, chart_type, col_cat, col_valor, color_col)
            if fig is None:
                st.warning("Sem dados numericos suficientes para o grafico.")
            else:
                st.plotly_chart(fig, width="stretch")
        else:
            st.warning(
                "Nenhuma coluna numerica encontrada. Ative a conversao numerica "
                "ou envie um arquivo com valores numericos."
            )

        st.subheader("Valor Total por Unidade de Operacao")
        if num_cols:
            col_valor_unidade = st.selectbox(
                "Coluna de valor para a unidade", num_cols, key="col_valor_unidade"
            )
            try:
                dfn = df.copy()
                dfn[col_valor_unidade] = pd.to_numeric(dfn[col_valor_unidade], errors="coerce")
                resumo_valor = (
                    dfn.groupby(col_unidade, dropna=False)[col_valor_unidade]
                    .sum()
                    .reset_index()
                    .sort_values(col_valor_unidade, ascending=False)
                )
                fig2 = px.bar(
                    resumo_valor,
                    x=col_unidade,
                    y=col_valor_unidade,
                    color=col_unidade,
                    text_auto=True,
                    title="Valor Total por Unidade de Operacao",
                )
                st.plotly_chart(fig2, width="stretch")
                st.download_button(
                    "Baixar valor por unidade",
                    to_csv_bytes(resumo_valor),
                    file_name="valor_por_unidade.csv",
                    mime="text/csv",
                    key="dl_valor_unidade",
                )
            except Exception as exc:  # noqa: BLE001
                st.error(f"Nao foi possivel gerar o grafico: {exc}")
        else:
            st.info("Sem coluna numerica disponivel para este dashboard.")

    # -------------------------- Aba Qualidade de Dados --------------------- #
    with tab_qualidade:
        st.subheader("Tipos das colunas")
        tipos = pd.DataFrame(
            {
                "Coluna": df.columns,
                "Tipo": [str(df[c].dtype) for c in df.columns],
                "Valores ausentes": [int(df[c].isna().sum()) for c in df.columns],
                "Percentual ausente": [
                    f"{df[c].isna().mean() * 100:.1f}%" for c in df.columns
                ],
                "Valores unicos": [df[c].nunique(dropna=True) for c in df.columns],
            }
        )
        st.dataframe(tipos, width="stretch")

        st.subheader("Inconsistencias encontradas")
        inconsistencias = []
        if df.isnull().values.any():
            inconsistencias.append("Existem valores ausentes na base.")
        if df.duplicated().any():
            inconsistencias.append("Existem linhas duplicadas na base.")
        vazias = df.columns[df.isna().all()].tolist()
        if vazias:
            inconsistencias.append(f"Colunas totalmente vazias: {vazias}")
        if not inconsistencias:
            st.success("Nenhuma inconsistencia encontrada.")
        else:
            for inc in inconsistencias:
                st.warning(inc)


if __name__ == "__main__":
    main()
