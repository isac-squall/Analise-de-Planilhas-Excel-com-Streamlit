# -*- coding: utf-8 -*-
# Aplicação Streamlit para análise de planilhas Excel

import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import openpyxl
from openpyxl import load_workbook  
from openpyxl.utils.dataframe import dataframe_to_rows
import tempfile


st.title("Análise de Planilhas Excel com Streamlit")
st.write("""
Esta aplicação permite:
- Carregar uma planilha Excel
- Visualizar abas e escolher a linha de cabeçalho
- Limpar e equalizar os dados (remover duplicatas, linhas/colunas vazias)
- Apurar carteiras, totais e pessoas físicas por unidade de operação
- Verificar inconsistências
- Visualizar dashboards de desempenho
""")

uploaded_file = st.file_uploader("Escolha um arquivo Excel", type=["xlsx"])
if uploaded_file is not None:
    # Salva o arquivo temporariamente
    temp_filename = 'planilha_temp.xlsx'
    with open(temp_filename, 'wb') as f:
        f.write(uploaded_file.getbuffer())
    st.success("Arquivo carregado com sucesso!")

    # Carrega o arquivo Excel
    excel_file = pd.ExcelFile(temp_filename)
    st.write("Abas disponíveis:", excel_file.sheet_names)

    # Visualiza as primeiras linhas da aba principal
    aba_padrao = st.selectbox("Selecione a aba para análise", excel_file.sheet_names)
    base_dados = pd.read_excel(excel_file, sheet_name=aba_padrao)
    st.write("Primeiras 15 linhas da base de dados:")
    st.dataframe(base_dados.head(15))

    # Escolha da linha de cabeçalho
    linha_cabecalho = st.number_input(
        "Informe o número da linha que contém os nomes das colunas (começando do 0)", 
        min_value=0, max_value=min(50, len(base_dados)-1), value=0
    )

    # Carrega novamente com o cabeçalho correto
    df = pd.read_excel(excel_file, sheet_name=aba_padrao, header=linha_cabecalho)

    # Limpeza e equalização
    df.columns = [str(col).strip().title() for col in df.columns]

    # Garante nomes únicos para as colunas
    def make_unique(cols):
        seen = {}
        result = []
        for col in cols:
            if col not in seen:
                seen[col] = 0
                result.append(col)
            else:
                seen[col] += 1
                result.append(f"{col}_{seen[col]}")
        return result
    df.columns = make_unique(df.columns)

    # Limpa espaços e remove duplicatas/linhas/colunas vazias
    df = df.applymap(lambda x: x.strip() if isinstance(x, str) else x)
    df = df.drop_duplicates()
    df = df.dropna(how='all').reset_index(drop=True)
    df = df.dropna(axis=1, how='all')

    st.write("Colunas disponíveis:", df.columns.tolist())
    st.write("Base de dados limpa:")
    st.dataframe(df)

    # Apuração de carteiras, totais e pessoas físicas por unidade de operação
    st.subheader("Apuração por Unidade de Operação")
    col_unidade = st.selectbox("Selecione a coluna de Unidade", df.columns)
    col_carteira = st.selectbox("Selecione a coluna de Carteira", df.columns)
    col_pessoa = st.selectbox("Selecione a coluna de Pessoa Física", df.columns)

    resumo = df.groupby([col_unidade, col_carteira], as_index=False)[col_pessoa].count().rename(columns={col_pessoa: "Total Pessoas"})
    st.dataframe(resumo)

    # Total por unidade (independente da carteira)
    resumo_unidade = df.groupby(col_unidade, as_index=False)[col_pessoa].count().rename(columns={col_pessoa: "Total Pessoas"})
    st.subheader("Total por Unidade de Operação")
    st.dataframe(resumo_unidade)

    # Verificação de inconsistências
    st.subheader("Inconsistências Encontradas")
    inconsistencias = []
    if df.isnull().values.any():
        inconsistencias.append("Existem valores ausentes na base.")
    if df.duplicated().any():
        inconsistencias.append("Existem linhas duplicadas na base.")
    if len(inconsistencias) == 0:
        st.success("Nenhuma inconsistência encontrada.")
    else:
        for inc in inconsistencias:
            st.warning(inc)

    # Dashboard de desempenho das carteiras
    st.subheader("Dashboard de Desempenho das Carteiras")
    tipo_grafico = st.selectbox("Tipo de gráfico", ["Barras", "Pizza"])
    col_valor = st.selectbox("Selecione a coluna de valor para análise (ex: Vendas, Total, etc)", df.columns)

    if tipo_grafico == "Barras":
        fig = px.bar(
            df,
            x=col_carteira,
            y=col_valor,
            color=col_unidade,
            title='Desempenho das Carteiras por Unidade',
            barmode='group'
        )
        st.plotly_chart(fig, use_container_width=True)
    else:
        resumo_pizza = df.groupby(col_carteira, as_index=False)[col_valor].sum()
        resumo_pizza[col_valor] = pd.to_numeric(resumo_pizza[col_valor], errors='coerce')
        resumo_pizza = resumo_pizza.dropna(subset=[col_valor])
        if resumo_pizza.empty:
            st.warning("Não há dados numéricos suficientes para gerar o gráfico de pizza.")
        else:
            fig = px.pie(
                resumo_pizza,
                names=col_carteira,
                values=col_valor,
                title='Participação das Carteiras'
            )
            st.plotly_chart(fig, use_container_width=True)

    # Dashboard: Valor total por unidade de operação
    st.subheader("Dashboard: Valor Total por Unidade de Operação")
    col_valor_unidade = st.selectbox("Selecione a coluna de valor para o dashboard por unidade", df.columns, key="valor_unidade")
    df[col_valor_unidade] = pd.to_numeric(df[col_valor_unidade], errors='coerce')
    resumo_valor_unidade = df.groupby(col_unidade, as_index=False)[col_valor_unidade].sum()

    fig_valor_unidade = px.bar(
        resumo_valor_unidade,
        x=col_unidade,
        y=col_valor_unidade,
        title="Valor Total por Unidade de Operação",
        color=col_unidade,
        text_auto=True
    )
    st.plotly_chart(fig_valor_unidade, use_container_width=True)





