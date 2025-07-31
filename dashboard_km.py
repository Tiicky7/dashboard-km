
import streamlit as st
import pandas as pd
import plotly.express as px

st.set_page_config(page_title="Dashboard KM", layout="wide")
st.title("📊 Dashboard KM - Emissões e Cancelamentos")

# Carregar dados
cancelamentos = pd.read_excel("CANCELAMENTOS_KM.xlsx")
emissoes = pd.read_excel("EMISSOES_KM.xlsx")

# Conversão de datas
cancelamentos["DATA_CANCELADO"] = pd.to_datetime(cancelamentos["DATA_CANCELADO"])
emissoes["DATA_EMISSAO"] = pd.to_datetime(emissoes["DATA_EMISSAO"])

# Filtros
meses = sorted(emissoes["MES"].unique())
expedicoes = sorted(emissoes["EXPEDICAO"].dropna().unique())

mes = st.sidebar.selectbox("📅 Selecione o mês", meses)
expedicao = st.sidebar.selectbox("🚚 Selecione a expedição", expedicoes)

cancel_filtrado = cancelamentos[
    (cancelamentos["MES"] == mes) & (cancelamentos["EXPEDICAO"] == expedicao)
]
emissoes_filtrado = emissoes[
    (emissoes["MES"] == mes) & (emissoes["EXPEDICAO"] == expedicao)
]

# Tabs
aba1, aba2, aba3, aba4 = st.tabs(
    ["📌 Resumo", "📦 Emissões", "❌ Cancelamentos", "📈 Comparativo"]
)

# Aba 1 - Resumo
with aba1:
    st.subheader("📌 Visão Geral do Mês")

    total_emissoes = emissoes_filtrado["CTRC_EMITIDO"].sum()
    total_cancelamentos = len(cancel_filtrado)
    taxa_cancelamento = (
        total_cancelamentos / total_emissoes if total_emissoes > 0 else 0
    )

    col1, col2, col3 = st.columns(3)
    col1.metric("Total de Emissões", total_emissoes)
    col2.metric("Total de Cancelamentos", total_cancelamentos)
    col3.metric("Taxa de Cancelamento (%)", f"{taxa_cancelamento*100:.2f}%")

# Aba 2 - Emissões
with aba2:
    st.subheader("📦 Emissões")

    fig1 = px.bar(
        emissoes_filtrado,
        x="DATA_EMISSAO",
        y="CTRC_EMITIDO",
        title="Emissões por Dia",
    )
    st.plotly_chart(fig1, use_container_width=True)

    emissoes_user = (
        emissoes_filtrado["USUÁRIO"].value_counts().reset_index()
    )
    emissoes_user.columns = ["USUÁRIO", "TOTAL"]

    fig2 = px.pie(emissoes_user, names="USUÁRIO", values="TOTAL", title="Por Usuário")
    st.plotly_chart(fig2, use_container_width=True)

# Aba 3 - Cancelamentos
with aba3:
    st.subheader("❌ Cancelamentos")

    motivos = (
        cancel_filtrado["MOTIVO"].value_counts().reset_index()
    )
    motivos.columns = ["MOTIVO", "TOTAL"]

    fig3 = px.bar(motivos, x="MOTIVO", y="TOTAL", title="Motivos de Cancelamento")
    st.plotly_chart(fig3, use_container_width=True)

    canc_user = (
        cancel_filtrado["USUARIO"].value_counts().reset_index()
    )
    canc_user.columns = ["USUARIO", "TOTAL"]

    fig4 = px.pie(canc_user, names="USUARIO", values="TOTAL", title="Por Usuário")
    st.plotly_chart(fig4, use_container_width=True)

# Aba 4 - Comparativo
with aba4:
    st.subheader("📈 Emissões x Cancelamentos (Linha do tempo)")

    emi_dia = (
        emissoes_filtrado.groupby("DATA_EMISSAO")["CTRC_EMITIDO"]
        .sum()
        .reset_index()
    )
    canc_dia = (
        cancel_filtrado.groupby("DATA_CANCELADO")
        .size()
        .reset_index(name="CANCELAMENTOS")
    )

    df_merge = pd.merge(
        emi_dia, canc_dia, left_on="DATA_EMISSAO", right_on="DATA_CANCELADO", how="outer"
    ).fillna(0)

    df_merge["DATA"] = df_merge["DATA_EMISSAO"].combine_first(df_merge["DATA_CANCELADO"])

    fig5 = px.line(
        df_merge,
        x="DATA",
        y=["CTRC_EMITIDO", "CANCELAMENTOS"],
        title="Linha do Tempo",
    )
    st.plotly_chart(fig5, use_container_width=True)

    # 📊 Índice de Cancelamento vs Meta
    st.subheader("📊 Índice de Cancelamento vs Meta (mensal)")

    emissoes["mes_ano"] = emissoes["DATA_EMISSAO"].dt.to_period("M").astype(str)
    cancelamentos["mes_ano"] = cancelamentos["DATA_CANCELADO"].dt.to_period("M").astype(str)

    emi_mes = emissoes.groupby("mes_ano")["CTRC_EMITIDO"].sum().reset_index(name="emissoes")
    canc_mes = cancelamentos.groupby("mes_ano").size().reset_index(name="cancelamentos")

    indice_mes = pd.merge(emi_mes, canc_mes, on="mes_ano", how="outer").fillna(0)
    indice_mes["indice_cancelamento"] = indice_mes["cancelamentos"] / indice_mes["emissoes"]
    indice_mes["meta"] = 0.0075

    fig_meta = px.line(
        indice_mes,
        x="mes_ano",
        y=["indice_cancelamento", "meta"],
        labels={"value": "Índice de Cancelamento", "mes_ano": "Mês", "variable": "Tipo"},
        markers=True,
        title="Índice de Cancelamento vs Meta (0,75%)"
    )
    fig_meta.update_traces(mode="lines+markers")
    st.plotly_chart(fig_meta, use_container_width=True)
