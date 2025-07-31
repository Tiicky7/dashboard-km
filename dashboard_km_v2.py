
import streamlit as st
import pandas as pd
import plotly.express as px

# ======= CONFIGURAÇÃO DA PÁGINA ========
st.set_page_config(page_title="Dashboard Cancelamentos KM", layout="wide")

# ======= FUNÇÕES ========
@st.cache_data
def carregar_dados():
    emissoes_df = pd.read_excel("EMISSOES_KM.xlsx")
    cancelamentos_df = pd.read_excel("CANCELAMENTOS_KM.xlsx")
    return emissoes_df, cancelamentos_df

def calcular_metricas(emissoes_df, cancelamentos_df):
    total_emissoes = emissoes_df.shape[0]
    total_cancelamentos = cancelamentos_df.shape[0]
    taxa_cancelamento = total_cancelamentos / total_emissoes if total_emissoes else 0
    return total_emissoes, total_cancelamentos, taxa_cancelamento

# ======= CARREGAMENTO DE DADOS ========
emissoes_df, cancelamentos_df = carregar_dados()
total_emissoes, total_cancelamentos, taxa_cancelamento = calcular_metricas(emissoes_df, cancelamentos_df)

# ======= HEADER ========
st.title("📊 Dashboard de Cancelamentos KM")

# ======= KPIs METRICAS ========
col1, col2, col3 = st.columns(3)
col1.metric("Total de Emissões", f"{total_emissoes}")
col2.metric("Total de Cancelamentos", f"{total_cancelamentos}")
col3.metric("Taxa de Cancelamento", f"{taxa_cancelamento:.2%}", delta="-0.25%" if taxa_cancelamento < 0.0075 else "+0.25%")

# ======= ABAS PRINCIPAIS ========
aba1, aba2, aba3, aba4 = st.tabs(["📈 Evolução", "📉 Cancelamentos", "📋 Comparativo com Meta", "⬇️ Exportar Dados"])

# ======= ABA 1: EVOLUÇÃO ========
with aba1:
    st.subheader("📈 Evolução de Emissões e Cancelamentos")
    emissoes_df['Data'] = pd.to_datetime(emissoes_df['Data'])
    cancelamentos_df['Data'] = pd.to_datetime(cancelamentos_df['Data'])

    emissoes_mensal = emissoes_df.groupby(emissoes_df['Data'].dt.to_period("M")).size().reset_index(name='Emissões')
    cancelamentos_mensal = cancelamentos_df.groupby(cancelamentos_df['Data'].dt.to_period("M")).size().reset_index(name='Cancelamentos')

    df_merge = pd.merge(emissoes_mensal, cancelamentos_mensal, on='Data', how='outer').fillna(0)
    df_merge['Data'] = df_merge['Data'].astype(str)
    df_merge['Taxa Cancelamento'] = df_merge['Cancelamentos'] / df_merge['Emissões']

    fig = px.line(df_merge, x="Data", y=["Emissões", "Cancelamentos"], markers=True)
    st.plotly_chart(fig, use_container_width=True)

# ======= ABA 2: CANCELAMENTOS ========
with aba2:
    st.subheader("📉 Distribuição de Cancelamentos por Motivo")
    if 'Motivo' in cancelamentos_df.columns:
        motivo_df = cancelamentos_df['Motivo'].value_counts().reset_index()
        motivo_df.columns = ['Motivo', 'Quantidade']
        fig_motivo = px.bar(motivo_df, x='Motivo', y='Quantidade', color='Quantidade', text='Quantidade')
        st.plotly_chart(fig_motivo, use_container_width=True)
    else:
        st.warning("Coluna 'Motivo' não encontrada no arquivo de cancelamentos.")

# ======= ABA 3: COMPARATIVO COM META ========
with aba3:
    st.subheader("📋 Comparativo com Meta de Cancelamento (0.75%)")
    fig_meta = px.line(df_merge, x="Data", y="Taxa Cancelamento", title="Taxa de Cancelamento x Meta")
    fig_meta.add_hline(y=0.0075, line_dash="dash", line_color="red", annotation_text="Meta (0.75%)", annotation_position="top left")
    st.plotly_chart(fig_meta, use_container_width=True)

# ======= ABA 4: EXPORTAÇÃO ========
with aba4:
    st.subheader("⬇️ Baixar Dados")
    st.download_button("Download Emissões (.xlsx)", data=emissoes_df.to_excel(index=False), file_name="emissoes.xlsx")
    st.download_button("Download Cancelamentos (.xlsx)", data=cancelamentos_df.to_excel(index=False), file_name="cancelamentos.xlsx")

