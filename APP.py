import pandas as pd
import streamlit as st
import plotly.express as px
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer
from reportlab.lib.styles import getSampleStyleSheet
import tempfile
from datetime import datetime

st.set_page_config(layout="wide")

# ===== ESTILO =====
st.markdown("""
<style>
.stApp {background-color: #f5f5f5;}
h1, h3 {color: #0a7d3b;}
section[data-testid="stSidebar"] {background-color: #0a7d3b; color: white;}
</style>
""", unsafe_allow_html=True)

# ===== LOGO =====
st.image("logo.png", width=250)

st.title("📊 Dashboard Banco de Horas")
st.markdown("---")

# ===== SIDEBAR =====
st.sidebar.title("⚙️ Configurações")

valor_hora = st.sidebar.number_input("💰 Valor da hora (R$)", value=17.52)

if st.sidebar.button("🔄 Atualizar Dados"):
    st.cache_data.clear()
    st.rerun()

# ===== GOOGLE SHEETS =====
@st.cache_data(ttl=300)
def carregar_google():
    sheet_id = "1zFKLz8SMEifA8si1f-ORLdGnbrNkOlw3vP7jNzDff9Y"
    url = f"https://docs.google.com/spreadsheets/d/{sheet_id}/export?format=csv"
    return pd.read_csv(url)

# ===== UPLOAD =====
file = st.file_uploader("📁 Envie o relatório CSV (opcional)", type=["csv"])

# ===== FONTE DE DADOS =====
if file is not None:
    st.success("📁 Dados carregados via CSV")
    df = pd.read_csv(file, sep=";")
else:
    try:
        df = carregar_google()
        st.info("🌐 Dados carregados automaticamente (Google Sheets)")
    except Exception as e:
        st.error(f"❌ Erro ao carregar dados: {e}")
        st.stop()

# ===== PADRONIZAÇÃO =====
df = df.rename(columns={
    df.columns[0]: "Funcionario",
    df.columns[1]: "Horas",
    df.columns[-1]: "Saldo"
})

# ===== LIMPEZA =====
df["Funcionario"] = df["Funcionario"].astype(str).str.strip()
df = df[(df["Funcionario"] != "") & (~df["Funcionario"].str.lower().isin(["null","none","nan"]))]

# ===== CONVERSÃO ROBUSTA =====
def converter_horas(valor):
    try:
        valor = str(valor).strip()

        if valor in ["", "nan", "None", "null"]:
            return None

        negativo = "-" in valor
        valor = valor.replace("-", "")

        partes = valor.split(":")

        if len(partes) == 2:
            h, m = partes
            total = float(h) + float(m)/60

        elif len(partes) == 3:
            h, m, s = partes
            total = float(h) + float(m)/60 + float(s)/3600

        else:
            total = float(valor)

        return -total if negativo else total

    except:
        return None

df["Saldo_horas"] = df["Saldo"].apply(converter_horas)

# ===== VALIDAÇÃO =====
total_registros = len(df)
total_convertidos = df["Saldo_horas"].notna().sum()
total_invalidos = total_registros - total_convertidos

df = df.dropna(subset=["Saldo_horas"])

df_pos = df[df["Saldo_horas"] > 0]

# ===== KPIs =====
total_pos = df_pos["Saldo_horas"].sum()
valor_pagar = total_pos * valor_hora

qtd_func = len(df_pos)
media_horas = df_pos["Saldo_horas"].mean() if not df_pos.empty else 0
media_valor = valor_pagar / qtd_func if qtd_func > 0 else 0

percentual_positivo = (qtd_func / len(df)) * 100 if len(df) > 0 else 0
top_funcionario = df_pos.sort_values("Saldo_horas", ascending=False).head(1)

# ===== SEMÁFORO =====
if media_horas == 0:
    cor = "green"
    status = "🟢 Saudável"
elif media_horas <= 10:
    cor = "orange"
    status = "🟡 Atenção"
else:
    cor = "red"
    status = "🔴 Crítico"

# ===== INFO GERAL =====
st.caption(f"🕒 Última atualização: {datetime.now().strftime('%d/%m/%Y %H:%M')}")

with st.expander("🔍 Validação dos dados"):
    st.write(f"Total registros: {total_registros}")
    st.write(f"Convertidos: {total_convertidos}")
    st.write(f"Inválidos: {total_invalidos}")

# ===== CARDS =====
col1, col2, col3, col4 = st.columns(4)

col1.metric("⏱️ Banco Positivo", f"{round(total_pos,2)} h")
col2.metric("💰 Valor a Pagar", f"R$ {valor_pagar:,.2f}")
col3.metric("👥 % com saldo", f"{percentual_positivo:.1f}%")
col4.metric("📊 Média por funcionário", f"{media_horas:.1f} h")

st.markdown(f"### 🚨 Situação: {status}")

# ===== GRÁFICOS =====
ranking = df_pos.sort_values("Saldo_horas", ascending=False).head(10)

fig = px.bar(
    ranking,
    x="Funcionario",
    y="Saldo_horas",
    text=ranking["Saldo_horas"].round(1)
)
st.plotly_chart(fig, use_container_width=True)

ranking["Valor_R$"] = ranking["Saldo_horas"] * valor_hora

fig2 = px.bar(
    ranking,
    x="Funcionario",
    y="Valor_R$",
    text=ranking["Valor_R$"].apply(lambda x: f"R$ {x:,.0f}")
)
st.plotly_chart(fig2, use_container_width=True)

# ===== LISTA =====
st.subheader("📋 Funcionários com saldo positivo")
st.dataframe(df_pos.sort_values("Saldo_horas", ascending=False), use_container_width=True)

# ===== PDF =====
def gerar_pdf():
    temp_file = tempfile.NamedTemporaryFile(delete=False)
    doc = SimpleDocTemplate(temp_file.name)
    styles = getSampleStyleSheet()

    conteudo = []

    conteudo.append(Paragraph("Relatório Executivo - Banco de Horas", styles['Title']))
    conteudo.append(Spacer(1, 12))

    conteudo.append(Paragraph(f"Total de horas positivas: {round(total_pos,2)}", styles['Normal']))
    conteudo.append(Paragraph(f"Valor estimado: R$ {valor_pagar:,.2f}", styles['Normal']))
    conteudo.append(Paragraph(f"Média por funcionário: {media_horas:.2f}", styles['Normal']))
    conteudo.append(Paragraph(f"Situação: {status}", styles['Normal']))

    doc.build(conteudo)

    return temp_file.name

if st.button("📄 Exportar PDF Executivo"):
    pdf_path = gerar_pdf()
    with open(pdf_path, "rb") as f:
        st.download_button("⬇️ Baixar PDF", f, file_name="relatorio_banco_horas.pdf")