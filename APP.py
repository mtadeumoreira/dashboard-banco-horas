import pandas as pd
import streamlit as st
import plotly.express as px
import tempfile
import os

from reportlab.platypus import SimpleDocTemplate, Paragraph, Table, TableStyle, Spacer, Image
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4

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
if os.path.exists("logo.png"):
    st.image("logo.png", width=250)

st.title("📊 Dashboard Banco de Horas")
st.markdown("---")

# ===== SIDEBAR =====
st.sidebar.title("⚙️ Configurações")

valor_hora = st.sidebar.number_input("💰 Valor da hora (R$)", value=17.52)

if st.sidebar.button("🔄 Atualizar Dados"):
    st.cache_data.clear()
    st.rerun()

st.caption("📊 Fonte: automática (Google Sheets) ou manual (CSV)")

# ===== FUNÇÕES =====
def normalizar_nome(nome):
    return str(nome).strip().upper()

@st.cache_data(ttl=300)
def carregar_google():
    sheet_id = "1zFKLz8SMEifA8si1f-ORLdGnbrNkOlw3vP7jNzDff9Y"
    url = f"https://docs.google.com/spreadsheets/d/{sheet_id}/export?format=csv"
    return pd.read_csv(url)

@st.cache_data(ttl=300)
def carregar_funcoes():
    sheet_id_funcoes = "1Hr9rxXCQydxocV8P8mbC12u_5V0konGiUxVN9PJ-cYI"
    url = f"https://docs.google.com/spreadsheets/d/{sheet_id_funcoes}/export?format=csv"
    return pd.read_csv(url)

# ===== UPLOAD =====
file = st.file_uploader("📁 Envie o relatório CSV (opcional)", type=["csv"])

if file is not None:
    st.success("📁 Dados carregados via CSV")
    df = pd.read_csv(file, sep=";")
else:
    try:
        df = carregar_google()
        st.info("🌐 Dados carregados automaticamente (Google Sheets)")
    except:
        st.error("❌ Erro ao carregar dados. Envie um CSV manualmente.")
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

# ===== CONVERSÃO =====
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
            return float(valor)

        return -total if negativo else total
    except:
        return None

df["Saldo_horas"] = df["Saldo"].apply(converter_horas)

# ===== VALIDAÇÃO =====
total_registros = len(df)
total_convertidos = df["Saldo_horas"].notna().sum()
total_positivos = (df["Saldo_horas"] > 0).sum()

with st.expander("🔍 Validação dos dados"):
    st.write(f"Total de registros: {total_registros}")
    st.write(f"Convertidos: {total_convertidos}")
    st.write(f"Saldo positivo: {total_positivos}")

df = df.dropna(subset=["Saldo_horas"])

# ===== DATA DE ATUALIZAÇÃO =====
data_importacao = datetime.now().strftime("%d/%m/%Y %H:%M")
st.caption(f"🕒 Última atualização: {data_importacao}")

# ===== KPIs =====
df_pos = df[df["Saldo_horas"] > 0]

total_pos = df_pos["Saldo_horas"].sum()
valor_pagar = total_pos * valor_hora

qtd_func = len(df_pos)
media_horas = df_pos["Saldo_horas"].mean() if not df_pos.empty else 0
media_valor = valor_pagar / qtd_func if qtd_func > 0 else 0

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

# ===== CARDS PRINCIPAIS =====
col1, col2, col3 = st.columns(3)

# CARD 1 - BANCO DE HORAS
col1.markdown(f"""
<div style="background:#0a7d3b;padding:20px;border-radius:10px;color:white;text-align:center">
<h4>⏱️ Banco Positivo</h4>
<h2>{round(total_pos,2)} horas</h2>
</div>
""", unsafe_allow_html=True)

# CARD 2 - VALOR
col2.markdown(f"""
<div style="background:#f2c94c;padding:20px;border-radius:10px;text-align:center">
<h4>💰 Valor a Pagar</h4>
<h2>R$ {valor_pagar:,.2f}</h2>
</div>
""", unsafe_allow_html=True)

# CARD 3 - MÉDIAS
col3.markdown(f"""
<div style="background:#ffffff;padding:20px;border-radius:10px;text-align:center;border:2px solid #0a7d3b">
<h4 style="color:#0a7d3b;">📊 Resumo Médio</h4>
<p>⏱️ Média Horas por Funcionario: <b>{media_horas:.2f} h</b></p>
<p>💰 Média Valor por Funcionario: <b>R$ {media_valor:,.2f}</b></p>
</div>
""", unsafe_allow_html=True)

st.markdown("")

# ===== CARD STATUS (LINHA SEPARADA) =====
st.markdown(f"""
<div style="background:{cor};padding:20px;border-radius:10px;color:white;text-align:center">
<h4>🚨 Situação do Banco de Horas</h4>
<p>👥 {qtd_func} funcionários impactados</p>
<h2>{status}</h2>
</div>
""", unsafe_allow_html=True)

st.markdown("---")

# ===== GRÁFICOS =====
st.subheader("🏆 Top 10 maiores saldos de horas")

ranking = df_pos.sort_values("Saldo_horas", ascending=False).head(10)

fig = px.bar(
    ranking,
    x="Funcionario",
    y="Saldo_horas",
    text=ranking["Saldo_horas"].round(1),
    color_discrete_sequence=["#0a7d3b"]
)
fig.update_traces(textposition='outside')
st.plotly_chart(fig, use_container_width=True)

# ===== FINANCEIRO =====
ranking["Valor_R$"] = ranking["Saldo_horas"] * valor_hora

st.subheader("💰 Custo por Funcionário")

fig2 = px.bar(
    ranking,
    x="Funcionario",
    y="Valor_R$",
    text=ranking["Valor_R$"].apply(lambda x: f"R$ {x:,.0f}"),
    color_discrete_sequence=["#f2c94c"]
)
fig2.update_traces(textposition='outside')
st.plotly_chart(fig2, use_container_width=True)

# ===== LISTA =====
st.subheader("📋 Funcionários com saldo positivo")
df_lista = df_pos.copy()

df_lista["Valor_R$"] = df_lista["Saldo_horas"] * valor_hora

st.dataframe(
    df_lista.sort_values("Saldo_horas", ascending=False)[
        ["Funcionario", "Saldo_horas", "Valor_R$"]
    ],
    use_container_width=True
)

# ===== PDF =====
if st.button("📄 Gerar Relatório PDF"):

    def normalizar_funcao(funcao):
        return str(funcao).replace("\n", " ").strip().upper()

    df_temp = df_pos.copy()
    df_temp["Funcao"] = df_temp.get("Funcao", "OUTROS")
    df_temp["Funcao"] = df_temp["Funcao"].apply(normalizar_funcao)

    def ordenar_funcao(funcao):
        if "MOTORISTA" in funcao:
            return (1, funcao)
        elif "AUXILIAR DE ENTREGA" in funcao:
            return (2, funcao)
        else:
            return (3, funcao)

    df_temp["ordem"] = df_temp["Funcao"].apply(ordenar_funcao)
    df_temp = df_temp.sort_values(["ordem", "Saldo_horas"], ascending=[True, False])

    temp_file = tempfile.NamedTemporaryFile(delete=False, suffix=".pdf")

    doc = SimpleDocTemplate(temp_file.name, pagesize=A4)
    styles = getSampleStyleSheet()
    content = []

    data_atual = datetime.now().strftime("%d/%m/%Y %H:%M")

    content.append(Paragraph("RELATÓRIO EXECUTIVO", styles["Title"]))
    content.append(Spacer(1, 10))
    content.append(Paragraph(f"Data de geração: {data_atual}", styles["Normal"]))
    content.append(Spacer(1, 15))

    resumo = [
        ["Indicador", "Valor"],
        ["Registros", total_registros],
        ["Convertidos", total_convertidos],
        ["Saldo positivo", total_positivos],
        ["Banco de horas", f"{round(total_pos,2)} h"],
        ["Valor a pagar", f"R$ {valor_pagar:,.2f}"],
        ["Média de horas", f"{media_horas:.2f}"],
        ["Status", status]
    ]

    tabela = Table(resumo)
    tabela.setStyle(TableStyle([
        ("BACKGROUND", (0,0), (-1,0), colors.HexColor("#0a7d3b")),
        ("TEXTCOLOR", (0,0), (-1,0), colors.white),
        ("GRID", (0,0), (-1,-1), 0.5, colors.grey),
    ]))

    content.append(tabela)
    content.append(Spacer(1, 20))

    dados_func = [["Função", "Funcionário", "Horas", "Valor"]]

    for _, row in df_temp.iterrows():
        valor = row["Saldo_horas"] * valor_hora
        dados_func.append([
            row["Funcao"],
            row["Funcionario"],
            f"{round(row['Saldo_horas'],2)} h",
            f"R$ {valor:,.2f}"
        ])

    tabela_func = Table(dados_func)
    tabela_func.setStyle(TableStyle([
        ("BACKGROUND", (0,0), (-1,0), colors.HexColor("#0a7d3b")),
        ("TEXTCOLOR", (0,0), (-1,0), colors.white),
        ("GRID", (0,0), (-1,-1), 0.3, colors.grey),
    ]))

    content.append(tabela_func)
    content.append(Spacer(1, 20))

    content.append(Paragraph(
        f"Relatório gerado automaticamente em {data_atual}",
        styles["Italic"]
    ))

    doc.build(content)

    with open(temp_file.name, "rb") as f:
        st.download_button("📥 Baixar PDF", f, file_name="relatorio.pdf")