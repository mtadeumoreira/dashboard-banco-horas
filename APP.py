import pandas as pd
import streamlit as st
import plotly.express as px
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, Image
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib.units import inch
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

# ===== NORMALIZAR =====
df["Funcionario"] = df["Funcionario"].apply(normalizar_nome)

# ===== CARREGAR FUNÇÕES E MERGE =====
try:
    df_funcoes = carregar_funcoes()
    df_funcoes["Funcionario"] = df_funcoes["Funcionario"].apply(normalizar_nome)

    df = df.merge(df_funcoes, on="Funcionario", how="left")
    df["Funcao"] = df["Funcao"].fillna("Não informado")

except:
    df["Funcao"] = "Não informado"

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
            total = float(valor)

        return -total if negativo else total
    except:
        return None

df["Saldo_horas"] = df["Saldo"].apply(converter_horas)

# ===== VALOR =====
df["Valor_R$"] = df["Saldo_horas"] * valor_hora

# ===== VALIDAÇÃO =====
total_registros = len(df)
total_convertidos = df["Saldo_horas"].notna().sum()
total_invalidos = total_registros - total_convertidos

df = df.dropna(subset=["Saldo_horas"])

df_pos = df[df["Saldo_horas"] > 0]

df_pos = df_pos.sort_values(
    by=["Funcao", "Saldo_horas"],
    ascending=[True, False]
)

# ===== KPIs =====
total_pos = df_pos["Saldo_horas"].sum()
valor_pagar = total_pos * valor_hora
qtd_func = len(df_pos)
media_horas = df_pos["Saldo_horas"].mean() if not df_pos.empty else 0
media_valor = valor_pagar / qtd_func if qtd_func > 0 else 0
percentual_positivo = (qtd_func / len(df)) * 100 if len(df) > 0 else 0

# ===== SEMÁFORO =====
if media_horas == 0:
    status = "🟢 Saudável"
elif media_horas <= 10:
    status = "🟡 Atenção"
else:
    status = "🔴 Crítico"

# ===== INFO =====
st.caption(f"🕒 Última atualização: {datetime.now().strftime('%d/%m/%Y %H:%M')}")

# ===== CARDS =====
col1, col2, col3 = st.columns(3)

col1.metric("⏱️ Banco Positivo", f"{round(total_pos,2)} h")
col2.metric("💰 Valor a Pagar", f"R$ {valor_pagar:,.2f}")
col3.metric("📊 Média", f"{media_horas:.1f} h")

st.markdown(f"### 🚨 Situação: {status}")

# ===== PDF =====
def gerar_pdf():
    temp_file = tempfile.NamedTemporaryFile(delete=False)

    doc = SimpleDocTemplate(temp_file.name)
    styles = getSampleStyleSheet()

    conteudo = []

    # ===== LOGO =====
    try:
        logo = Image("logo.png", width=2*inch, height=1*inch)
        logo.hAlign = 'CENTER'
        conteudo.append(logo)
    except:
        pass

    # ===== TÍTULO =====
    conteudo.append(Paragraph("<b>RELATÓRIO EXECUTIVO</b>", styles['Title']))
    conteudo.append(Paragraph("Banco de Horas", styles['Heading2']))
    conteudo.append(Spacer(1, 15))

    # ===== LINHA VERDE =====
    linha = Table([[""]], colWidths=[500])
    linha.setStyle(TableStyle([
        ('LINEABOVE', (0,0), (-1,-1), 2, colors.HexColor("#0a7d3b"))
    ]))
    conteudo.append(linha)
    conteudo.append(Spacer(1, 20))

    # ===== RESUMO EXECUTIVO (PADRÃO MANTIDO) =====
    conteudo.append(Paragraph("<b>Resumo Executivo</b>", styles['Heading3']))
    conteudo.append(Spacer(1, 10))

    dados_resumo = [
        ["Indicador", "Valor"],
        ["Total de registros", total_registros],
        ["Convertidos", total_convertidos],
        ["Saldo positivo", qtd_func],
        ["Banco de horas (h)", round(total_pos,2)],
        ["Valor a pagar", f"R$ {valor_pagar:,.2f}"],
        ["Funcionários impactados", qtd_func],
        ["Média por funcionário", f"R$ {media_valor:,.2f}"],
        ["Média de horas", f"{media_horas:.2f}"],
        ["% com saldo positivo", f"{percentual_positivo:.1f}%"],
        ["Status", status]
    ]

    tabela_resumo = Table(dados_resumo, colWidths=[250, 250])

    tabela_resumo.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,0), colors.HexColor("#0a7d3b")),
        ('TEXTCOLOR', (0,0), (-1,0), colors.white),
        ('GRID', (0,0), (-1,-1), 0.5, colors.grey),
        ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'),
    ]))

    conteudo.append(tabela_resumo)
    conteudo.append(Spacer(1, 25))

    # ===== FUNCIONÁRIOS AGRUPADOS POR FUNÇÃO =====
conteudo.append(Paragraph("<b>Funcionários com saldo positivo</b>", styles['Heading3']))
conteudo.append(Spacer(1, 10))

from collections import defaultdict

grupos = defaultdict(list)

for _, row in df_pos.iterrows():
    grupos[row["Funcao"]].append(row)

total_geral_horas = 0
total_geral_valor = 0

for funcao, lista in grupos.items():

    # ===== TÍTULO DA FUNÇÃO =====
    conteudo.append(Paragraph(f"<b>Função: {funcao}</b>", styles['Normal']))
    conteudo.append(Spacer(1, 6))

    dados_func = [["Funcionário", "Horas", "Valor (R$)"]]

    subtotal_horas = 0
    subtotal_valor = 0

    for row in lista:
        horas = round(row["Saldo_horas"], 2)
        valor = row["Valor_R$"]

        subtotal_horas += horas
        subtotal_valor += valor

        dados_func.append([
            Paragraph(str(row["Funcionario"]), styles['Normal']),
            f"{horas} h",
            f"R$ {valor:,.2f}"
        ])

    # ===== TABELA =====
    tabela = Table(dados_func, colWidths=[250, 80, 100])

    tabela.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor("#0a7d3b")),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
        ('GRID', (0, 0), (-1, -1), 0.3, colors.grey),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
    ]))

    conteudo.append(tabela)

    # ===== SUBTOTAL =====
    conteudo.append(Spacer(1, 6))
    conteudo.append(
        Paragraph(
            f"<b>Subtotal:</b> {round(subtotal_horas,2)} h | R$ {subtotal_valor:,.2f}",
            styles['Normal']
        )
    )

    conteudo.append(Spacer(1, 15))

    total_geral_horas += subtotal_horas
    total_geral_valor += subtotal_valor

# ===== TOTAL GERAL =====
conteudo.append(Spacer(1, 10))

conteudo.append(Paragraph("<b>TOTAL GERAL</b>", styles['Heading3']))
conteudo.append(Spacer(1, 5))

conteudo.append(
    Paragraph(
        f"<b>Horas:</b> {round(total_geral_horas,2)} h",
        styles['Normal']
    )
)

conteudo.append(
    Paragraph(
        f"<b>Valor:</b> R$ {total_geral_valor:,.2f}",
        styles['Normal']
    )
)

    # ===== RODAPÉ =====
    conteudo.append(Paragraph(
        "Relatório gerado automaticamente pelo sistema de gestão de banco de horas.",
        styles['Italic']
    ))

    doc.build(conteudo)

    return temp_file.name

if st.button("📄 Exportar PDF Executivo"):
    pdf_path = gerar_pdf()
    with open(pdf_path, "rb") as f:
        st.download_button("⬇️ Baixar PDF", f, file_name="relatorio_banco_horas.pdf")