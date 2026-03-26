import pandas as pd
import streamlit as st
import plotly.express as px
from reportlab.platypus import SimpleDocTemplate, Paragraph
from reportlab.lib.styles import getSampleStyleSheet
import tempfile

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
    st.rerun()

# ===== FUNÇÃO GOOGLE SHEETS =====
@st.cache_data(ttl=60)
def carregar_google():
    sheet_id = "1zFKLz8SMEifA8si1f-ORLdGnbrNkOlw3vP7jNzDff9Y"
    url = f"https://docs.google.com/spreadsheets/d/{sheet_id}/export?format=csv"
    return pd.read_csv(url)

# ===== UPLOAD =====
file = st.file_uploader("📁 Envie o relatório CSV (opcional)", type=["csv"])

# ===== ESCOLHA AUTOMÁTICA =====
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

# ===== PADRONIZAÇÃO DE COLUNAS =====
df = df.rename(columns={
    df.columns[0]: "Funcionario",
    df.columns[1]: "Horas",
    df.columns[-1]: "Saldo"
})
    # ===== LIMPEZA =====
    df["Funcionario"] = df["Funcionario"].astype(str).str.strip()
    df = df[(df["Funcionario"] != "") & (~df["Funcionario"].str.lower().isin(["null","none","nan"]))]

    # ===== CONVERSÃO ROBUSTA (HH:MM e HH:MM:SS) =====
    def converter_horas(valor):
        try:
            valor = str(valor).strip()

            if valor in ["", "nan", "None", "null"]:
                return None

            negativo = "-" in valor
            valor = valor.replace("-", "")

            partes = valor.split(":")

            if len(partes) == 2:  # HH:MM
                h, m = partes
                total = float(h) + float(m)/60

            elif len(partes) == 3:  # HH:MM:SS
                h, m, s = partes
                total = float(h) + float(m)/60 + float(s)/3600

            else:
                return float(valor)

            return -total if negativo else total

        except:
            return None

    df["Saldo_horas"] = df["Saldo"].apply(converter_horas)

    # ===== VALIDAÇÃO COMPLETA =====
    total_registros = len(df)
    total_convertidos = df["Saldo_horas"].notna().sum()
    total_positivos = (df["Saldo_horas"] > 0).sum()

    with st.expander("🔍 Validação dos dados"):
        st.write(f"Total de registros no arquivo: {total_registros}")
        st.write(f"Convertidos com sucesso: {total_convertidos}")
        st.write(f"Saldo positivo (>0): {total_positivos}")

    # ===== REMOVER APENAS OS INVÁLIDOS =====
    df = df.dropna(subset=["Saldo_horas"])

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

    # ===== CARDS =====
    col1, col2, col3 = st.columns(3)

    col1.markdown(f"""
    <div style="background:#0a7d3b;padding:20px;border-radius:10px;color:white;text-align:center">
    <h4>⏱️ Banco Positivo</h4>
    <h2>{round(total_pos,2)} horas</h2>
    </div>
    """, unsafe_allow_html=True)

    col2.markdown(f"""
    <div style="background:#f2c94c;padding:20px;border-radius:10px;text-align:center">
    <h4>💰 Valor a Pagar</h4>
    <h2>R$ {valor_pagar:,.2f}</h2>
    </div>
    """, unsafe_allow_html=True)

    col3.markdown(f"""
    <div style="background:{cor};padding:20px;border-radius:10px;color:white;text-align:center">
    <h4>🚨 Situação do Banco</h4>
    <p>👥 {qtd_func} funcionários</p>
    <p>💰 Média: R$ {media_valor:,.2f}</p>
    <h3>{status}</h3>
    </div>
    """, unsafe_allow_html=True)

    st.markdown("---")

    # ===== GRÁFICO HORAS =====
    st.subheader("🏆 Top maiores saldos de horas")

    ranking = df_pos.sort_values("Saldo_horas", ascending=False).head(10)

    fig = px.bar(
        ranking,
        x="Funcionario",
        y="Saldo_horas",
        text=ranking["Saldo_horas"].round(1),
        color_discrete_sequence=["#0a7d3b"]
    )

    fig.update_traces(textposition='outside')

    fig.update_layout(
        height=600,
        margin=dict(t=80),
        xaxis_tickangle=-45
    )

    st.plotly_chart(fig, use_container_width=True)

    # ===== GRÁFICO FINANCEIRO =====
    ranking["Valor_R$"] = ranking["Saldo_horas"] * valor_hora
    ranking["Valor_formatado"] = ranking["Valor_R$"].apply(lambda x: f"R$ {x:,.2f}")

    st.subheader("💰 Custo por Funcionário")

    fig2 = px.bar(
        ranking,
        x="Funcionario",
        y="Valor_R$",
        text="Valor_formatado",
        color_discrete_sequence=["#f2c94c"]
    )

    fig2.update_traces(textposition='outside')

    fig2.update_layout(
        height=600,
        margin=dict(t=80),
        xaxis_tickangle=-45
    )

    st.plotly_chart(fig2, use_container_width=True)

    st.markdown("---")

    # ===== LISTA =====
    st.subheader("📋 Funcionários com saldo positivo")

    df_lista = df_pos.sort_values("Saldo_horas", ascending=False)

    st.dataframe(df_lista[["Funcionario","Saldo_horas"]], use_container_width=True)

    # ===== PDF =====
# ===== PDF PREMIUM =====
if st.button("📄 Gerar Relatório PDF"):

    from reportlab.platypus import Table, TableStyle, Spacer, Image
    from reportlab.lib import colors
    from reportlab.lib.pagesizes import A4
    from reportlab.lib.units import cm

    temp_file = tempfile.NamedTemporaryFile(delete=False, suffix=".pdf")

    doc = SimpleDocTemplate(
        temp_file.name,
        pagesize=A4,
        rightMargin=2*cm,
        leftMargin=2*cm,
        topMargin=2*cm,
        bottomMargin=2*cm
    )

    styles = getSampleStyleSheet()
    content = []

    # ===== LOGO =====
    try:
        logo = Image("logo.png", width=120, height=50)
        content.append(logo)
    except:
        pass

    content.append(Spacer(1, 10))

    # ===== TÍTULO =====
    content.append(Paragraph("RELATÓRIO EXECUTIVO", styles["Title"]))
    content.append(Paragraph("Banco de Horas", styles["Heading2"]))
    content.append(Spacer(1, 15))

    # ===== LINHA SEPARADORA =====
    linha = Table([[""]], colWidths=[500])
    linha.setStyle(TableStyle([
        ("LINEBELOW", (0,0), (-1,-1), 2, colors.HexColor("#0a7d3b"))
    ]))
    content.append(linha)
    content.append(Spacer(1, 15))

    # ===== RESUMO EM FORMATO EXECUTIVO =====
    resumo_data = [
        ["Indicador", "Valor"],
        ["Total de registros", total_registros],
        ["Convertidos", total_convertidos],
        ["Saldo positivo", total_positivos],
        ["Banco de horas (h)", round(total_pos,2)],
        ["Valor a pagar", f"R$ {valor_pagar:,.2f}"],
        ["Funcionários impactados", qtd_func],
        ["Média por funcionário", f"R$ {media_valor:,.2f}"],
        ["Status", status]
    ]

    tabela_resumo = Table(resumo_data, colWidths=[250, 200])

    tabela_resumo.setStyle(TableStyle([
        ("BACKGROUND", (0,0), (-1,0), colors.HexColor("#0a7d3b")),
        ("TEXTCOLOR", (0,0), (-1,0), colors.white),
        ("FONTNAME", (0,0), (-1,0), "Helvetica-Bold"),

        ("GRID", (0,0), (-1,-1), 0.5, colors.grey),
        ("BACKGROUND", (0,1), (-1,-1), colors.whitesmoke),
    ]))

    content.append(Paragraph("Resumo Executivo", styles["Heading2"]))
    content.append(Spacer(1, 10))
    content.append(tabela_resumo)

    content.append(Spacer(1, 20))

    # ===== TABELA PRINCIPAL =====
    content.append(Paragraph("Funcionários com saldo positivo", styles["Heading2"]))
    content.append(Spacer(1, 10))

    df_lista = df_pos.sort_values("Saldo_horas", ascending=False)

    dados = [["Funcionário", "Horas"]]

    for i, row in df_lista.iterrows():
        nome = str(row["Funcionario"])
        horas = round(row["Saldo_horas"], 2)

        dados.append([nome, f"{horas} h"])

    tabela = Table(dados, colWidths=[300, 100])

    tabela.setStyle(TableStyle([
        ("BACKGROUND", (0,0), (-1,0), colors.HexColor("#0a7d3b")),
        ("TEXTCOLOR", (0,0), (-1,0), colors.white),
        ("FONTNAME", (0,0), (-1,0), "Helvetica-Bold"),

        ("ALIGN", (1,1), (-1,-1), "CENTER"),

        ("GRID", (0,0), (-1,-1), 0.5, colors.grey),

        ("BACKGROUND", (0,1), (-1,-1), colors.whitesmoke),
    ]))

    content.append(tabela)

    content.append(Spacer(1, 30))

    # ===== RODAPÉ =====
    content.append(Paragraph(
        "Relatório gerado automaticamente pelo sistema de gestão de banco de horas.",
        styles["Italic"]
    ))

    # ===== GERAR PDF =====
    doc.build(content)

    # ===== DOWNLOAD =====
    with open(temp_file.name, "rb") as f:
        st.download_button(
            "📥 Baixar Relatório Executivo",
            f,
            file_name="relatorio_executivo_banco_horas.pdf"
        )