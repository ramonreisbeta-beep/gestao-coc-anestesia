import streamlit as st
import pandas as pd
from datetime import datetime, date
import gspread
from google.oauth2.service_account import Credentials
from io import BytesIO

# PDF
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib import colors
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.units import inch
from reportlab.lib.pagesizes import A4

# Excel
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils.dataframe import dataframe_to_rows

st.set_page_config(page_title="Gestão COC Anestesia", layout="wide")
st.title("Registro de Cirurgias e Dashboard Financeiro")

# ==========================================================
# CONEXÃO GOOGLE SHEETS
# ==========================================================
@st.cache_resource
def conectar_sheets():
    scopes = ["https://www.googleapis.com/auth/spreadsheets",
              "https://www.googleapis.com/auth/drive"]
    credenciais_dict = dict(st.secrets["connections"]["gsheets"])
    if "spreadsheet" in credenciais_dict:
        del credenciais_dict["spreadsheet"]
    credentials = Credentials.from_service_account_info(
        credenciais_dict, scopes=scopes)
    return gspread.authorize(credentials)

try:
    client = conectar_sheets()
    URL_PLANILHA = st.secrets["connections"]["gsheets"]["spreadsheet"]
    planilha = client.open_by_url(URL_PLANILHA)

    aba_cirurgias = planilha.worksheet("CIRURGIAS")
    aba_convenios = planilha.worksheet("Página2")
    aba_cbhpm = planilha.worksheet("Página3")

    df_cirurgias = pd.DataFrame(aba_cirurgias.get_all_records())
    df_convenios = pd.DataFrame(aba_convenios.get_all_records())
    df_cbhpm = pd.DataFrame(aba_cbhpm.get_all_records())

except Exception as e:
    st.error(f"Erro de conexão com as planilhas: {e}")
    st.stop()

# ==========================================================
# FUNÇÕES AUXILIARES
# ==========================================================
def limpar_moeda(valor):
    if pd.isna(valor) or str(valor).strip() in ['-', '', '0']:
        return 0.0
    valor_str = str(valor).replace('R$', '').replace('.', '').replace(',', '.').strip()
    try:
        return float(valor_str)
    except:
        return 0.0

def formatar_real(valor):
    return f"R$ {valor:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.')

def converter_para_horas(duracao_str):
    try:
        texto = str(duracao_str).strip()
        if not texto or texto == "nan":
            return None
        partes = texto.split(':')
        return int(partes[0]) + (int(partes[1]) / 60.0)
    except:
        return None

# Mapas para performance
mapa_cbhpm = df_cbhpm.set_index('Código').to_dict('index')
mapa_convenios = df_convenios.set_index('Convênio').to_dict('index')

def calcular_faturamento_memoria(row):
    convenio = str(row.get('CONVÊNIO', '')).strip()
    procs_str = str(row.get('PROCEDIMENTO', '')).strip()

    if not convenio or not procs_str or procs_str == 'nan':
        return 0.0

    linha_convenio = mapa_convenios.get(convenio)
    if not linha_convenio:
        return 0.0

    valor_total = 0.0
    lista_procs = procs_str.split('\n')

    for i, proc in enumerate(lista_procs):
        codigo = proc.split(" - ")[0].strip()
        linha_cbhpm = mapa_cbhpm.get(codigo)
        if not linha_cbhpm:
            continue

        porte = str(linha_cbhpm.get('Porte Anest.', '')).strip()
        preco = 0.0

        if porte.isdigit():
            col_an = f"AN{porte}"
            if col_an in linha_convenio:
                preco = limpar_moeda(linha_convenio[col_an])

        if i == 0:
            valor_total += preco
        else:
            valor_total += preco * 0.5

    return valor_total

# ==========================================================
# ABAS
# ==========================================================
tab_registro, tab_dashboard = st.tabs(["📝 Novo Registro", "📊 Dashboard Financeiro"])

# ==========================================================
# DASHBOARD
# ==========================================================
with tab_dashboard:

    st.subheader("🏆 Ranking de Rentabilidade")

    if df_cirurgias.empty:
        st.info("Nenhuma cirurgia registrada.")
    else:
        df_cirurgias['Valor Virtual'] = df_cirurgias.apply(calcular_faturamento_memoria, axis=1)

        if 'DURAÇÃO' not in df_cirurgias.columns:
            st.error("Coluna DURAÇÃO não encontrada.")
            st.stop()

        df_cirurgias['Horas'] = df_cirurgias['DURAÇÃO'].apply(converter_para_horas)

        df_cirurgias['R$/Hora'] = df_cirurgias.apply(
            lambda row: row['Valor Virtual'] / row['Horas']
            if row['Horas'] and row['Horas'] > 0 else None,
            axis=1
        )

        faturamento_total = df_cirurgias['Valor Virtual'].sum()
        total_cirurgias = len(df_cirurgias)
        ticket_medio = faturamento_total / total_cirurgias if total_cirurgias else 0

        col1, col2, col3 = st.columns(3)
        col1.metric("💰 Faturamento Total", formatar_real(faturamento_total))
        col2.metric("🏥 Nº Cirurgias", total_cirurgias)
        col3.metric("📊 Ticket Médio", formatar_real(ticket_medio))

        df_validos = df_cirurgias.dropna(subset=['R$/Hora']).copy()

        if not df_validos.empty:

            df_ranking = df_validos.sort_values(by='R$/Hora', ascending=False)

            st.dataframe(df_ranking, use_container_width=True)

            # ======================================================
            # EXPORTAR PDF
            # ======================================================
            if st.button("📄 Exportar PDF"):
                buffer = BytesIO()
                doc = SimpleDocTemplate(buffer, pagesize=A4)
                elementos = []

                estilo = ParagraphStyle(name="Normal", fontSize=10)

                elementos.append(Paragraph("Relatório Financeiro - COC Anestesia", estilo))
                elementos.append(Spacer(1, 0.3 * inch))
                elementos.append(Paragraph(f"Faturamento Total: {formatar_real(faturamento_total)}", estilo))
                elementos.append(Paragraph(f"Nº Cirurgias: {total_cirurgias}", estilo))
                elementos.append(Paragraph(f"Ticket Médio: {formatar_real(ticket_medio)}", estilo))
                elementos.append(Spacer(1, 0.3 * inch))

                dados = [df_ranking.columns.tolist()] + df_ranking.values.tolist()
                tabela = Table(dados, repeatRows=1)
                tabela.setStyle(TableStyle([
                    ('BACKGROUND', (0,0), (-1,0), colors.lightgrey),
                    ('GRID', (0,0), (-1,-1), 0.5, colors.grey),
                    ('FONTSIZE', (0,0), (-1,-1), 7),
                ]))

                elementos.append(tabela)
                doc.build(elementos)
                buffer.seek(0)

                st.download_button(
                    "⬇️ Baixar PDF",
                    data=buffer,
                    file_name="Relatorio_COC_Anestesia.pdf",
                    mime="application/pdf"
                )

            # ======================================================
            # EXPORTAR EXCEL
            # ======================================================
            if st.button("📊 Exportar Excel"):
                wb = Workbook()

                ws1 = wb.active
                ws1.title = "Resumo"
                ws1["A1"] = "Relatório Financeiro - COC Anestesia"
                ws1["A3"] = "Faturamento Total:"
                ws1["B3"] = faturamento_total
                ws1["A4"] = "Nº Cirurgias:"
                ws1["B4"] = total_cirurgias
                ws1["A5"] = "Ticket Médio:"
                ws1["B5"] = ticket_medio

                ws2 = wb.create_sheet("Ranking")

                for r in dataframe_to_rows(df_ranking, index=False, header=True):
                    ws2.append(r)

                for cell in ws2[1]:
                    cell.font = Font(bold=True)

                buffer = BytesIO()
                wb.save(buffer)
                buffer.seek(0)

                st.download_button(
                    "⬇️ Baixar Excel",
                    data=buffer,
                    file_name="Relatorio_COC_Anestesia.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )