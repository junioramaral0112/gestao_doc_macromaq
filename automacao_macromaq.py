import streamlit as st
import pandas as pd
from datetime import datetime, timedelta
from docx import Document
from docx.shared import Inches, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from pptx import Presentation
import openpyxl
import io
import os
import zipfile
import unicodedata
import base64
from urllib.parse import quote

# --- CONFIGURAÇÕES E CAMINHOS ORIGINAIS ---
st.set_page_config(page_title="Automação SSMA Macromaq", layout="wide")

if os.path.exists(r"C:\Users\dilceu.gomes\Documents"):
    BASE_PATH = r"C:\Users\dilceu.gomes\Documents"
else:
    BASE_PATH = os.path.dirname(__file__) if "__file__" in locals() else "."

# Caminhos dos seus arquivos existentes
FUNDO_PATH = os.path.join(BASE_PATH, "fundo.png")
LOGO_PATH = os.path.join(BASE_PATH, "logo.png")
LOGO_CLINICA = os.path.join(BASE_PATH, "adivitta.png") # Adicionado para o ASO
TEMPLATE_FICHA = os.path.join(BASE_PATH, "template_ficha.xlsx")
TEMPLATE_OS_JUNIOR = os.path.join(BASE_PATH, "template_os_Junior.docx")
TEMPLATE_NR06_JUNIOR = os.path.join(BASE_PATH, "template_nr06_Junior.pptx")
TEMPLATE_OS_SIMONE = os.path.join(BASE_PATH, "template_os_simone.docx")
TEMPLATE_NR06_SIMONE = os.path.join(BASE_PATH, "template_nr06_simone.pptx")

# Planilhas
SHEET_ID_DOCS = "1y98U3eK7JXJqQaMC0i7eFbwpvp97Nuyeml5Dis0UCUg"
SHEET_ID_ASO = "1G_oVT9gK-n_jGh5R4g65qUwK_MfQGvCX-SA4NHNNflU"

# --- SUAS FUNÇÕES ORIGINAIS (MANTIDAS) ---
def get_base64(bin_file):
    if not os.path.exists(bin_file): return ""
    with open(bin_file, 'rb') as f: data = f.read()
    return base64.b64encode(data).decode()

def aplicar_layout():
    try:
        fundo = get_base64(FUNDO_PATH)
        logo = get_base64(LOGO_PATH)
        st.markdown(f"""
        <style>
        .stApp {{ background-image: url("data:image/png;base64,{fundo}"); background-size: cover; background-position: center; background-attachment: fixed; }}
        .stSelectbox label, div[data-testid="stCheckbox"] label p {{ color: white !important; background: rgba(0,0,0,0.7); padding: 5px 12px; border-radius: 8px; font-weight: bold; }}
        .stButton > button {{ background: #2c3e50; color: #f9cc0b; border: 2px solid #f9cc0b; border-radius: 10px; height: 55px; font-weight: bold; width: 100%; font-size: 18px; }}
        .header-container {{ display: flex; align-items: center; background: rgba(255,255,255,0.9); padding: 20px; border-radius: 15px; margin-bottom: 30px; }}
        .footer {{ position: fixed; left: 0; bottom: 0; width: 100%; background: rgba(0,0,0,0.8); color: white; text-align: center; padding: 10px; font-size: 13px; z-index: 999; }}
        </style>
        <div class="header-container">
            <img src="data:image/png;base64,{logo}" width="320">
            <h1 style="margin-left:25px;color:#2c3e50;">Gestão SSMA Documentos</h1>
        </div>
        """, unsafe_allow_html=True)
    except: st.title("Gestão SSMA Macromaq")

# ... (Mantenha aqui todas as outras funções: remover_acentos, carregar_aba, preencher_excel_ficha, etc.) ...

# --- NOVA LOGICA DE MENU ---
aplicar_layout()

# Menu lateral para trocar de sistema
st.sidebar.markdown("### 🗂️ Menu Principal")
opcao_menu = st.sidebar.radio("Escolha o Módulo:", ["Emissão de Documentos", "Controle de ASO"])

if opcao_menu == "Emissão de Documentos":
    # --- COLOQUE AQUI O RESTO DO SEU CODIGO ORIGINAL (Colaboradores, Cargos, Checkbox, etc.) ---
    st.markdown("### Gerador de Fichas, OS e Certificados")
    # (Inserir aqui a lógica de df_colab e df_cargos que você já tem)

else:
    # --- MÓDULO DE ASO (ESTILO NOVO MAS COM SEU LAYOUT) ---
    st.markdown("### 🛡️ Controle de Vencimento de ASO")
    
    UNIDADES_ASO = {
        "CURITIBA": "145843404", "SÃO JOSÉ": "1537243911", 
        "ITAJAÍ": "517454238", "JOINVILLE": "1940989392"
    }

    def load_aso(gid):
        url = f"https://docs.google.com/spreadsheets/d/{SHEET_ID_ASO}/export?format=csv&gid={gid}"
        return pd.read_csv(url)

    unidade_sel = st.sidebar.selectbox("Selecione a Unidade ASO", list(UNIDADES_ASO.keys()))
    df_aso = load_aso(UNIDADES_ASO[unidade_sel])

    if not df_aso.empty:
        hoje = datetime.now()
        prazo_10 = hoje + timedelta(days=10)
        df_aso['Venc'] = pd.to_datetime(df_aso['Venc'], dayfirst=True, errors='coerce')
        alertas = df_aso[df_aso['Venc'] <= prazo_10].copy()
        
        st.metric("ASOs em Alerta (10 dias)", len(alertas))
        st.dataframe(alertas[['Nome', 'Cargo', 'Setor', 'Venc']], use_container_width=True)
        # (Lógica de gerar solicitação de agendamento aqui dentro)

# RODAPÉ (MANTIDO)
st.markdown("""<div class="footer">© 2026 Gestão Documentos | Versão 1.0 | Desenvolvido por: Dilceu Junior</div>""", unsafe_allow_html=True)
