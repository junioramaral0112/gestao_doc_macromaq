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
from urllib.parse import quote
import base64

# --- CONFIGURAÇÕES GLOBAIS ---
st.set_page_config(page_title="Portal SSMA - Junior 360", layout="wide")

# LÓGICA DE CAMINHOS (Adaptada para GitHub/Local)
BASE_PATH = os.path.dirname(__file__) if "__file__" in locals() else "."
ASSETS_PATH = os.path.join(BASE_PATH, "assets")

# Arquivos de Imagem e Templates
FUNDO_PATH = os.path.join(ASSETS_PATH, "fundo.png")
LOGO_PATH = os.path.join(ASSETS_PATH, "logo.png")
LOGO_CLINICA = os.path.join(ASSETS_PATH, "adivitta.png")
TEMPLATE_FICHA = os.path.join(ASSETS_PATH, "template_ficha.xlsx")
# ... (Mapeie todos os templates aqui usando ASSETS_PATH)

# IDs das Planilhas (ASO e Documentos)
SHEET_ID_ASO = "1G_oVT9gK-n_jGh5R4g65qUwK_MfQGvCX-SA4NHNNflU"
SHEET_ID_DOCS = "1y98U3eK7JXJqQaMC0i7eFbwpvp97Nuyeml5Dis0UCUg"

# --- FUNÇÕES DE LAYOUT (Mantendo seu estilo original) ---
def get_base64(bin_file):
    if not os.path.exists(bin_file): return ""
    with open(bin_file, 'rb') as f: data = f.read()
    return base64.b64encode(data).decode()

def aplicar_layout():
    fundo = get_base64(FUNDO_PATH)
    logo = get_base64(LOGO_PATH)
    st.markdown(f"""
        <style>
        .stApp {{ background-image: url("data:image/png;base64,{fundo}"); background-size: cover; background-attachment: fixed; }}
        .header-container {{ display: flex; align-items: center; background: rgba(255,255,255,0.9); padding: 20px; border-radius: 15px; margin-bottom: 30px; }}
        .footer {{ position: fixed; left: 0; bottom: 0; width: 100%; background: rgba(0,0,0,0.8); color: white; text-align: center; padding: 10px; z-index: 999; }}
        </style>
        <div class="header-container">
            <img src="data:image/png;base64,{logo}" width="300">
            <h1 style="margin-left:25px;color:#2c3e50;">Portal SSMA Integrado</h1>
        </div>
    """, unsafe_allow_html=True)

aplicar_layout()

# --- MENU DE NAVEGAÇÃO ---
st.sidebar.markdown("## 🛠️ Menu de Ferramentas")
escolha = st.sidebar.radio("Selecione o Módulo:", ["Painel de Controle ASO", "Automação de Docs (EPI/OS)"])

# ==========================================
# MÓDULO 1: CONTROLE DE ASO
# ==========================================
if escolha == "Painel de Controle ASO":
    st.subheader("🛡️ Gestão de Vencimentos de ASO")
    
    UNIDADES_ASO = {
        "CURITIBA": "145843404", "SÃO JOSÉ": "1537243911", 
        "ITAJAÍ": "517454238", "JOINVILLE": "1940989392"
    }
    
    def load_aso(gid):
        url = f"https://docs.google.com/spreadsheets/d/{SHEET_ID_ASO}/export?format=csv&gid={gid}"
        return pd.read_csv(url)

    unidade_aso = st.sidebar.selectbox("Selecione a Unidade", list(UNIDADES_ASO.keys()))
    df_aso = load_aso(UNIDADES_ASO[unidade_aso])

    if not df_aso.empty:
        # Lógica de Alertas
        hoje = datetime.now()
        prazo = hoje + timedelta(days=10)
        df_aso['Venc'] = pd.to_datetime(df_aso['Venc'], dayfirst=True, errors='coerce')
        alertas = df_aso[df_aso['Venc'] <= prazo].copy()
        
        st.metric("ASOs em Alerta (10 dias)", len(alertas))
        st.dataframe(alertas[['Nome', 'Cargo', 'Setor', 'Venc']], use_container_width=True)
        
        # Gerador de Documento ASO (Sua função gerar_docx entra aqui)
        # ... [Insira sua função gerar_docx aqui] ...

# ==========================================
# MÓDULO 2: AUTOMAÇÃO DE DOCS (SEU CÓDIGO ORIGINAL)
# ==========================================
else:
    st.subheader("📝 Gerador de Fichas, OS e Certificados")
    # Coloque aqui toda a lógica do seu segundo código 
    # (selectbox de colaborador, botões de processar, etc.)

# RODAPÉ
st.markdown("""<div class="footer">© 2026 Gestão Documentos | Desenvolvido por: Dilceu Junior</div>""", unsafe_allow_html=True)
