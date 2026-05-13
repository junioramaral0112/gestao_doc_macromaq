import streamlit as st
import pandas as pd
from docx import Document
from docx.shared import Inches, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from pptx import Presentation
import openpyxl
import io
from io import BytesIO  # Importação vital para o ASO funcionar
import os
import zipfile
import unicodedata
from datetime import datetime, timedelta
from urllib.parse import quote
import base64

# --- CONFIGURAÇÕES GLOBAIS ---
st.set_page_config(page_title="Portal Integrado SSMA - Junior 360", layout="wide")

# LÓGICA DE CAMINHO DINÂMICO (MANTIDA)
if os.path.exists(r"C:\Users\dilceu.gomes\Documents"):
    BASE_PATH = r"C:\Users\dilceu.gomes\Documents"
else:
    BASE_PATH = os.path.dirname(__file__) if "__file__" in locals() else "."

# CAMINHOS DOS ARQUIVOS (ORIGINAIS)
FUNDO_PATH = os.path.join(BASE_PATH, "fundo.png")
LOGO_PATH = os.path.join(BASE_PATH, "logo.png")
LOGO_CLINICA = r"C:\Users\dilceu.gomes\Desktop\sistema_aso\adivitta.png"
TEMPLATE_FICHA = os.path.join(BASE_PATH, "template_ficha.xlsx")
TEMPLATE_OS_JUNIOR = os.path.join(BASE_PATH, "template_os_Junior.docx")
TEMPLATE_NR06_JUNIOR = os.path.join(BASE_PATH, "template_nr06_Junior.pptx")
TEMPLATE_OS_SIMONE = os.path.join(BASE_PATH, "template_os_simone.docx")
TEMPLATE_NR06_SIMONE = os.path.join(BASE_PATH, "template_nr06_simone.pptx")

SHEET_ID_DOCS = "1y98U3eK7JXJqQaMC0i7eFbwpvp97Nuyeml5Dis0UCUg"
SHEET_ID_ASO = "1G_oVT9gK-n_jGh5R4g65qUwK_MfQGvCX-SA4NHNNflU"

# UNIDADES PARA EMISSÃO DE DOCS (ORIGINAL)
UNIDADES_DOCS = {
    "SÃO JOSÉ": {"CNPJ": "83.675.413/0001-01", "ENDERECO": "BR 101, km 210 / Bairro: Picadas do Sul – São José – SC / CEP: 88106-100"},
    "CHAPECÓ": {"CNPJ": "83.675.413/0002-84", "ENDERECO": "Rua Xanxerê, 360E – Bairro Líder – Chapecó/SC"},
    "JOINVILLE": {"CNPJ": "83.675.413/0011-75", "ENDERECO": "BR101, KM17 – Sentido Norte – Bairro Sta Catarina – Joinville / SC"},
    "PARANÁ": {"CNPJ": "83.675.413/0004-46", "ENDERECO": "Av. Juscelino K. de Oliveira, 3628 – Bairro CIC – Curitiba / PR"},
    "SÃO PAULO": {"CNPJ": "83.675.413/0008-70", "ENDERECO": "Rua Goiabeira 105/125 – Bairro Roseira de Cima – Jaguariúna / SP"},
    "MINAS GERAIS": {"CNPJ": "83.675.413/0014-18", "ENDERECO": "Anel Rodoviário Celso Mello Azevedo, 3713 - Bom Sucesso - BH/MG"},
    "ITAJAÍ": {"CNPJ": "83.675.413/0013-37", "ENDERECO": "Av. Vereador Abrahão João Francisco, 2300 - Dom Bosco - Itajaí / SC"}
}

UNIDADES_ASO = {"CURITIBA": "145843404", "SÃO JOSÉ": "1537243911", "ITAJAÍ": "517454238", "JOINVILLE": "1940989392"}

# --- ESTILO E LAYOUT (ORIGINAL) ---
def get_base64(bin_file):
    if not os.path.exists(bin_file): return ""
    with open(bin_file, 'rb') as f: data = f.read()
    return base64.b64encode(data).decode()

def aplicar_layout():
    try:
        fundo, logo = get_base64(FUNDO_PATH), get_base64(LOGO_PATH)
        st.markdown(f"""
        <style>
        .stApp {{ background-image: url("data:image/png;base64,{fundo}"); background-size: cover; background-attachment: fixed; }}
        .stSelectbox label, div[data-testid="stCheckbox"] label p {{ color: white !important; background: rgba(0,0,0,0.7); padding: 5px 12px; border-radius: 8px; font-weight: bold; }}
        .stButton > button {{ background: #2c3e50; color: #f9cc0b; border: 2px solid #f9cc0b; border-radius: 10px; height: 55px; font-weight: bold; width: 100%; font-size: 18px; }}
        .header-container {{ display: flex; align-items: center; background: rgba(255,255,255,0.9); padding: 20px; border-radius: 15px; margin-bottom: 30px; }}
        .footer {{ position: fixed; left: 0; bottom: 0; width: 100%; background: rgba(0,0,0,0.8); color: white; text-align: center; padding: 10px; z-index: 999; }}
        </style>
        <div class="header-container"><img src="data:image/png;base64,{logo}" width="320"><h1 style="margin-left:25px;color:#2c3e50;">Gestão SSMA Macromaq</h1></div>
        """, unsafe_allow_html=True)
    except: st.title("Gestão SSMA Macromaq")

# --- FUNÇÕES AUXILIARES (ORIGINAIS) ---
def remover_acentos(t): return "".join(c for c in unicodedata.normalize('NFD', str(t).strip()) if unicodedata.category(c) != 'Mn').lower()
def formatar_cpf(c): c = ''.join(filter(str.isdigit, str(c))).zfill(11); return f"{c[:3]}.{c[3:6]}.{c[6:9]}-{c[9:]}"
def limpar_v(v): return "" if pd.isna(v) or str(v).strip().lower() in ["nan", "na"] else str(v).strip()

@st.cache_data
def carregar_aba(aba, sid):
    try: url = f"https://docs.google.com/spreadsheets/d/{sid}/gviz/tq?tqx=out:csv&sheet={quote(aba)}"; return pd.read_csv(url)
    except: return pd.DataFrame()

# --- PROCESSAMENTO DE DOCS ---
def substituir_docx(doc, mapa):
    for p in doc.paragraphs:
        for t, v in mapa.items():
            if t in p.text: p.text = p.text.replace(t, str(v))
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    for t, v in mapa.items():
                        if t in p.text: p.text = p.text.replace(t, str(v))

def substituir_pptx(prs, mapa):
    for slide in prs.slides:
        for shape in slide.shapes:
            if hasattr(shape, "text_frame"):
                for p in shape.text_frame.paragraphs:
                    for run in p.runs:
                        for t, v in mapa.items():
                            if t in run.text: run.text = run.text.replace(t, str(v))

def gerar_aso_docx(dados, tipo, data_s, unid):
    doc = Document()
    if os.path.exists(LOGO_CLINICA):
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.add_run().add_picture(LOGO_CLINICA, width=Inches(1.5))
    run_t = doc.add_paragraph().add_run(f"FORMULÁRIO PARA AGENDAMENTO {tipo}")
    run_t.bold, run_t.font.size = True, Pt(14)
    table = doc.add_table(rows=7, cols=2)
    table.style = 'Table Grid'
    l = ["Nome Completo", "Cargo", "Setor", "Unidade", "Cliente", "Local", "Data Sugestão"]
    v = [str(dados['Nome']), str(dados['Cargo']), str(dados['Setor']), unid, "MACROMAQ", "Arapoti", data_s.strftime('%d/%m/%Y')]
    for i, (label, val) in enumerate(zip(l, v)):
        table.rows[i].cells[0].text = label
        table.rows[i].cells[1].text = val
        table.rows[i].cells[0].paragraphs[0].runs[0].bold = True
    target = BytesIO()
    doc.save(target)
    return target.getvalue()

# --- EXECUÇÃO ---
aplicar_layout()
st.sidebar.markdown("### 🗂️ Menu Principal")
menu = st.sidebar.radio("Selecione:", ["Documentos SST", "Controle de ASO"])

if menu == "Documentos SST":
    df_colab, df_cargos = carregar_aba("Colaboradores", SHEET_ID_DOCS), carregar_aba("Cargos", SHEET_ID_DOCS)
    if not df_colab.empty:
        c1, c2, c3 = st.columns(3)
        with c1:
            n_sel = st.selectbox("Colaborador:", df_colab['Nome Colaborador'].dropna().unique())
            d_colab = df_colab[df_colab['Nome Colaborador'] == n_sel].iloc[0]
        with c2:
            u_p = str(d_colab.get('Filial', d_colab.get('Unidade', ''))).upper().strip()
            l_u = list(UNIDADES_DOCS.keys())
            u_sel = st.selectbox("Unidade:", l_u, index=l_u.index(u_p) if u_p in l_u else 0)
        with c3: t_resp = st.selectbox("Técnico:", ["Técnico Junior", "Técnica Simone"])
        
        t_os = TEMPLATE_OS_JUNIOR if t_resp == "Técnico Junior" else TEMPLATE_OS_SIMONE
        
        if st.button("🚀 PROCESSAR DOCUMENTOS"):
            cargo = str(d_colab['Cargo']).strip()
            df_cargos['f_l'] = df_cargos['Função'].astype(str).apply(remover_acentos)
            desc_f = df_cargos[df_cargos['f_l'] == remover_acentos(cargo)]
            if not desc_f.empty:
                doc = Document(t_os)
                substituir_docx(doc, {"{{NOME}}": n_sel, "{{FUNCAO}}": cargo.upper(), "{{CNPJ}}": UNIDADES_DOCS[u_sel]["CNPJ"], "{{ENDERECO}}": UNIDADES_DOCS[u_sel]["ENDERECO"], "{{SETOR}}": str(d_colab.get('NomeLocal', '')), "{{DATA}}": datetime.now().strftime("%d/%m/%Y")})
                b = io.BytesIO(); doc.save(b)
                st.success("✅ Documento Pronto!")
                st.download_button("📥 Baixar OS", b.getvalue(), f"OS_{n_sel}.docx")

else:
    u_aso = st.sidebar.selectbox("Unidade ASO:", list(UNIDADES_ASO.keys()))
    df_aso = carregar_aba(u_aso, SHEET_ID_ASO)
    if not df_aso.empty:
        hoje = datetime.now()
        df_aso['Venc'] = pd.to_datetime(df_aso['Venc'], dayfirst=True, errors='coerce')
        alertas = df_aso[df_aso['Venc'] <= (hoje + timedelta(days=10))].copy()
        st.metric("ASOs em Alerta (10 dias)", len(alertas))
        st.dataframe(alertas[['Nome', 'Cargo', 'Setor', 'Venc']], use_container_width=True)
        
        st.markdown("---")
        sc1, sc2, sc3 = st.columns(3)
        with sc1: n_aso = st.selectbox("Colaborador:", alertas['Nome'].tolist())
        with sc2: t_aso = st.selectbox("Tipo:", ["PERIÓDICO", "MUDANÇA DE RISCO", "RETORNO"])
        with sc3: d_aso = st.date_input("Data Sugestão:", value=hoje + timedelta(days=2))
        
        if n_aso:
            d_sel = alertas[alertas['Nome'] == n_aso].iloc[0]
            btn_aso = gerar_aso_docx(d_sel, t_aso, d_aso, u_aso)
            st.download_button("📥 Baixar Solicitação", btn_aso, f"ASO_{n_aso}.docx")

st.markdown("""<div class="footer">© 2026 Gestão Documentos | Desenvolvido por: Dilceu Junior</div>""", unsafe_allow_html=True)
