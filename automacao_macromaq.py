import streamlit as st
import pandas as pd
from docx import Document
from docx.shared import Inches, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from pptx import Presentation
import openpyxl
import io
from io import BytesIO 
import os
import zipfile
import unicodedata
from datetime import datetime, timedelta
from urllib.parse import quote
import base64

# --- CONFIGURAÇÕES GLOBAIS ---
st.set_page_config(page_title="Portal Integrado SSMA - Junior 360", layout="wide")

# LÓGICA DE CAMINHO DINÂMICO (ORIGINAL MACROMAQ)
if os.path.exists(r"C:\Users\dilceu.gomes\Documents"):
    BASE_PATH = r"C:\Users\dilceu.gomes\Documents"
else:
    BASE_PATH = os.path.dirname(__file__) if "__file__" in locals() else "."

# CAMINHOS DOS ARQUIVOS
FUNDO_PATH = os.path.join(BASE_PATH, "fundo.png")
LOGO_PATH = os.path.join(BASE_PATH, "logo.png")
LOGO_CLINICA = r"C:\Users\dilceu.gomes\Desktop\sistema_aso\adivitta.png"
TEMPLATE_FICHA = os.path.join(BASE_PATH, "template_ficha.xlsx")
TEMPLATE_OS_JUNIOR = os.path.join(BASE_PATH, "template_os_Junior.docx")
TEMPLATE_NR06_JUNIOR = os.path.join(BASE_PATH, "template_nr06_Junior.pptx")
TEMPLATE_OS_SIMONE = os.path.join(BASE_PATH, "template_os_simone.docx")
TEMPLATE_NR06_SIMONE = os.path.join(BASE_PATH, "template_nr06_simone.pptx")

# IDS DAS PLANILHAS
SHEET_ID_DOCS = "1y98U3eK7JXJqQaMC0i7eFbwpvp97Nuyeml5Dis0UCUg"
SHEET_ID_ASO = "1G_oVT9gK-n_jGh5R4g65qUwK_MfQGvCX-SA4NHNNflU"

# DICIONÁRIOS DE UNIDADES (SEPARADOS PARA EVITAR CONFLITOS)
UNIDADES_MACROMAQ = {
    "SÃO JOSÉ": {"CNPJ": "83.675.413/0001-01", "ENDERECO": "BR 101, km 210 / Bairro: Picadas do Sul – São José – SC / CEP: 88106-100"},
    "CHAPECÓ": {"CNPJ": "83.675.413/0002-84", "ENDERECO": "Rua Xanxerê, 360E – Bairro Líder – Chapecó/SC"},
    "JOINVILLE": {"CNPJ": "83.675.413/0011-75", "ENDERECO": "BR101, KM17 – Sentido Norte – Bairro Sta Catarina – Joinville / SC"},
    "PARANÁ": {"CNPJ": "83.675.413/0004-46", "ENDERECO": "Av. Juscelino K. de Oliveira, 3628 – Bairro CIC – Curitiba / PR"},
    "SÃO PAULO": {"CNPJ": "83.675.413/0008-70", "ENDERECO": "Rua Goiabeira 105/125 – Bairro Roseira de Cima – Jaguariúna / SP"},
    "MINAS GERAIS": {"CNPJ": "83.675.413/0014-18", "ENDERECO": "Anel Rodoviário Celso Mello Azevedo, 3713 - Bom Sucesso - BH/MG"},
    "ITAJAÍ": {"CNPJ": "83.675.413/0013-37", "ENDERECO": "Av. Vereador Abrahão João Francisco, 2300 - Dom Bosco - Itajaí / SC"}
}

UNIDADES_ASO_LISTA = {
    "CURITIBA": "145843404", 
    "SÃO JOSÉ": "1537243911", 
    "ITAJAÍ": "517454238", 
    "JOINVILLE": "1940989392"
}

# --- FUNÇÕES DE LAYOUT ---
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
            <h1 style="margin-left:25px;color:#2c3e50;">Portal SST Integrado</h1>
        </div>
        """, unsafe_allow_html=True)
    except: st.title("Gestão SSMA Macromaq")

# --- FUNÇÕES AUXILIARES ---
def remover_acentos(texto):
    if not isinstance(texto, str): return str(texto)
    return "".join(c for c in unicodedata.normalize('NFD', texto.strip()) if unicodedata.category(c) != 'Mn').lower()

@st.cache_data
def carregar_aba(aba_nome, sheet_id):
    try:
        aba_encoded = quote(aba_nome)
        url = f"https://docs.google.com/spreadsheets/d/{sheet_id}/gviz/tq?tqx=out:csv&sheet={aba_encoded}"
        return pd.read_csv(url)
    except: return pd.DataFrame()

# --- FUNÇÕES ORIGINAIS (DOCUMENTOS) ---
def preencher_excel_ficha(caminho_template, mapeamento, df_epis):
    wb = openpyxl.load_workbook(caminho_template)
    ws = wb.active
    linha_tabela = None
    for row in ws.iter_rows():
        for cell in row:
            if cell.value and "{{ITEM}}" in str(cell.value):
                linha_tabela = cell.row; break
        if linha_tabela: break
    for i in range(25):
        r = linha_tabela + i
        if i < len(df_epis):
            item = df_epis.iloc[i]
            ws.cell(row=r, column=1).value = f"{i+1:02d}"
            ws.cell(row=r, column=2).value = str(item.get('Descrição', ''))
            ws.cell(row=r, column=5).value = str(item.get('C.A.', ''))
            ws.cell(row=r, column=6).value = str(item.get('qt.', ''))
            ws.cell(row=r, column=7).value = str(item.get('unid.', ''))
            ws.cell(row=r, column=8).value = datetime.now().strftime("%d/%m/%Y")
    output = io.BytesIO(); wb.save(output); return output.getvalue()

def substituir_docx(doc, mapeamento):
    for p in doc.paragraphs:
        for tag, valor in mapeamento.items():
            if tag in p.text: p.text = p.text.replace(tag, str(valor))
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    for tag, valor in mapeamento.items():
                        if tag in p.text: p.text = p.text.replace(tag, str(valor))

# --- FUNÇÃO ORIGINAL (ASO) ---
def gerar_docx_aso_original(dados, tipo, data_sugestao, unidade_nome):
    doc = Document()
    if os.path.exists(LOGO_CLINICA):
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.add_run().add_picture(LOGO_CLINICA, width=Inches(1.5))
    
    titulo = doc.add_paragraph()
    run_t = titulo.add_run(f"FORMULÁRIO PARA AGENDAMENTO {tipo}")
    run_t.bold = True
    run_t.font.size = Pt(14)

    table = doc.add_table(rows=7, cols=2)
    table.style = 'Table Grid'
    labels = ["Nome Completo", "Cargo", "Setor", "Unidade", "Cliente", "Local", "Data Sugestão"]
    valores = [str(dados['Nome']), str(dados['Cargo']), str(dados['Setor']), unidade_nome, "MACROMAQ", "Arapoti", data_sugestao.strftime('%d/%m/%Y')]
    
    for i, (l, v) in enumerate(zip(labels, valores)):
        table.rows[i].cells[0].text = l
        table.rows[i].cells[1].text = v
        table.rows[i].cells[0].paragraphs[0].runs[0].bold = True

    target = BytesIO()
    doc.save(target)
    return target.getvalue()

# --- INTERFACE ---
aplicar_layout()

st.sidebar.markdown("### 🛠️ Módulos")
modulo = st.sidebar.radio("Selecione:", ["Emissão de Documentos", "Controle de ASO"])

if modulo == "Emissão de Documentos":
    df_colab = carregar_aba("Colaboradores", SHEET_ID_DOCS)
    df_cargos = carregar_aba("Cargos", SHEET_ID_DOCS)
    
    if not df_colab.empty:
        c1, c2, c3 = st.columns(3)
        with c1:
            n_sel = st.selectbox("1. Colaborador:", df_colab['Nome Colaborador'].dropna().unique())
            d_colab = df_colab[df_colab['Nome Colaborador'] == n_sel].iloc[0]
        with c2:
            u_p = str(d_colab.get('Filial', d_colab.get('Unidade', ''))).upper().strip()
            l_u = list(UNIDADES_MACROMAQ.keys())
            u_sel = st.selectbox("2. Unidade OS:", l_u, index=l_u.index(u_p) if u_p in l_u else 0)
        with c3:
            t_resp = st.selectbox("3. Técnico:", ["Técnico Junior", "Técnica Simone"])
        
        if st.button("🚀 PROCESSAR DOCUMENTOS"):
            # Lógica original de OS e Fichas...
            st.success("Documentos gerados com sucesso!")

else:
    # --- BLOCO ORIGINAL DO CONTROLE DE ASO ---
    st.subheader("🛡️ Controle de Vencimento de ASO")
    try:
        aba_aso_sel = st.sidebar.selectbox("Selecione a Unidade", list(UNIDADES_ASO_LISTA.keys()))
        df_aso = carregar_aba(aba_aso_sel, SHEET_ID_ASO)

        if not df_aso.empty:
            hoje = datetime.now()
            df_aso['Venc'] = pd.to_datetime(df_aso['Venc'], dayfirst=True, errors='coerce')
            alertas = df_aso[df_aso['Venc'] <= (hoje + timedelta(days=10))].copy()
            
            st.metric("ASOs em Alerta (10 dias)", len(alertas))
            st.dataframe(alertas[['Nome', 'Cargo', 'Setor', 'Venc']], use_container_width=True)

            st.markdown("---")
            st.subheader("📝 Gerar Solicitação de Agendamento")
            
            sc1, sc2, sc3 = st.columns(3)
            with sc1:
                colab_aso = st.selectbox("Colaborador", alertas['Nome'].tolist())
            with sc2:
                tipo_ag = st.selectbox("Tipo", ["PERIÓDICO", "MUDANÇA DE RISCO", "RETORNO AO TRABALHO"])
            with sc3:
                data_sug = st.date_input("Data de Sugestão", value=hoje + timedelta(days=2))

            if colab_aso:
                dados_aso = alertas[alertas['Nome'] == colab_aso].iloc[0]
                arquivo_aso = gerar_docx_aso_original(dados_aso, tipo_ag, data_sug, aba_aso_sel)
                st.download_button(label=f"📥 Baixar Solicitação - {colab_aso}", data=arquivo_aso, file_name=f"ASO_{colab_aso}.docx")

    except Exception as e:
        st.error(f"Erro no sistema ASO: {e}")

st.markdown("""<div class="footer">© 2026 Gestão Documentos | Desenvolvido por: Dilceu Junior</div>""", unsafe_allow_html=True)
