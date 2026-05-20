import streamlit as st
import pandas as pd
from docx import Document
from pptx import Presentation
import openpyxl
import io
import os
import zipfile
import unicodedata
import subprocess  # Adicionado para a conversão de PDF
from datetime import datetime
from urllib.parse import quote
import base64

# --- CONFIGURAÇÕES ---
st.set_page_config(page_title="Automação SSMA Macromaq", layout="wide")

# Usar o diretório atual para facilitar o deploy no GitHub/Streamlit Cloud
BASE_PATH = os.getcwd()

# Mantendo seu padrão original de caminhos na raiz do repositório
FUNDO_PATH = os.path.join(BASE_PATH, "fundo.png")
LOGO_PATH = os.path.join(BASE_PATH, "logo.png")
TEMPLATE_FICHA = os.path.join(BASE_PATH, "template_ficha.xlsx")
TEMPLATE_OS_JUNIOR = os.path.join(BASE_PATH, "template_os_Junior.docx")
TEMPLATE_NR06_JUNIOR = os.path.join(BASE_PATH, "template_nr06_Junior.pptx")
TEMPLATE_OS_SIMONE = os.path.join(BASE_PATH, "template_os_simone.docx")
TEMPLATE_NR06_SIMONE = os.path.join(BASE_PATH, "template_nr06_simone.pptx")

SHEET_ID = "1y98U3eK7JXJqQaMC0i7eFbwpvp97Nuyeml5Dis0UCUg"

# --- UNIDADES ---
UNIDADES = {
    "SÃO JOSÉ": {
        "CNPJ": "83.675.413/0001-01",
        "ENDERECO": "BR 101, km 210 / Bairro: Picadas do Sul – São José – SC / CEP: 88106-100"
    },
    "CHAPECÓ": {
        "CNPJ": "83.675.413/0002-84",
        "ENDERECO": "Rua Xanxerê, 360E – Bairro Líder – Chapecó/SC"
    },
    "JOINVILLE": {
        "CNPJ": "83.675.413/0011-75",
        "ENDERECO": "BR101, KM17 – Sentido Norte – Bairro Sta Catarina – Joinville / SC"
    },
    "PARANÁ": {
        "CNPJ": "83.675.413/0004-46",
        "ENDERECO": "Av. Juscelino K. de Oliveira, 3628 – Bairro CIC – Curitiba / PR"
    },
    "SÃO PAULO": {
        "CNPJ": "83.675.413/0008-70",
        "ENDERECO": "Rua Goiabeira 105/125 – Bairro Roseira de Cima – Jaguariúna / SP"
    },
    "MINAS GERAIS": {
        "CNPJ": "83.675.413/0014-18",
        "ENDERECO": "Anel Rodoviário Celso Mello Azevedo, 3713 - Bom Sucesso - BH/MG"
    },
    "ITAJAÍ": {
        "CNPJ": "83.675.413/0013-37",
        "ENDERECO": "Av. Vereador Abrahão João Francisco, 2300 - Dom Bosco - Itajaí / SC"
    }
}

# --- LAYOUT E CSS ---
def get_base64(bin_file):
    if not os.path.exists(bin_file):
        return ""
    with open(bin_file, 'rb') as f:
        data = f.read()
    return base64.b64encode(data).decode()

def aplicar_layout():
    try:
        fundo = get_base64(FUNDO_PATH)
        logo = get_base64(LOGO_PATH)
        st.markdown(f"""
        <style>
        /* REMOVER BARRA LATERAL E NAVEGAÇÃO NATIVA */
        [data-testid="stSidebar"], [data-testid="stSidebarNav"] {{
            display: none;
        }}

        .stApp {{
            background-image: url("data:image/png;base64,{fundo}");
            background-size: cover;
            background-position: center;
            background-attachment: fixed;
        }}
        .stSelectbox label, div[data-testid="stCheckbox"] label p {{
            color: white !important;
            background: rgba(0,0,0,0.7);
            padding: 5px 12px;
            border-radius: 8px;
            font-weight: bold;
        }}
        .stButton > button {{
            background: #2c3e50;
            color: #f9cc0b;
            border: 2px solid #f9cc0b;
            border-radius: 10px;
            height: 55px;
            font-weight: bold;
            width: 100%;
            font-size: 18px;
        }}
        .header-container {{
            display: flex;
            align-items: center;
            background: rgba(255,255,255,0.9);
            padding: 20px;
            border-radius: 15px;
            margin-bottom: 30px;
        }}
        .footer {{
            position: fixed;
            left: 0;
            bottom: 0;
            width: 100%;
            background: rgba(0,0,0,0.8);
            color: white;
            text-align: center;
            padding: 10px;
            font-size: 13px;
            z-index: 999;
        }}
        </style>
        <div class="header-container">
            <img src="data:image/png;base64,{logo}" width="320">
            <h1 style="margin-left:25px;color:#2c3e50;">Gestão SSMA Documentos</h1>
        </div>
        """, unsafe_allow_html=True)
    except:
        st.title("Automação SSMA Macromaq")

# --- FUNÇÕES AUXILIARES ---
def remover_acentos(texto):
    if not isinstance(texto, str): return str(texto)
    return "".join(c for c in unicodedata.normalize('NFD', texto.strip()) if unicodedata.category(c) != 'Mn').lower()

@st.cache_data(ttl=300)
def carregar_aba(aba_nome):
    try:
        url = f"https://docs.google.com/spreadsheets/d/{SHEET_ID}/gviz/tq?tqx=out:csv&sheet={quote(aba_nome)}"
        return pd.read_csv(url)
    except:
        return pd.DataFrame()

def data_extenso_pt():
    meses = {1:"Janeiro", 2:"Fevereiro", 3:"Março", 4:"Abril", 5:"Maio", 6:"Junho", 7:"Julho", 8:"Agosto", 9:"Setembro", 10:"Outubro", 11:"Novembro", 12:"Dezembro"}
    agora = datetime.now()
    return agora.strftime(f"%d de {meses[agora.month]} de %Y")

def formatar_matricula(valor):
    try: return str(int(float(valor))) if not pd.isna(valor) else ""
    except: return str(valor)

def formatar_cpf(cpf):
    cpf = ''.join(filter(str.isdigit, str(cpf))).zfill(11)
    return f"{cpf[:3]}.{cpf[3:6]}.{cpf[6:9]}-{cpf[9:]}"

def limpar_valor(valor):
    if pd.isna(valor): return ""
    texto = str(valor).strip()
    return "" if texto.lower() in ["nan", "na"] else texto

# Função para converter arquivos do Office para PDF usando o LibreOffice do Servidor
def converter_para_pdf_linux(conteudo_arquivo, nome_original):
    try:
        temp_input = os.path.join(BASE_PATH, nome_original)
        with open(temp_input, "wb") as f:
            f.write(conteudo_arquivo)
        
        subprocess.run([
            'libreoffice', '--headless', '--convert-to', 'pdf', temp_input,
            '--outdir', BASE_PATH
        ], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, check=True)
        
        nome_pdf = os.path.splitext(nome_original)[0] + ".pdf"
        temp_pdf_path = os.path.join(BASE_PATH, nome_pdf)
        
        if os.path.exists(temp_pdf_path):
            with open(temp_pdf_path, "rb") as f:
                pdf_bytes = f.read()
            
            os.remove(temp_input)
            os.remove(temp_pdf_path)
            return pdf_bytes, nome_pdf
    except Exception as e:
        pass
    return None, None

# --- PROCESSAMENTO DE DOCUMENTOS ---
def preencher_excel_ficha(caminho_template, mapeamento, df_epis):
    wb = openpyxl.load_workbook(caminho_template)
    ws = wb.active
    for row in ws.iter_rows():
        for cell in row:
            if isinstance(cell.value, str):
                for tag, valor in mapeamento.items():
                    if tag in cell.value:
                        cell.value = cell.value.replace(tag, str(valor))
        
    linha_tabela = None
    for row in ws.iter_rows():
        for cell in row:
            if cell.value and "{{ITEM}}" in str(cell.value):
                linha_tabela = cell.row
                break
        if linha_tabela: break
        
    if not linha_tabela: raise Exception("Tag {{ITEM}} não encontrada.")
    
    for i in range(25):
        r = linha_tabela + i
        if i < len(df_epis):
            item = df_epis.iloc[i]
            ws.cell(row=r, column=1).value = f"{i+1:02d}"
            ws.cell(row=r, column=2).value = limpar_valor(item.get('Descrição', ''))
            ws.cell(row=r, column=5).value = limpar_valor(item.get('C.A.', ''))
            ws.cell(row=r, column=6).value = limpar_valor(item.get('qt.', ''))
            ws.cell(row=r, column=7).value = limpar_valor(item.get('unid.', ''))
            ws.cell(row=r, column=8).value = datetime.now().strftime("%d/%m/%Y")
        else:
            for c in [1, 2, 5, 6, 7, 8]: ws.cell(row=r, column=c).value = ""
            
    output = io.BytesIO()
    wb.save(output)
    return output.getvalue()

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

def substituir_pptx(prs, mapeamento):
    for slide in prs.slides:
        for shape in slide.shapes:
            if hasattr(shape, "text_frame"):
                for paragraph in shape.text_frame.paragraphs:
                    for run in paragraph.runs:
                        for tag, valor in mapeamento.items():
                            if tag in run.text: run.text = run.text.replace(tag, str(valor))

# --- APP LOGIC ---
aplicar_layout()

if st.button("🔄 Atualizar Dados da Planilha Agora"):
    st.cache_data.clear()
    st.rerun()

df_colab = carregar_aba("Colaboradores")
df_cargos = carregar_aba("Cargos")

if not df_colab.empty and not df_cargos.empty:
    df_colab['Nome_Formatado'] = df_colab['Nome Colaborador'].astype(str).str.strip().str.title()
    
    col1, col2, col3 = st.columns(3)
    with col1:
        lista_nomes = sorted(df_colab['Nome_Formatado'].dropna().unique())
        nome_sel = st.selectbox("1. Selecione o Colaborador:", lista_nomes)
        dados_colab = df_colab[df_colab['Nome_Formatado'] == nome_sel].iloc[0]
        
    with col2:
        unidade_plan = str(dados_colab.get('Filial', dados_colab.get('Unidade', ''))).upper().strip()
        lista_unid = list(UNIDADES.keys())
        idx = lista_unid.index(unidade_plan) if unidade_plan in lista_unid else 0
        unid_sel = st.selectbox("2. Unidade para OS:", lista_unid, index=idx)
    with col3:
        tecnico_sel = st.selectbox("3. Técnico Responsável:", ["Técnico Junior", "Técnica Simone"])

    t_os = TEMPLATE_OS_JUNIOR if tecnico_sel == "Técnico Junior" else TEMPLATE_OS_SIMONE
    t_nr = TEMPLATE_NR06_JUNIOR if tecnico_sel == "Técnico Junior" else TEMPLATE_NR06_SIMONE

    st.markdown("<br>", unsafe_allow_html=True)
    c1, c2, c3 = st.columns(3)
    g_os, g_ficha, g_cert = c1.checkbox("OS", True), c2.checkbox("Ficha EPI", True), c3.checkbox("Certificado", True)

    # Caixa para decidir se quer embutir PDFs
    incluir_pdf = st.checkbox("📄 Incluir cópias em formato PDF no Kit", True)

    if st.button("🚀 PROCESSAR DOCUMENTOS"):
        with st.spinner("Gerando documentos..."):
            cargo = str(dados_colab['Cargo']).strip()
            arquivos = {}
            df_cargos['f_l'] = df_cargos['Função'].astype(str).apply(remover_acentos)
            desc_f = df_cargos[df_cargos['f_l'] == remover_acentos(cargo)]

            if not desc_f.empty:
                desc_atv = desc_f['Descrição da Atividade'].values[0]
                
                # 1. Processando Ordem de Serviço (DOCX -> PDF)
                if g_os:
                    doc = Document(t_os)
                    substituir_docx(doc, {"{{NOME}}": dados_colab['Nome Colaborador'], "{{FUNCAO}}": cargo.upper(), "{{CNPJ}}": UNIDADES[unid_sel]["CNPJ"], "{{ENDERECO}}": UNIDADES[unid_sel]["ENDERECO"], "{{SETOR}}": str(dados_colab.get('NomeLocal', '')), "{{DESCRICAO_ATIVIDADE}}": str(desc_atv), "{{DATA}}": datetime.now().strftime("%d/%m/%Y")})
                    b = io.BytesIO(); doc.save(b)
                    conteudo_docx = b.getvalue()
                    nome_docx = f"OS {nome_sel}.docx"
                    arquivos[nome_docx] = conteudo_docx
                    
                    if incluir_pdf:
                        pdf_bytes, nome_pdf = converter_para_pdf_linux(conteudo_docx, nome_docx)
                        if pdf_bytes: arquivos[nome_pdf] = pdf_bytes

                # 2. Processando Ficha de EPI (XLSX -> PDF) - ATUALIZADO
                if g_ficha:
                    df_e = carregar_aba(cargo)
                    if df_e.empty: df_e = carregar_aba(remover_acentos(cargo))
                    if not df_e.empty:
                        m_f = {"{{NOME}}": dados_colab['Nome Colaborador'], "{{MATRICULA}}": formatar_matricula(dados_colab.get('Matrícula', '')), "{{FUNCAO}}": cargo, "{{DATA_ADMISSAO}}": datetime.now().strftime("%d/%m/%Y"), "{{SETOR}}": str(dados_colab.get('NomeLocal', '')), "{{CENTRO_CUSTO}}": ""}
                        conteudo_xlsx = preencher_excel_ficha(TEMPLATE_FICHA, m_f, df_e)
                        nome_xlsx = f"Ficha EPI {nome_sel}.xlsx"
                        arquivos[nome_xlsx] = conteudo_xlsx
                        
                        # Nova verificação: Converte a Planilha gerada para PDF
                        if incluir_pdf:
                            pdf_bytes, nome_pdf = converter_para_pdf_linux(conteudo_xlsx, nome_xlsx)
                            if pdf_bytes: arquivos[nome_pdf] = pdf_bytes

                # 3. Processando Certificado NR06 (PPTX -> PDF)
                if g_cert:
                    prs = Presentation(t_nr)
                    substituir_pptx(prs, {"{{NOME}}": dados_colab['Nome Colaborador'], "{{CPF}}": formatar_cpf(dados_colab.get('CPF', '')), "{{FUNCAO}}": cargo, "{{DATA_TREINAMENTO}}": datetime.now().strftime("%d/%m/%Y"), "{{LOCAL_DATA}}": f"{unid_sel.title()}, {data_extenso_pt()}."})
                    b = io.BytesIO(); prs.save(b)
                    conteudo_pptx = b.getvalue()
                    nome_pptx = f"NR06 {nome_sel}.pptx"
                    arquivos[nome_pptx] = conteudo_pptx
                    
                    if incluir_pdf:
                        pdf_bytes, nome_pdf = converter_para_pdf_linux(conteudo_pptx, nome_pptx)
                        if pdf_bytes: arquivos[nome_pdf] = pdf_bytes

                # Se houver arquivos gerados, compacta tudo em ZIP
                if arquivos:
                    z_b = io.BytesIO()
                    with zipfile.ZipFile(z_b, "w") as z:
                        for n, d in arquivos.items(): z.writestr(n, d)
                    st.success("✅ Documentos prontos!")
                    st.download_button("📦 BAIXAR KIT COMPLETO (ZIP)", z_b.getvalue(), f"Kit_{nome_sel}.zip", use_container_width=True)
            else:
                st.error(f"Cargo '{cargo}' não encontrado na aba Cargos.")

st.markdown("""<div class="footer">© 2026 Gestão Documentos | Desenvolvido por: Dilceu Junior</div>""", unsafe_allow_html=True)
