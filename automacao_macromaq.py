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

# --- CONFIGURAÇÕES ---
st.set_page_config(page_title="Portal Integrado SSMA - Junior 360", layout="wide")

# LÓGICA DE CAMINHO DINÂMICO
if os.path.exists(r"C:\Users\dilceu.gomes\Documents"):
    BASE_PATH = r"C:\Users\dilceu.gomes\Documents"
else:
    BASE_PATH = os.path.dirname(__file__) if "__file__" in locals() else "."

# Caminhos dos arquivos
FUNDO_PATH = os.path.join(BASE_PATH, "fundo.png")
LOGO_PATH = os.path.join(BASE_PATH, "logo.png")
LOGO_CLINICA = os.path.join(BASE_PATH, "adivitta.png") # Caminho relativo para GitHub
TEMPLATE_FICHA = os.path.join(BASE_PATH, "template_ficha.xlsx")
TEMPLATE_OS_JUNIOR = os.path.join(BASE_PATH, "template_os_Junior.docx")
TEMPLATE_NR06_JUNIOR = os.path.join(BASE_PATH, "template_nr06_Junior.pptx")
TEMPLATE_OS_SIMONE = os.path.join(BASE_PATH, "template_os_simone.docx")
TEMPLATE_NR06_SIMONE = os.path.join(BASE_PATH, "template_nr06_simone.pptx")

# IDs das Planilhas
SHEET_ID_DOCS = "1y98U3eK7JXJqQaMC0i7eFbwpvp97Nuyeml5Dis0UCUg"
SHEET_ID_ASO = "1G_oVT9gK-n_jGh5R4g65qUwK_MfQGvCX-SA4NHNNflU"

# --- UNIDADES EMISSÃO DOCS ---
UNIDADES_DOCS = {
    "SÃO JOSÉ": {"CNPJ": "83.675.413/0001-01", "ENDERECO": "BR 101, km 210 / Bairro: Picadas do Sul – São José – SC / CEP: 88106-100"},
    "CHAPECÓ": {"CNPJ": "83.675.413/0002-84", "ENDERECO": "Rua Xanxerê, 360E – Bairro Líder – Chapecó/SC"},
    "JOINVILLE": {"CNPJ": "83.675.413/0011-75", "ENDERECO": "BR101, KM17 – Sentido Norte – Bairro Sta Catarina – Joinville / SC"},
    "PARANÁ": {"CNPJ": "83.675.413/0004-46", "ENDERECO": "Av. Juscelino K. de Oliveira, 3628 – Bairro CIC – Curitiba / PR"},
    "SÃO PAULO": {"CNPJ": "83.675.413/0008-70", "ENDERECO": "Rua Goiabeira 105/125 – Bairro Roseira de Cima – Jaguariúna / SP"},
    "MINAS GERAIS": {"CNPJ": "83.675.413/0014-18", "ENDERECO": "Anel Rodoviário Celso Mello Azevedo, 3713 - Bom Sucesso - BH/MG"},
    "ITAJAÍ": {"CNPJ": "83.675.413/0013-37", "ENDERECO": "Av. Vereador Abrahão João Francisco, 2300 - Dom Bosco - Itajaí / SC"}
}

# --- UNIDADES CONTROLE ASO ---
UNIDADES_ASO = {
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
        .stSelectbox label, div[data-testid="stCheckbox"] label p {{
            color: white !important; background: rgba(0,0,0,0.7); padding: 5px 12px; border-radius: 8px; font-weight: bold;
        }}
        .stButton > button {{
            background: #2c3e50; color: #f9cc0b; border: 2px solid #f9cc0b; border-radius: 10px; height: 55px; font-weight: bold; width: 100%; font-size: 18px;
        }}
        .header-container {{
            display: flex; align-items: center; background: rgba(255,255,255,0.9); padding: 20px; border-radius: 15px; margin-bottom: 30px;
        }}
        .footer {{
            position: fixed; left: 0; bottom: 0; width: 100%; background: rgba(0,0,0,0.8); color: white; text-align: center; padding: 10px; font-size: 13px; z-index: 999;
        }}
        </style>
        <div class="header-container">
            <img src="data:image/png;base64,{logo}" width="320">
            <h1 style="margin-left:25px;color:#2c3e50;">Portal Integrado SSMA</h1>
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

def data_extenso_pt():
    meses = {1: "Janeiro", 2: "Fevereiro", 3: "Março", 4: "Abril", 5: "Maio", 6: "Junho", 7: "Julho", 8: "Agosto", 9: "Setembro", 10: "Outubro", 11: "Novembro", 12: "Dezembro"}
    agora = datetime.now()
    return agora.strftime(f"%d de {meses[agora.month]} de %Y")

def formatar_matricula(valor):
    if pd.isna(valor) or valor == "": return ""
    try: return str(int(float(valor)))
    except: return str(valor)

def formatar_cpf(cpf):
    cpf = ''.join(filter(str.isdigit, str(cpf))).zfill(11)
    return f"{cpf[:3]}.{cpf[3:6]}.{cpf[6:9]}-{cpf[9:]}"

def limpar_valor(valor):
    if pd.isna(valor): return ""
    texto = str(valor).strip()
    return "" if texto.lower() in ["nan", "na"] else texto

# --- PROCESSAMENTO DOCUMENTOS ---
def preencher_excel_ficha(caminho_template, mapeamento, df_epis):
    wb = openpyxl.load_workbook(caminho_template)
    ws = wb.active
    for row in ws.iter_rows():
        for cell in row:
            if isinstance(cell.value, str):
                for tag, valor in mapeamento.items():
                    if tag in cell.value: cell.value = cell.value.replace(tag, str(valor))
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
            ws.cell(row=r, column=2).value = limpar_valor(item.get('Descrição', ''))
            ws.cell(row=r, column=5).value = limpar_valor(item.get('C.A.', ''))
            ws.cell(row=r, column=6).value = limpar_valor(item.get('qt.', ''))
            ws.cell(row=r, column=7).value = limpar_valor(item.get('unid.', ''))
            ws.cell(row=r, column=8).value = datetime.now().strftime("%d/%m/%Y")
        else:
            for c in [1, 2, 5, 6, 7, 8]: ws.cell(row=r, column=c).value = ""
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

def substituir_pptx(prs, mapeamento):
    for slide in prs.slides:
        for shape in slide.shapes:
            if hasattr(shape, "text_frame"):
                for paragraph in shape.text_frame.paragraphs:
                    for run in paragraph.runs:
                        for tag, valor in mapeamento.items():
                            if tag in run.text: run.text = run.text.replace(tag, str(valor))

# --- FUNÇÃO GERADORA ASO ---
def gerar_docx_aso(dados, tipo, data_sugestao, unidade_nome):
    doc = Document()
    if os.path.exists(LOGO_CLINICA):
        section = doc.sections[0]
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = p.add_run()
        run.add_picture(LOGO_CLINICA, width=Inches(1.5))
    
    titulo = doc.add_paragraph()
    titulo.alignment = WD_ALIGN_PARAGRAPH.LEFT
    run_titulo = titulo.add_run(f"FORMULÁRIO PARA AGENDAMENTO {tipo}")
    run_titulo.bold = True
    run_titulo.font.size = Pt(14)

    table = doc.add_table(rows=7, cols=2)
    table.style = 'Table Grid'
    labels = ["Nome Completo", "Cargo", "Setor", "Unidade", "Cliente", "Local", "Data Sugestão"]
    valores = [str(dados['Nome']), str(dados['Cargo']), str(dados['Setor']), unidade_nome, "MACROMAQ", "Arapoti", data_sugestao.strftime('%d/%m/%Y')]
    
    for i, (label, val) in enumerate(zip(labels, valores)):
        table.rows[i].cells[0].text = label
        table.rows[i].cells[1].text = val
        table.rows[i].cells[0].paragraphs[0].runs[0].bold = True

    target = BytesIO()
    doc.save(target)
    return target.getvalue()

# --- APP ---
aplicar_layout()

# NAVEGAÇÃO LATERAL
st.sidebar.markdown("### 🛠️ Menu Principal")
menu_modulo = st.sidebar.radio("Selecione o Módulo:", ["Emissão de Documentos", "Controle de ASO"])

# ==========================================
# MÓDULO 1: EMISSÃO DE DOCUMENTOS (ORIGINAL)
# ==========================================
if menu_modulo == "Emissão de Documentos":
    df_colab = carregar_aba("Colaboradores", SHEET_ID_DOCS)
    df_cargos = carregar_aba("Cargos", SHEET_ID_DOCS)

    if not df_colab.empty and not df_cargos.empty:
        col1, col2, col3 = st.columns(3)
        with col1:
            nome_sel = st.selectbox("1. Selecione o Colaborador:", df_colab['Nome Colaborador'].dropna().unique())
            dados_colab = df_colab[df_colab['Nome Colaborador'] == nome_sel].iloc[0]
        with col2:
            unidade_p = str(dados_colab.get('Filial', dados_colab.get('Unidade', ''))).upper().strip()
            lista_u = list(UNIDADES_DOCS.keys())
            idx = lista_u.index(unidade_p) if unidade_p in lista_u else 0
            unid_sel = st.selectbox("2. Unidade para OS:", lista_u, index=idx)
        with col3:
            tecnico_sel = st.selectbox("3. Técnico Responsável:", ["Técnico Junior", "Técnica Simone"])

        t_os = TEMPLATE_OS_JUNIOR if tecnico_sel == "Técnico Junior" else TEMPLATE_OS_SIMONE
        t_nr = TEMPLATE_NR06_JUNIOR if tecnico_sel == "Técnico Junior" else TEMPLATE_NR06_SIMONE

        st.markdown("<br>", unsafe_allow_html=True)
        g1, g2, g3 = st.columns(3)
        g_os, g_ficha, g_cert = g1.checkbox("OS", True), g2.checkbox("Ficha EPI", True), g3.checkbox("Certificado", True)

        if st.button("🚀 PROCESSAR DOCUMENTOS"):
            with st.spinner("Gerando documentos..."):
                cargo = str(dados_colab['Cargo']).strip()
                df_cargos['f_l'] = df_cargos['Função'].astype(str).apply(remover_acentos)
                cargo_norm = remover_acentos(cargo)
                desc_f = df_cargos[df_cargos['f_l'] == cargo_norm]

                if not desc_f.empty:
                    desc_atv = desc_f['Descrição da Atividade'].values[0]
                    arquivos = {}

                    if g_os:
                        doc = Document(t_os)
                        substituir_docx(doc, {"{{NOME}}": nome_sel, "{{FUNCAO}}": cargo.upper(), "{{CNPJ}}": UNIDADES_DOCS[unid_sel]["CNPJ"], "{{ENDERECO}}": UNIDADES_DOCS[unid_sel]["ENDERECO"], "{{SETOR}}": str(dados_colab.get('NomeLocal', '')), "{{DESCRICAO_ATIVIDADE}}": str(desc_atv), "{{DATA}}": datetime.now().strftime("%d/%m/%Y")})
                        b = io.BytesIO(); doc.save(b); arquivos[f"OS {nome_sel}.docx"] = b.getvalue()

                    if g_ficha:
                        df_e = carregar_aba(cargo, SHEET_ID_DOCS)
                        if df_e.empty: df_e = carregar_aba(remover_acentos(cargo), SHEET_ID_DOCS)
                        if not df_e.empty:
                            m_f = {"{{NOME}}": nome_sel, "{{MATRICULA}}": formatar_matricula(dados_colab.get('Matrícula', '')), "{{FUNCAO}}": cargo, "{{DATA_ADMISSAO}}": datetime.now().strftime("%d/%m/%Y"), "{{SETOR}}": str(dados_colab.get('NomeLocal', ''))}
                            arquivos[f"Ficha EPI {nome_sel}.xlsx"] = preencher_excel_ficha(TEMPLATE_FICHA, m_f, df_e)

                    if g_cert:
                        prs = Presentation(t_nr)
                        substituir_pptx(prs, {"{{NOME}}": nome_sel, "{{CPF}}": formatar_cpf(dados_colab.get('CPF', '')), "{{FUNCAO}}": cargo, "{{DATA_TREINAMENTO}}": datetime.now().strftime("%d/%m/%Y"), "{{LOCAL_DATA}}": f"{unid_sel.title()}, {data_extenso_pt()}."})
                        b = io.BytesIO(); prs.save(b); arquivos[f"NR06 {nome_sel}.pptx"] = b.getvalue()

                    if arquivos:
                        z_b = io.BytesIO()
                        with zipfile.ZipFile(z_b, "w") as z:
                            for n, d in arquivos.items(): z.writestr(n, d)
                        st.success("✅ Documentos prontos!")
                        st.download_button("📦 BAIXAR KIT COMPLETO (ZIP)", z_b.getvalue(), f"Kit_{nome_sel}.zip", use_container_width=True)

# ==========================================
# MÓDULO 2: CONTROLE DE ASO (NOVO)
# ==========================================
else:
    st.subheader("🛡️ Controle de Vencimento de ASO")
    try:
        aba_nome = st.sidebar.selectbox("Selecione a Unidade ASO", list(UNIDADES_ASO.keys()))
        df_aso = carregar_aba(aba_nome, SHEET_ID_ASO)

        if not df_aso.empty:
            hoje = datetime.now()
            df_aso['Venc'] = pd.to_datetime(df_aso['Venc'], dayfirst=True, errors='coerce')
            alertas = df_aso[df_aso['Venc'] <= (hoje + timedelta(days=10))].copy()
            
            st.metric("ASOs em Alerta (10 dias)", len(alertas))
            st.subheader(f"⚠️ Colaboradores em Alerta - {aba_nome}")
            st.dataframe(alertas[['Nome', 'Cargo', 'Setor', 'Venc']], use_container_width=True)

            st.markdown("---")
            st.subheader("📝 Gerar Solicitação de Agendamento")
            sc1, sc2, sc3 = st.columns(3)
            with sc1:
                n_sel_aso = st.selectbox("Colaborador", alertas['Nome'].tolist())
            with sc2:
                t_ag_aso = st.selectbox("Tipo de Agendamento", ["PERIÓDICO", "MUDANÇA DE RISCO", "RETORNO AO TRABALHO"])
            with sc3:
                dt_s_aso = st.date_input("Data de Sugestão", value=hoje + timedelta(days=2))

            if n_sel_aso:
                dados_aso = alertas[alertas['Nome'] == n_sel_aso].iloc[0]
                doc_aso = gerar_docx_aso(dados_aso, t_ag_aso, dt_s_aso, aba_nome)
                st.download_button(label=f"📥 Baixar Solicitação - {n_sel_aso}", data=doc_aso, file_name=f"ASO_{n_sel_aso.replace(' ', '_')}.docx")

    except Exception as e:
        st.error(f"Erro no sistema ASO: {e}")

# RODAPÉ
st.markdown("""<div class="footer">© 2026 Gestão Documentos | Versão 1.0 | Desenvolvido por: Dilceu Junior</div>""", unsafe_allow_html=True)
