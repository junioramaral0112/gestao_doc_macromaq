import streamlit as st
import pandas as pd
from docx import Document
from docx.shared import Inches, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from pptx import Presentation
import openpyxl
import io
import os
import zipfile
import unicodedata
from datetime import datetime, timedelta
from urllib.parse import quote
import base64

# --- CONFIGURAÇÕES GLOBAIS ---
st.set_page_config(page_title="Portal Integrado SSMA - Junior 360", layout="wide")

# LÓGICA DE CAMINHO DINÂMICO (ORIGINAL)
if os.path.exists(r"C:\Users\dilceu.gomes\Documents"):
    BASE_PATH = r"C:\Users\dilceu.gomes\Documents"
else:
    BASE_PATH = os.path.dirname(__file__) if "__file__" in locals() else "."

# CAMINHOS DOS ARQUIVOS (MANTIDOS)
FUNDO_PATH = os.path.join(BASE_PATH, "fundo.png")
LOGO_PATH = os.path.join(BASE_PATH, "logo.png")
LOGO_CLINICA = r"C:\Users\dilceu.gomes\Desktop\sistema_aso\adivitta.png"
TEMPLATE_FICHA = os.path.join(BASE_PATH, "template_ficha.xlsx")
TEMPLATE_OS_JUNIOR = os.path.join(BASE_PATH, "template_os_Junior.docx")
TEMPLATE_NR06_JUNIOR = os.path.join(BASE_PATH, "template_nr06_Junior.pptx")
TEMPLATE_OS_SIMONE = os.path.join(BASE_PATH, "template_os_simone.docx")
TEMPLATE_NR06_SIMONE = os.path.join(BASE_PATH, "template_nr06_simone.pptx")

# PLANILHAS
SHEET_ID_DOCS = "1y98U3eK7JXJqQaMC0i7eFbwpvp97Nuyeml5Dis0UCUg"
SHEET_ID_ASO = "1G_oVT9gK-n_jGh5R4g65qUwK_MfQGvCX-SA4NHNNflU"

# --- DICIONÁRIOS DE UNIDADES (MANTIDOS SEPARADOS PARA EVITAR NAMEERROR) ---
UNIDADES_DOCS = {
    "SÃO JOSÉ": {"CNPJ": "83.675.413/0001-01", "ENDERECO": "BR 101, km 210 / Bairro: Picadas do Sul – São José – SC / CEP: 88106-100"},
    "CHAPECÓ": {"CNPJ": "83.675.413/0002-84", "ENDERECO": "Rua Xanxerê, 360E – Bairro Líder – Chapecó/SC"},
    "JOINVILLE": {"CNPJ": "83.675.413/0011-75", "ENDERECO": "BR101, KM17 – Sentido Norte – Bairro Sta Catarina – Joinville / SC"},
    "PARANÁ": {"CNPJ": "83.675.413/0004-46", "ENDERECO": "Av. Juscelino K. de Oliveira, 3628 – Bairro CIC – Curitiba / PR"},
    "SÃO PAULO": {"CNPJ": "83.675.413/0008-70", "ENDERECO": "Rua Goiabeira 105/125 – Bairro Roseira de Cima – Jaguariúna / SP"},
    "MINAS GERAIS": {"CNPJ": "83.675.413/0014-18", "ENDERECO": "Anel Rodoviário Celso Mello Azevedo, 3713 - Bom Sucesso - BH/MG"},
    "ITAJAÍ": {"CNPJ": "83.675.413/0013-37", "ENDERECO": "Av. Vereador Abrahão João Francisco, 2300 - Dom Bosco - Itajaí / SC"}
}

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
        .stSelectbox label, div[data-testid="stCheckbox"] label p {{ color: white !important; background: rgba(0,0,0,0.7); padding: 5px 12px; border-radius: 8px; font-weight: bold; }}
        .stButton > button {{ background: #2c3e50; color: #f9cc0b; border: 2px solid #f9cc0b; border-radius: 10px; height: 55px; font-weight: bold; width: 100%; font-size: 18px; }}
        .header-container {{ display: flex; align-items: center; background: rgba(255,255,255,0.9); padding: 20px; border-radius: 15px; margin-bottom: 30px; }}
        .footer {{ position: fixed; left: 0; bottom: 0; width: 100%; background: rgba(0,0,0,0.8); color: white; text-align: center; padding: 10px; font-size: 13px; z-index: 999; }}
        </style>
        <div class="header-container">
            <img src="data:image/png;base64,{logo}" width="320">
            <h1 style="margin-left:25px;color:#2c3e50;">Gestão SSMA Macromaq</h1>
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

# --- FUNÇÕES DE DOCUMENTOS (ORIGINAIS) ---
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
                for p in shape.text_frame.paragraphs:
                    for run in p.runs:
                        for tag, valor in mapeamento.items():
                            if tag in run.text: run.text = run.text.replace(tag, str(valor))

def formatar_matricula(v):
    if pd.isna(v) or v == "": return ""
    try: return str(int(float(v)))
    except: return str(v)

def formatar_cpf(cpf):
    cpf = ''.join(filter(str.isdigit, str(cpf))).zfill(11)
    return f"{cpf[:3]}.{cpf[3:6]}.{cpf[6:9]}-{cpf[9:]}"

def limpar_valor(v):
    if pd.isna(v): return ""
    t = str(v).strip()
    return "" if t.lower() in ["nan", "na"] else t

def data_extenso_pt():
    meses = {1: "Janeiro", 2: "Fevereiro", 3: "Março", 4: "Abril", 5: "Maio", 6: "Junho", 7: "Julho", 8: "Agosto", 9: "Setembro", 10: "Outubro", 11: "Novembro", 12: "Dezembro"}
    agora = datetime.now()
    return agora.strftime(f"%d de {meses[agora.month]} de %Y")

# --- FUNÇÕES ASO (ORIGINAIS DO NOSSO CONTROLE) ---
def gerar_aso_docx(dados, tipo, data_sugestao, unidade):
    doc = Document()
    if os.path.exists(LOGO_CLINICA):
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.add_run().add_picture(LOGO_CLINICA, width=Inches(1.5))
    
    run_titulo = doc.add_paragraph().add_run(f"FORMULÁRIO PARA AGENDAMENTO {tipo}")
    run_titulo.bold, run_titulo.font.size = True, Pt(14)

    table = doc.add_table(rows=7, cols=2)
    table.style = 'Table Grid'
    labels = ["Nome Completo", "Cargo", "Setor", "Unidade", "Cliente", "Local", "Data Sugestão"]
    valores = [str(dados['Nome']), str(dados['Cargo']), str(dados['Setor']), unidade, "MACROMAQ", "Arapoti", data_sugestao.strftime('%d/%m/%Y')]
    
    for i, (l, v) in enumerate(zip(labels, valores)):
        table.rows[i].cells[0].text = l
        table.rows[i].cells[1].text = v
        table.rows[i].cells[0].paragraphs[0].runs[0].bold = True

    target = BytesIO()
    doc.save(target)
    return target.getvalue()

# --- EXECUÇÃO DO PORTAL ---
aplicar_layout()

st.sidebar.markdown("### 🗂️ Navegação")
menu_modulo = st.sidebar.radio("Selecione o Módulo:", ["Emissão de Documentos", "Controle de ASO"])

# ==========================================
# MÓDULO 1: EMISSÃO DE DOCUMENTOS (MACROMAQ)
# ==========================================
if menu_modulo == "Emissão de Documentos":
    df_colab = carregar_aba("Colaboradores", SHEET_ID_DOCS)
    df_cargos = carregar_aba("Cargos", SHEET_ID_DOCS)

    if not df_colab.empty and not df_cargos.empty:
        c1, c2, c3 = st.columns(3)
        with c1:
            nome_sel = st.selectbox("1. Selecione o Colaborador:", df_colab['Nome Colaborador'].dropna().unique())
            dados_colab = df_colab[df_colab['Nome Colaborador'] == nome_sel].iloc[0]
        with c2:
            unidade_p = str(dados_colab.get('Filial', dados_colab.get('Unidade', ''))).upper().strip()
            lista_u = list(UNIDADES_DOCS.keys())
            idx = lista_u.index(unidade_p) if unidade_p in lista_u else 0
            unid_sel = st.selectbox("2. Unidade para OS:", lista_u, index=idx)
        with c3:
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
                desc_f = df_cargos[df_cargos['f_l'] == remover_acentos(cargo)]

                if not desc_f.empty:
                    desc_atv, arquivos = desc_f['Descrição da Atividade'].values[0], {}
                    
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
# MÓDULO 2: CONTROLE DE ASO (JUNIOR 360)
# ==========================================
else:
    aba_aso = st.sidebar.selectbox("Selecione a Unidade ASO", list(UNIDADES_ASO.keys()))
    df_aso = carregar_aba(aba_aso, SHEET_ID_ASO)

    if not df_aso.empty:
        hoje = datetime.now()
        df_aso['Venc'] = pd.to_datetime(df_aso['Venc'], dayfirst=True, errors='coerce')
        alertas = df_aso[df_aso['Venc'] <= (hoje + timedelta(days=10))].copy()
        
        st.metric("ASOs em Alerta (10 dias)", len(alertas))
        st.subheader(f"⚠️ Colaboradores em Alerta - {aba_aso}")
        st.dataframe(alertas[['Nome', 'Cargo', 'Setor', 'Venc']], use_container_width=True)

        st.markdown("---")
        st.subheader("📝 Gerar Solicitação de Agendamento")
        sc1, sc2, sc3 = st.columns(3)
        with sc1:
            n_sel = st.selectbox("Colaborador", alertas['Nome'].tolist())
        with sc2:
            t_ag = st.selectbox("Tipo", ["PERIÓDICO", "MUDANÇA DE RISCO", "RETORNO AO TRABALHO"])
        with sc3:
            dt_s = st.date_input("Data de Sugestão", value=hoje + timedelta(days=2))

        if n_sel:
            d_colab = alertas[alertas['Nome'] == n_sel].iloc[0]
            doc_aso = gerar_aso_docx(d_colab, t_ag, dt_s, aba_aso)
            st.download_button(label=f"📥 Baixar Solicitação - {n_sel}", data=doc_aso, file_name=f"ASO_{n_sel.replace(' ', '_')}.docx")

st.markdown("""<div class="footer">© 2026 Gestão Documentos | Versão 1.0 | Desenvolvido por: Dilceu Junior</div>""", unsafe_allow_html=True)
