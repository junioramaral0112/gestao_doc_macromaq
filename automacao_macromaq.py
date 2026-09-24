from __future__ import annotations
import base64
import hashlib
import json
import os
from datetime import datetime
import requests
import streamlit as st

# ==========================================================================
# CONFIGURAÇÃO DA PÁGINA
# ==========================================================================
st.set_page_config(
    page_title="Portal SSMA Macromaq",
    page_icon="🛡️",
    layout="wide",
    initial_sidebar_state="collapsed",
)

# ==========================================================================
# CONSTANTES
# ==========================================================================
_BASE_DIR = os.path.dirname(os.path.abspath(__file__))
_DATA_DIR = os.path.join(_BASE_DIR, "data")
_USERS_FILE = os.path.join(_DATA_DIR, "usuarios.json")
_IMAGES_DIR = os.path.join(_BASE_DIR, "assets", "images")
_FUNDO_JPG = os.path.join(_IMAGES_DIR, "fundo.jpg")
_LOGO_PNG = os.path.join(_IMAGES_DIR, "logo.png")

DEFAULT_USER = "macrossma"
DEFAULT_PASS = "macromaq2026"
TRIAL_DIAS = 15

# Templates por Técnico
TEMPLATE_FICHA = os.path.join(BASE_PATH, "template_ficha.docx")
TEMPLATE_OS_DAIANE = os.path.join(BASE_PATH, "template_os_Daiane.docx")
TEMPLATE_NR06_DAIANE = os.path.join(BASE_PATH, "template_nr06_Daiane.pptx")
TEMPLATE_OS_SIMONE = os.path.join(BASE_PATH, "template_os_simone.docx")
TEMPLATE_NR06_SIMONE = os.path.join(BASE_PATH, "template_nr06_simone.pptx")
TEMPLATE_OS_JUNIOR = os.path.join(BASE_PATH, "template_os_Junior.docx")
TEMPLATE_NR06_JUNIOR = os.path.join(BASE_PATH, "template_nr06_Junior.pptx")

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
    "SÃO LEOPOLDO": {
        "CNPJ": "83.675.413/0016-80",
        "ENDERECO": "Avenida Senador Salgado Filho, 1970 – São Leopoldo – RS"
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

# ==========================================================================
# FUNÇÕES DE USUÁRIO
# ==========================================================================
def _hash_senha(senha: str) -> str: 
    return hashlib.sha256(senha.encode("utf-8")).hexdigest()

def _carregar_usuarios() -> dict:
    os.makedirs(_DATA_DIR, exist_ok=True)
    default_data = {
        DEFAULT_USER: {
            "senha_hash": _hash_senha(DEFAULT_PASS),
            "senha_padrao": False,
            "criado_em": datetime.now().isoformat(),
        }
    }
    
    if not os.path.exists(_USERS_FILE):
        _salvar_usuarios(default_data)
        return default_data
    
    try:
        with open(_USERS_FILE, "r", encoding="utf-8") as f: 
            data = json.load(f)
            if DEFAULT_USER not in data:
                data.update(default_data)
                _salvar_usuarios(data)
            else:
                data[DEFAULT_USER]["senha_padrao"] = False
                _salvar_usuarios(data)
            return data
    except:
        _salvar_usuarios(default_data)
        return default_data

def _salvar_usuarios(data: dict) -> None:
    os.makedirs(_DATA_DIR, exist_ok=True)
    with open(_USERS_FILE, "w", encoding="utf-8") as f: 
        json.dump(data, f, ensure_ascii=False, indent=2)

def _verificar_senha(usuario: str, senha: str) -> bool:
    usuarios = _carregar_usuarios()
    user_data = usuarios.get(usuario)
    return user_data["senha_hash"] == _hash_senha(senha) if user_data else False

def _dias_desde_cadastro(usuario: str) -> int:
    usuarios = _carregar_usuarios()
    criado_str = usuarios.get(usuario, {}).get("criado_em", "")
    return (datetime.now() - datetime.fromisoformat(criado_str)).days if criado_str else 999

# ==========================================================================
# INTEGRAÇÃO ASAAS
# ==========================================================================
@st.cache_data(ttl=600)
def verificar_adimplencia() -> dict:
    ASAAS_API_KEY = st.secrets.get("ASAAS_API_KEY", os.getenv("ASAAS_API_KEY", ""))
    ASAAS_BASE_URL = "https://api.asaas.com/v3"
    if not ASAAS_API_KEY: return {"adimplente": False, "status": "ERRO", "mensagem": "API não configurada."}
    try:
        resp = requests.get(f"{ASAAS_BASE_URL}/payments", headers={"access_token": ASAAS_API_KEY}, timeout=15)
        data = resp.json().get("data", [])
        return {"adimplente": any(p.get("status") in ["RECEIVED", "CONFIRMED"] for p in data)}
    except: return {"adimplente": False, "mensagem": "Erro na verificação financeira."}

# ==========================================================================
# CSS E LAYOUT
# ==========================================================================
def _para_base64(caminho: str) -> str:
    if not os.path.exists(caminho): return ""
    with open(caminho, "rb") as f: return base64.b64encode(f.read()).decode()

def _injetar_css():
    b64_fundo = _para_base64(_FUNDO_JPG)
    if b64_fundo:
        fundo_css = f"background: linear-gradient(rgba(15,23,42,0.88), rgba(15,23,42,0.92)), url(data:image/jpeg;base64,{b64_fundo}); background-size: cover; background-position: center; background-attachment: fixed;"
    else:
        fundo_css = "background: linear-gradient(135deg, #0f172a, #1e293b);"

    st.markdown(f"""
        <style>
        [data-testid="stSidebar"], [data-testid="stSidebarNav"], [data-testid="collapsedControl"] {{ display: none !important; }}
        header[data-testid="stHeader"] {{ display: none !important; }}
        .stApp {{ {fundo_css} }}

        .login-container {{
            display: flex;
            justify-content: center;
            align-items: center;
            height: 80vh;
        }}
        .login-box {{
            width: 100%;
            max-width: 420px;
            background: rgba(255,255,255,0.96);
            border-radius: 16px;
            padding: 36px 32px;
            box-shadow: 0 8px 32px rgba(0,0,0,0.25);
            text-align: center;
        }}
        .login-box h2 {{ color: #1e293b; margin-bottom: 6px; font-weight: 800; }}

        .portal-header {{ display: flex; align-items: center; gap: 20px; background: rgba(255,255,255,0.95); backdrop-filter: blur(10px); border-radius: 16px; padding: 18px 24px; margin-bottom: 28px; box-shadow: 0 4px 24px rgba(0,0,0,0.2); }}
        .portal-header h1 {{ font-size: 1.6rem; font-weight: 800; color: #1e293b; margin: 0; }}
        .card {{ background: rgba(255,255,255,0.95); backdrop-filter: blur(8px); border-radius: 16px; padding: 26px 18px; text-align: center; box-shadow: 0 4px 20px rgba(0,0,0,0.15); height: 100%; display: flex; flex-direction: column; justify-content: space-between; }}
        .card-icon {{ font-size: 2.8rem; margin-bottom: 10px; }}
        .card-btn {{ display: inline-block; padding: 10px 20px; border-radius: 10px; font-weight: 700; text-decoration: none; }}
        .btn-aso {{ background: linear-gradient(135deg,#3b82f6,#2563eb); color: white; }}
        .btn-docs {{ background: linear-gradient(135deg,#10b981,#059669); color: white; }}
        .btn-apr {{ background: linear-gradient(135deg,#f59e0b,#d97706); color: #1e293b; }}
        .btn-audit {{ background: linear-gradient(135deg,#8b5cf6,#6d28d9); color: white; }}
        .portal-footer {{ position: fixed; bottom: 0; left: 0; right: 0; text-align: center; padding: 12px; background: rgba(15,23,42,0.95); color: #94a3b8; font-size: 0.75rem; z-index: 100; }}
        </style>
    """, unsafe_allow_html=True)

# --- FUNÇÕES AUXILIARES DE PROCESSAMENTO ---
def remover_acentos(texto):
    if not isinstance(texto, str): return str(texto)
    return "".join(c for c in unicodedata.normalize('NFD', texto.strip()) if unicodedata.category(c) != 'Mn').lower()

def limpar_valor(valor):
    if pd.isna(valor): return ""
    return str(valor).strip()

def limpar_quebras_linha(texto):
    if not isinstance(texto, str): texto = str(texto)
    texto = texto.replace('\r', ' ').replace('\n', ' ')
    while "  " in texto:
        texto = texto.replace("  ", " ")
    return texto.strip()

@st.cache_data(ttl=300)
def carregar_aba(aba_nome):
    try:
        url = f"https://docs.google.com/spreadsheets/d/{SHEET_ID}/gviz/tq?tqx=out:csv&sheet={quote(aba_nome.strip())}"
        df = pd.read_csv(url, dtype=str)
        return df
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
    try:
        if pd.isna(cpf): return "Não informado"
        cpf = str(cpf).strip()
        if cpf.endswith('.0'): cpf = cpf[:-2]
        if cpf.lower() in ["", "nan", "none", "0"]: return "Não informado"
        cpf = ''.join(filter(str.isdigit, cpf))
        if len(cpf) == 10: cpf = "0" + cpf
        if len(cpf) == 11: return f"{cpf[:3]}.{cpf[3:6]}.{cpf[6:9]}-{cpf[9:]}"
        return cpf if cpf != "" else "Não informado"
    except:
        return "Não informado"

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
        st.sidebar.warning(f"Aviso técnico: Falha ao converter {nome_original} para PDF. Detalhes: {e}")
    return None, None

def substituir_docx(doc, mapeamento):
    for p in doc.paragraphs:
        for tag, valor in mapeamento.items():
            if tag in p.text:
                for run in p.runs:
                    if tag in run.text:
                        run.text = run.text.replace(tag, str(valor))
                if tag in p.text:
                    texto_inteiro = p.text
                    for t, v in mapeamento.items():
                        texto_inteiro = texto_inteiro.replace(t, str(v))
                    if p.runs:
                        p.runs[0].text = texto_inteiro
                        for run in p.runs[1:]:
                            run.text = ""

    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    for tag, valor in mapeamento.items():
                        if tag in p.text:
                            for run in p.runs:
                                if tag in run.text:
                                    run.text = run.text.replace(tag, str(valor))
                            if tag in p.text:
                                texto_inteiro = p.text
                                for t, v in mapeamento.items():
                                    texto_inteiro = texto_inteiro.replace(t, str(v))
                                if p.runs:
                                    p.runs[0].text = texto_inteiro
                                    for run in p.runs[1:]:
                                        run.text = ""

def substituir_pptx(prs, mapeamento):
    for slide in prs.slides:
        for shape in slide.shapes:
            if hasattr(shape, "text_frame"):
                for paragraph in shape.text_frame.paragraphs:
                    for run in paragraph.runs:
                        for tag, valor in mapeamento.items():
                            if tag in run.text: run.text = run.text.replace(tag, str(valor))

def preencher_ficha_docx(caminho_template, mapeamento, df_epis):
    doc = Document(caminho_template)
    substituir_docx(doc, mapeamento)
    
    tabela_alvo = None
    linhas_tags = []
    
    for tabela in doc.tables:
        for row in tabela.rows:
            texto_linha = "".join(cell.text for cell in row.cells)
            if "ITEM" in texto_linha or "DESC" in texto_linha:
                tabela_alvo = tabela
                linhas_tags.append(row)
                    
    if not tabela_alvo or len(linhas_tags) == 0:
        raise Exception("Nenhuma linha contendo a tag {{ITEM}} foi localizada no template do Word.")
        
    qtd_items = len(df_epis)
    linha_modelo = linhas_tags[0]
    
    def atualizar_celula_preservando_estilo(celula, novo_texto, alinhamento=WD_ALIGN_PARAGRAPH.CENTER):
        if not celula.paragraphs:
            celula.add_paragraph()
        p = celula.paragraphs[0]
        p.alignment = alinhamento
        if p.runs:
            p.runs[0].text = novo_texto
            for r in p.runs[1:]:
                r.text = ""
        else:
            p.add_run(novo_texto)

    for i, row_item in enumerate(linhas_tags):
        if i < qtd_items:
            item = df_epis.iloc[i]
            num_seq = f"{i + 1:02d}"
            
            ca_valor = limpar_valor(item.get('C.A.', ''))
            if ca_valor == "" or ca_valor.lower() in ["nan", "none", "0", "0.0"]:
                ca_valor = "N/A"
            
            for cell in row_item.cells:
                texto_celula = cell.text
                if "ITEM" in texto_celula: 
                    atualizar_celula_preservando_estilo(cell, num_seq, WD_ALIGN_PARAGRAPH.CENTER)
                elif "DESC" in texto_celula: 
                    atualizar_celula_preservando_estilo(cell, limpar_valor(item.get('Descrição', '')), WD_ALIGN_PARAGRAPH.LEFT)
                elif "CA" in texto_celula: 
                    atualizar_celula_preservando_estilo(cell, ca_valor, WD_ALIGN_PARAGRAPH.CENTER)
                elif "QT" in texto_celula: 
                    atualizar_celula_preservando_estilo(cell, limpar_valor(item.get('qt.', '')), WD_ALIGN_PARAGRAPH.CENTER)
                elif "unid" in texto_celula: 
                    atualizar_celula_preservando_estilo(cell, limpar_valor(item.get('unid.', 'unid')), WD_ALIGN_PARAGRAPH.CENTER)
                elif "DATA" in texto_celula: 
                    atualizar_celula_preservando_estilo(cell, datetime.now().strftime("%d/%m/%Y"), WD_ALIGN_PARAGRAPH.CENTER)
        else:
            for cell in row_item.cells:
                for paragraph in cell.paragraphs:
                    paragraph.text = ""

    if qtd_items > len(linhas_tags):
        tr_modelo = linha_modelo._tr
        for i in range(len(linhas_tags), qtd_items):
            item = df_epis.iloc[i]
            num_seq = f"{i + 1:02d}"
            
            ca_valor = limpar_valor(item.get('C.A.', ''))
            if ca_valor == "" or ca_valor.lower() in ["nan", "none", "0", "0.0"]:
                ca_valor = "N/A"
            
            nova_tr = copy.deepcopy(tr_modelo)
            nova_linha = tabela_alvo.add_row()
            nova_linha._tr.getparent().replace(nova_linha._tr, nova_tr)
            nova_linha._tr = nova_tr
            
            for cell in nova_linha.cells:
                texto_celula = cell.text
                if "ITEM" in texto_celula: 
                    atualizar_celula_preservando_estilo(cell, num_seq, WD_ALIGN_PARAGRAPH.CENTER)
                elif "DESC" in texto_celula: 
                    atualizar_celula_preservando_estilo(cell, limpar_valor(item.get('Descrição', '')), WD_ALIGN_PARAGRAPH.LEFT)
                elif "CA" in texto_celula: 
                    atualizar_celula_preservando_estilo(cell, ca_valor, WD_ALIGN_PARAGRAPH.CENTER)
                elif "QT" in texto_celula: 
                    atualizar_celula_preservando_estilo(cell, limpar_valor(item.get('qt.', '')), WD_ALIGN_PARAGRAPH.CENTER)
                elif "unid" in texto_celula: 
                    atualizar_celula_preservando_estilo(cell, limpar_valor(item.get('unid.', 'unid')), WD_ALIGN_PARAGRAPH.CENTER)
                elif "DATA" in texto_celula: 
                    atualizar_celula_preservando_estilo(cell, datetime.now().strftime("%d/%m/%Y"), WD_ALIGN_PARAGRAPH.CENTER)
            
    output = io.BytesIO()
    doc.save(output)
    return output.getvalue()

# ==========================================================================
# TELAS E FLUXO PRINCIPAL
# ==========================================================================
def tela_login():
    _injetar_css()
    
    b64_logo = _para_base64(_LOGO_PNG)
    logo_tag = f'<img src="data:image/png;base64,{b64_logo}" style="height:55px; margin-bottom:12px;" alt="Logo">' if b64_logo else '<span style="font-size:2.5rem;">🛡️</span>'

    st.markdown(f"""
        <div class="login-container">
            <div class="login-box">
                {logo_tag}
                <h2>Portal Central - SSMA</h2>
                <p style="color:#64748b; font-size:0.85rem; margin-bottom:20px;">Gestão Integrada de Segurança, Saúde e Meio Ambiente</p>
    """, unsafe_allow_html=True)

    with st.form("form_login"):
        user = st.text_input("Usuário", placeholder="Digite seu usuário")
        pw = st.text_input("Senha", type="password", placeholder="Digite sua senha")
        submitted = st.form_submit_button("🔐 Entrar", use_container_width=True, type="primary")

        if submitted:
            if _verificar_senha(user.strip(), pw):
                st.session_state.autenticado = True
                st.session_state.usuario = user.strip()
                st.rerun()
            else:
                st.error("Usuário ou senha inválidos.")

    st.markdown('</div></div>', unsafe_allow_html=True)

def tela_portal():
    _injetar_css()
    b64_logo = _para_base64(_LOGO_PNG)
    logo_tag = f'<img src="data:image/png;base64,{b64_logo}" style="height:45px; width:auto;" alt="Logo">' if b64_logo else '<span style="font-size:2rem;">🛡️</span>'
    usuario = st.session_state.get("usuario", "")

    st.markdown(f"""
        <div class="portal-header">
            {logo_tag}
            <div style="flex-grow:1;">
                <h1>Portal Central - SSMA</h1>
                <p style="color:#64748b; font-size:0.8rem; margin:2px 0 0;">Gestão Integrada de Segurança, Saúde e Meio Ambiente</p>
            </div>
            <div style="text-align:right;">
                <span style="color:#64748b; font-size:0.75rem;">👤 {usuario}</span>
            </div>
        </div>
    """, unsafe_allow_html=True)

    col_logout = st.columns([10, 1])
    with col_logout[1]:
        if st.button("🚪 Sair", use_container_width=True):
            st.session_state.clear()
            st.rerun()

    df_colab = carregar_aba("Colaboradores")
    df_cargos = carregar_aba("Cargos")

    if not df_colab.empty and not df_cargos.empty:
        df_colab['Nome_Formatado'] = df_colab['Nome Colaborador'].astype(str).str.strip().str.title()
        
        colab_query = st.query_params.get("colaborador", None)
        lista_nomes = sorted(df_colab['Nome_Formatado'].dropna().unique())
        idx_padrao = 0

        if colab_query:
            nome_query_limpo = remover_acentos(unquote(str(colab_query)))
            for i, nome in enumerate(lista_nomes):
                if remover_acentos(nome) == nome_query_limpo or nome_query_limpo in remover_acentos(nome):
                    idx_padrao = i
                    break

        col1, col2, col3 = st.columns(3)
        with col1:
            nome_sel = st.selectbox("1. Selecione o Colaborador:", lista_nomes, index=idx_padrao)
            dados_colab = df_colab[df_colab['Nome_Formatado'] == nome_sel].iloc[0]
            
        with col2:
            unidade_plan = str(dados_colab.get('Filial', dados_colab.get('Unidade', ''))).upper().strip()
            lista_unid = list(UNIDADES.keys())
            idx = lista_unid.index(unidade_plan) if unidade_plan in lista_unid else 0
            unid_sel = st.selectbox("2. Unidade para OS:", lista_unid, index=idx)
        with col3:
            tecnico_sel = st.selectbox("3. Técnico Responsável:", ["Técnica Daiane Sales", "Técnica Simone", "Técnico Dilceu Junior"])

        # Seleção dos templates com base no técnico escolhido
        if tecnico_sel == "Técnica Daiane Sales":
            t_os = TEMPLATE_OS_DAIANE
            t_nr = TEMPLATE_NR06_DAIANE
        elif tecnico_sel == "Técnica Simone":
            t_os = TEMPLATE_OS_SIMONE
            t_nr = TEMPLATE_NR06_SIMONE
        else:
            t_os = TEMPLATE_OS_JUNIOR
            t_nr = TEMPLATE_NR06_JUNIOR

        st.markdown("<br>", unsafe_allow_html=True)
        c1, c2, c3 = st.columns(3)
        g_os, g_ficha, g_cert = c1.checkbox("OS", True), c2.checkbox("Ficha EPI", True), c3.checkbox("Certificado", True)

        incluir_pdf = st.checkbox("📄 Incluir cópias em formato PDF no Kit", True)

        if st.button("🚀 PROCESSAR DOCUMENTOS"):
            with st.spinner("Gerando documentos..."):
                cargo = str(dados_colab['Cargo']).strip()
                arquivos = {}
                df_cargos['f_l'] = df_cargos['Função'].astype(str).apply(remover_acentos)
                desc_f = df_cargos[df_cargos['f_l'] == remover_acentos(cargo)]

                if not desc_f.empty:
                    desc_atv = limpar_quebras_linha(desc_f['Descrição da Atividade'].fillna('').values[0]) if 'Descrição da Atividade' in desc_f.columns else ''
                    riscos_agentes = limpar_quebras_linha(desc_f['Riscos e Agentes Existentes'].fillna('').values[0]) if 'Riscos e Agentes Existentes' in desc_f.columns else ''
                    medidas_protecao = limpar_quebras_linha(desc_f['Medidas de Proteção'].fillna('').values[0]) if 'Medidas de Proteção' in desc_f.columns else ''
                    
                    setor_original = limpar_valor(dados_colab.get('NomeLocal', dados_colab.get('Setor', '')))
                    setor_final = setor_original if setor_original != "" else unid_sel.title()
                    
                    coluna_cpf = [col for col in df_colab.columns if 'CPF' in col.upper()]
                    if coluna_cpf:
                        cpf_bruto = dados_colab[coluna_cpf[0]]
                    else:
                        try: cpf_bruto = dados_colab.iloc[18]
                        except: cpf_bruto = ""
                            
                    cpf_final = formatar_cpf(cpf_bruto)
                    
                    # 1. Ordem de Serviço
                    if g_os:
                        doc = Document(t_os)
                        substituir_docx(doc, {
                            "{{NOME}}": dados_colab['Nome Colaborador'], 
                            "{{FUNCAO}}": cargo.upper(), 
                            "{{CNPJ}}": UNIDADES[unid_sel]["CNPJ"], 
                            "{{ENDERECO}}": UNIDADES[unid_sel]["ENDERECO"], 
                            "{{SETOR}}": setor_final, 
                            "{{DESCRICAO_ATIVIDADE}}": desc_atv, 
                            "{{MEDIDAS_PROTECAO}}": medidas_protecao,    
                            "{{RISCOS_AGENTES}}": riscos_agentes,
                            "{{DATA}}": datetime.now().strftime("%d/%m/%Y")
                        })
                        b = io.BytesIO(); doc.save(b)
                        conteudo_docx = b.getvalue()
                        nome_docx = f"OS {nome_sel}.docx"
                        arquivos[nome_docx] = conteudo_docx
                        
                        if incluir_pdf:
                            pdf_bytes, nome_pdf = converter_para_pdf_linux(conteudo_docx, nome_docx)
                            if pdf_bytes: arquivos[nome_pdf] = pdf_bytes

                    # 2. Ficha de EPI
                    if g_ficha:
                        cargo_limpo = cargo.strip()
                        df_e = carregar_aba(cargo_limpo)
                        
                        if df_e.empty: df_e = carregar_aba(f"{cargo_limpo} ")
                        if df_e.empty: df_e = carregar_aba(cargo_limpo.title())
                        if df_e.empty: df_e = carregar_aba(remover_acentos(cargo_limpo))
                        if df_e.empty and "jr" in cargo_limpo.lower():
                            df_e = carregar_aba(cargo_limpo.lower().replace("jr", "Jr"))

                        if not df_e.empty:
                            m_f = {"{{NOME}}": dados_colab['Nome Colaborador'], "{{MATRICULA}}": formatar_matricula(dados_colab.get('Matrícula', '')), "{{FUNCAO}}": cargo, "{{DATA_ADMISSAO}}": datetime.now().strftime("%d/%m/%Y"), "{{SETOR}}": setor_final, "{{CENTRO_CUSTO}}": ""}
                            conteudo_ficha_docx = preencher_ficha_docx(TEMPLATE_FICHA, m_f, df_e)
                            nome_ficha_docx = f"Ficha EPI {nome_sel}.docx"
                            arquivos[nome_ficha_docx] = conteudo_ficha_docx
                            
                            if incluir_pdf:
                                pdf_bytes, nome_pdf = converter_para_pdf_linux(conteudo_ficha_docx, nome_ficha_docx)
                                if pdf_bytes: arquivos[nome_pdf] = pdf_bytes
                        else:
                            st.error(f"❌ Erro crítico: A aba de EPIs para o cargo '{cargo_limpo}' não pôde ser baixada.")

                    # 3. Certificado NR06
                    if g_cert:
                        prs = Presentation(t_nr)
                        
                        if "SÃO JOSÉ" in unid_sel.upper():
                            local_data_string = f"{data_extenso_pt()}."
                        else:
                            local_data_string = f"{unid_sel.title()}, {data_extenso_pt()}."

                        substituir_pptx(prs, {
                            "{{NOME}}": dados_colab['Nome Colaborador'], 
                            "{{CPF}}": cpf_final, 
                            "{{FUNCAO}}": cargo, 
                            "{{DATA_TREINAMENTO}}": datetime.now().strftime("%d/%m/%Y"), 
                            "{{LOCAL_DATA}}": local_data_string
                        })
                        b = io.BytesIO(); prs.save(b)
                        conteudo_pptx = b.getvalue()
                        nome_pptx = f"NR06 {nome_sel}.pptx"
                        arquivos[nome_pptx] = conteudo_pptx
                        
                        if incluir_pdf:
                            pdf_bytes, nome_pdf = converter_para_pdf_linux(conteudo_pptx, nome_pptx)
                            if pdf_bytes: arquivos[nome_pdf] = pdf_bytes

                    if arquivos:
                        z_b = io.BytesIO()
                        with zipfile.ZipFile(z_b, "w") as z:
                            for n, d in arquivos.items(): z.writestr(n, d)
                        st.success("✅ Documentos prontos!")
                        st.download_button("📦 BAIXAR KIT COMPLETO (ZIP)", z_b.getvalue(), f"Kit_{nome_sel}.zip", use_container_width=True)
                else:
                    st.error(f"Cargo '{cargo}' não encontrado na aba Cargos.")
    else:
        st.error("Erro ao carregar dados da planilha Google. Verifique o acesso público ou as abas.")

    st.markdown('<div class="portal-footer">© 2026 Gestão Documentos | Desenvolvido por: Dilceu Junior</div>', unsafe_allow_html=True)

def main():
    if "autenticado" not in st.session_state:
        st.session_state.autenticado = False

    if not st.session_state.autenticado:
        tela_login()
    else:
        if _dias_desde_cadastro(st.session_state.usuario) < TRIAL_DIAS or verificar_adimplencia().get("adimplente"):
            tela_portal()
        else:
            _injetar_css()
            st.error("Acesso bloqueado por inadimplência.")

if __name__ == "__main__":
    main()
