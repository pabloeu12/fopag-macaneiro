"""
app.py
──────
Interface Streamlit da Conferência de Férias — Maçaneiro e Bwise.
"""

import streamlit as st
import pandas as pd
import pdfplumber
import re
import os
from datetime import datetime
from dateutil.relativedelta import relativedelta
from decimal import Decimal, ROUND_HALF_UP

# ════════════════════════════════════════════════════════════
# CONFIGURAÇÃO DA PÁGINA
# ════════════════════════════════════════════════════════════
st.set_page_config(
    page_title="Conferência de Férias",
    page_icon="🏖️",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ════════════════════════════════════════════════════════════
# BUSCA INTELIGENTE DE LOGOS (Resolve os erros de JPG/PNG)
# ════════════════════════════════════════════════════════════
def buscar_logo(palavra_chave):
    """
    Varre a pasta atual e procura qualquer arquivo de imagem que 
    contenha a palavra-chave (ex: 'bwise' ou 'macaneiro').
    """
    try:
        arquivos = os.listdir('.')
        for arq in arquivos:
            if arq.lower().endswith(('.png', '.jpg', '.jpeg')) and palavra_chave in arq.lower():
                return arq
    except Exception:
        pass
    return None

logo_bwise = buscar_logo("bwise")
logo_macaneiro = buscar_logo("macaneiro")

# ════════════════════════════════════════════════════════════
# CSS PERSONALIZADO — Layout Escuro 
# ════════════════════════════════════════════════════════════
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Syne:wght@400;600;700;800&family=DM+Sans:wght@300;400;500&display=swap');

html, body, [class*="css"] {
    font-family: 'DM Sans', sans-serif;
}
.stApp {
    background-color: #0d1117 !important;
    color: #e6edf3 !important;
}
[data-testid="stSidebar"] {
    background-color: #161b22 !important;
    border-right: 1px solid #30363d !important;
}
[data-testid="stFileUploader"] {
    background: #0d1117 !important;
    border: 1.5px dashed #30363d !important;
    border-radius: 10px !important;
}
[data-testid="stFileUploader"]:hover {
    border-color: #4ea8de !important;
}
.titulo-central {
    text-align: center; 
    font-family: 'Syne', sans-serif;
    font-size: 1.8rem; 
    font-weight: 800; 
    line-height: 1.2; 
    margin-top: 15px;
    color: #ffffff;
}
.subtitulo {
    font-family: 'Syne', sans-serif;
    color: #4ea8de;
    font-weight: bold;
    margin-top: 20px;
    margin-bottom: 10px;
}
.stButton > button {
    background: linear-gradient(135deg, #1f4e79, #0078d4) !important;
    color: white !important;
    border: none !important;
    border-radius: 10px !important;
    font-family: 'Syne', sans-serif !important;
    font-weight: 700 !important;
    font-size: 1rem !important;
    padding: 14px 36px !important;
    width: 100% !important;
    transition: all 0.2s !important;
}
.stButton > button:hover {
    box-shadow: 0 4px 24px rgba(0,120,212,0.3) !important;
}
.stDataFrame {
    border: 1px solid #21262d !important;
    border-radius: 8px !important;
}
.stAlert {
    border-radius: 10px !important;
}
</style>
""", unsafe_allow_html=True)

# ════════════════════════════════════════════════════════════
# MENU LATERAL (SIDEBAR)
# ════════════════════════════════════════════════════════════
with st.sidebar:
    st.markdown("### 📁 Importação de Dados")
    st.markdown("<br>", unsafe_allow_html=True)
    
    st.markdown("**1. Upload: RECIBO FÉRIAS (PDF)**")
    pdf_file = st.file_uploader(
        "Upload PDF", 
        type=['pdf'], 
        key="pdf", 
        label_visibility="collapsed"
    )
    
    st.markdown("<br>", unsafe_allow_html=True)
    
    st.markdown("**2. Upload: LISTA DE EVENTOS**")
    eventos_file = st.file_uploader(
        "Upload Eventos", 
        type=['csv', 'xlsx', 'xls'], 
        key="evt", 
        label_visibility="collapsed"
    )

    st.markdown("<br>", unsafe_allow_html=True)
    
    st.markdown("**3. Upload: HISTÓRICO DE CARGOS**")
    historico_file = st.file_uploader(
        "Upload Historico", 
        type=['csv', 'xlsx', 'xls'], 
        key="hist", 
        label_visibility="collapsed"
    )

# ════════════════════════════════════════════════════════════
# CABEÇALHO PRINCIPAL
# ════════════════════════════════════════════════════════════
col1, col2, col3 = st.columns([1, 2.5, 1])

with col1:
    if logo_bwise:
        st.image(logo_bwise, use_column_width=True)
    else:
        st.warning("Logo Bwise não encontrada.")

with col2:
    st.markdown('<div class="titulo-central">CONFERÊNCIA<br>DE FÉRIAS</div>', unsafe_allow_html=True)

with col3:
    if logo_macaneiro:
        st.image(logo_macaneiro, use_column_width=True)
    else:
        st.warning("Logo Maçaneiro não encontrada.")

st.markdown("---")

# ════════════════════════════════════════════════════════════
# REGRAS DE NEGÓCIO E FUNÇÕES AUXILIARES
# ════════════════════════════════════════════════════════════
EVENTOS_MEDIAS = [158, 115, 161, 117, 120, 619, 575, 642, 644, 645, 643, 153, 465, 700, 699, 460]

def arredondar(valor):
    try:
        return float(Decimal(str(valor)).quantize(Decimal('0.01'), rounding=ROUND_HALF_UP))
    except:
        return valor

def parse_br_float(val):
    if pd.isna(val): return 0.0
    if isinstance(val, (int, float)): return float(val)
    v_str = str(val).strip()
    if ',' in v_str:
        v_str = v_str.replace('.', '').replace(',', '.')
    return float(v_str)

def obter_salario_epoca(df_hist, mes_evt, ano_evt):
    if df_hist is None or df_hist.empty:
        return None
    df_temp = df_hist.copy()
    df_temp['Data de reajuste'] = pd.to_datetime(df_temp['Data de reajuste'], errors='coerce', dayfirst=True)
    df_temp = df_temp.dropna(subset=['Data de reajuste']).sort_values(by='Data de reajuste', ascending=True)
    
    if df_temp.empty: return None
    salario_vigente = float(df_temp.iloc[0]['Salário'])
    
    for _, row in df_temp.iterrows():
        data_reajuste = row['Data de reajuste']
        if (ano_evt > data_reajuste.year) or (ano_evt == data_reajuste.year and mes_evt >= data_reajuste.month):
            salario_vigente = float(row['Salário novo'])
        else:
            break
    return salario_vigente

def extrair_dados_pdf(pdf_obj):
    dados = {'eventos': {}}
    with pdfplumber.open(pdf_obj) as pdf:
        texto = ""
        for page in pdf.pages:
            texto += page.extract_text() + "\n"
            
    match_salario = re.search(r'Salário Contratual:[\s]*([\d\.,]+)', texto)
    if match_salario: dados['salario'] = float(match_salario.group(1).replace('.', '').replace(',', '.'))
        
    match_periodo = re.search(r'Período Aquisitivo:[\s]*([\d]{2}/[\d]{2}/[\d]{4})\s*a\s*([\d]{2}/[\d]{2}/[\d]{4})', texto)
    if match_periodo:
        dados['inicio_aquisitivo'] = datetime.strptime(match_periodo.group(1), '%d/%m/%Y')
        dados['fim_aquisitivo'] = datetime.strptime(match_periodo.group(2), '%d/%m/%Y')
        
    match_mat = re.search(r'Matrícula:[\s]*(\d+)', texto)
    if match_mat: dados['matricula'] = int(match_mat.group(1))

    linhas = texto.split('\n')
    for linha in linhas:
        match_evento = re.search(r'^(\d{4})\s*-\s*(.*?)\s+([\d\.,]+)\s+([\d\.,]+)', linha)
        if match_evento:
            cod = int(match_evento.group(1))
            ref = float(match_evento.group(3).replace(',', '.'))
            provento = float(match_evento.group(4).replace('.', '').replace(',', '.'))
            dados['eventos'][cod] = {'referencia': ref, 'provento': provento}
            
    return dados

def carregar_historico(historico_obj, matricula):
    nome_ficheiro = historico_obj.name.lower()
    if nome_ficheiro.endswith('.csv'): df_hist = pd.read_csv(historico_obj)
    elif nome_ficheiro.endswith('.xlsx'): df_hist = pd.read_excel(historico_obj, engine='openpyxl')
    elif nome_ficheiro.endswith('.xls'): df_hist = pd.read_excel(historico_obj, engine='xlrd')
        
    df_hist.columns = df_hist.columns.str.strip()
    df_hist = df_hist[df_hist['Matrícula'] == matricula]
    return df_hist[['Data de reajuste', 'Salário', 'Salário novo']].dropna()

def carregar_eventos(eventos_obj, matricula):
    nome_ficheiro = eventos_obj.name.lower()
    if nome_ficheiro.endswith('.csv'): df_evt = pd.read_csv(eventos_obj, skiprows=1)
    elif nome_ficheiro.endswith('.xlsx'): df_evt = pd.read_excel(eventos_obj, skiprows=1, engine='openpyxl')
    elif nome_ficheiro.endswith('.xls'): df_evt = pd.read_excel(eventos_obj, skiprows=1, engine='xlrd')
    
    df_evt.columns = df_evt.columns.str.strip()
    df_evt = df_evt[df_evt['Matrícula'] == matricula]
    if 'Valor Provento' in df_evt.columns:
        df_evt['Valor Provento'] = df_evt['Valor Provento'].apply(parse_br_float)
    return df_evt

# ════════════════════════════════════════════════════════════
# TELA INICIAL (SEM ARQUIVOS) E PASSO A PASSO
# ════════════════════════════════════════════════════════════
if not pdf_file or not eventos_file or not historico_file:
    with st.expander("📖 Como extrair os documentos do sistema (Passo a Passo)"):
        st.markdown("""
        ### 1️⃣ RECIBO DE FÉRIAS (PDF)
        **Caminho:** `Férias` ➔ `Controle de Período Aquisitivo e Concessivo` ➔ `Impressão de Documentos...`
        * Clicar em **Avançar** ➔ Selecionar empregado ➔ Selecionar o último período de férias ➔ **Recibo de Férias**.
        * Salvar o arquivo no formato **PDF**.

        ---
        ### 2️⃣ LISTA DE EVENTOS DE RECIBO DE PAGAMENTO
        **Caminho:** `Folha de Pagamento` ➔ `Folha de Pagamento` ➔ `Lista de Eventos de Recibos de Pagamento...`
        * **Competência Inicial e Competência Final:** Selecionar o período aquisitivo.
        * **Tipo de Recibo:** `1 Normal`.
        * Clicar em **Filtrar** e salvar o arquivo (formatos **.CSV**, **.XLS** ou **.XLSX**).

        ---
        ### 3️⃣ HISTÓRICO DE CARGOS E SALÁRIOS
        **Caminho:** `Folha de Pagamento` ➔ `Cadastros` ➔ `Cargos` ➔ `Lista de Histórico de Cargos e Salários...`
        * **Situação do funcionário:** `Ativos`.
        * Salvar o arquivo (formatos **.CSV**, **.XLS** ou **.XLSX**).
        """)
        
    st.info("👉 Por favor, anexe os TRÊS documentos no menu lateral para iniciar a conferência.")
    st.stop()

# ════════════════════════════════════════════════════════════
# COMPARAÇÃO E RESULTADOS
# ════════════════════════════════════════════════════════════
col_btn, _ = st.columns([1, 3])
with col_btn:
    btn_comparar = st.button("⚡  INICIAR CONFERÊNCIA")

# O sistema roda automaticamente ou após clicar no botão
if btn_comparar or (pdf_file and eventos_file and historico_file):
    
    with st.spinner("⚙️ Processando informações e calculando médias..."):
        dados_pdf = extrair_dados_pdf(pdf_file)
        matricula = dados_pdf.get('matricula', 0)
        salario_atual = dados_pdf.get('salario', 0.0)
    
    st.markdown('<div class="subtitulo">📋 Resumo Extraído</div>', unsafe_allow_html=True)
    st.write(f"**Matrícula:** {matricula}")
    st.write(f"**Salário Contratual:** R$ {salario_atual:,.2f}")
    if 'inicio_aquisitivo' in dados_pdf:
        st.write(f"**Período Aquisitivo:** {dados_pdf['inicio_aquisitivo'].strftime('%d/%m/%Y')} a {dados_pdf['fim_aquisitivo'].strftime('%d/%m/%Y')}")
    st.markdown("---")

    # --- PARTE 1: CONFERÊNCIA FÉRIAS E ABONO BASE ---
    st.markdown('<div class="subtitulo">1. Conferência Férias e Abono</div>', unsafe_allow_html=True)
    valor_dia = salario_atual / 30

    col_a, col_b = st.columns(2)
    
    if 189 in dados_pdf['eventos']:
        ref_ferias = dados_pdf['eventos'][189]['referencia']
        valor_pdf_ferias = dados_pdf['eventos'][189]['provento']
        calc_ferias = arredondar(valor_dia * ref_ferias)
        with col_a:
            st.info(f"**0189 - FÉRIAS NORMAIS**\n\n"
                    f"Cálculo: R$ {salario_atual:,.2f} / 30 * {ref_ferias} = **R$ {calc_ferias:,.2f}**\n\n"
                    f"Valor no PDF: **R$ {valor_pdf_ferias:,.2f}**\n\n"
                    f"Diferença: R$ {(calc_ferias - valor_pdf_ferias):,.2f}")
                    
    if 191 in dados_pdf['eventos']:
        ref_abono = dados_pdf['eventos'][191]['referencia']
        valor_pdf_abono = dados_pdf['eventos'][191]['provento']
        calc_abono = arredondar(valor_dia * ref_abono)
        with col_b:
            st.info(f"**0191 - ABONO PECUNIÁRIO**\n\n"
                    f"Cálculo: R$ {salario_atual:,.2f} / 30 * {ref_abono} = **R$ {calc_abono:,.2f}**\n\n"
                    f"Valor no PDF: **R$ {valor_pdf_abono:,.2f}**\n\n"
                    f"Diferença: R$ {(calc_abono - valor_pdf_abono):,.2f}")

    st.markdown("---")

    # --- PARTE 2: CONFERÊNCIA DE MÉDIAS VARIÁVEIS ---
    st.markdown('<div class="subtitulo">2. Conferência de Médias de Variáveis</div>', unsafe_allow_html=True)
    
    if 'inicio_aquisitivo' in dados_pdf:
        meses_calculo = []
        data_iter = dados_pdf['inicio_aquisitivo']
        for _ in range(12):
            meses_calculo.append((data_iter.month, data_iter.year))
            data_iter += relativedelta(months=1)

        try:
            df_hist = carregar_historico(historico_file, matricula)
            df_evt = carregar_eventos(eventos_file, matricula)
            
            total_medias_ajustadas = 0.0
            detalhes_medias = []
            
            for _, row in df_evt.iterrows():
                try:
                    mes_evt = int(row['Mês'])
                    ano_evt = int(row['Ano'])
                    cod_evt = int(row['Cód. Evento'])
                except:
                    continue
                
                if (mes_evt, ano_evt) in meses_calculo and cod_evt in EVENTOS_MEDIAS:
                    valor_original = float(row['Valor Provento'])
                    salario_epoca = obter_salario_epoca(df_hist, mes_evt, ano_evt)
                    
                    if salario_epoca and salario_epoca < salario_atual:
                        valor_ajustado = arredondar((valor_original / salario_epoca) * salario_atual)
                    else:
                        valor_ajustado = arredondar(valor_original)
                        
                    total_medias_ajustadas += valor_ajustado
                    detalhes_medias.append(f"{mes_evt:02d}/{ano_evt} - Cód {cod_evt}: R$ {valor_original:,.2f} (Base: R$ {salario_epoca or salario_atual:,.2f}) -> Corrigido: R$ {valor_ajustado:,.2f}")

            media_mensal = arredondar(total_medias_ajustadas / 12)
            
            st.write(f"**Soma Total de Proventos Atualizados:** R$ {total_medias_ajustadas:,.2f}")
            st.write(f"**Média Apurada (A dividir por 12):** R$ {media_mensal:,.2f}")
            
            with st.expander("Ver detalhamento dos eventos de médias da competência"):
                for det in detalhes_medias:
                    st.write(det)

            st.write("")
            col_c, col_d = st.columns(2)
            
            if 223 in dados_pdf['eventos']: 
                ref_med_ferias = dados_pdf['eventos'][223]['referencia']
                calc_med_ferias = arredondar((media_mensal / 30) * ref_med_ferias)
                with col_c:
                    st.success(f"**0223 - MEDIAS S/ VARIAVEIS - FÉRIAS**\n\n"
                               f"Cálculo: R$ {media_mensal:,.2f} / 30 * {ref_med_ferias} = **R$ {calc_med_ferias:,.2f}**\n\n"
                               f"Valor PDF: **R$ {dados_pdf['eventos'][223]['provento']:,.2f}**\n\n"
                               f"Diferença: R$ {(calc_med_ferias - dados_pdf['eventos'][223]['provento']):,.2f}")
                               
            if 224 in dados_pdf['eventos']:
                ref_med_abono = dados_pdf['eventos'][224]['referencia']
                calc_med_abono = arredondar((media_mensal / 30) * ref_med_abono)
                with col_d:
                    st.success(f"**0224 - MEDIAS S/ VARIAVEIS - ABONO**\n\n"
                               f"Cálculo: R$ {media_mensal:,.2f} / 30 * {ref_med_abono} = **R$ {calc_med_abono:,.2f}**\n\n"
                               f"Valor PDF: **R$ {dados_pdf['eventos'][224]['provento']:,.2f}**\n\n"
                               f"Diferença: R$ {(calc_med_abono - dados_pdf['eventos'][224]['provento']):,.2f}")
        
        except Exception as e:
            st.error(f"❌ Erro durante o processamento das planilhas: {e}")
