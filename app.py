"""
app.py
──────
Interface Streamlit da Conferência de Folha — Maçaneiro e Bwise.
"""

import pandas as pd
import streamlit as st
import os
from io import BytesIO
from comparador import executar_comparacao, gerar_excel

# ════════════════════════════════════════════════════════════
# CONFIGURAÇÃO DA PÁGINA
# ════════════════════════════════════════════════════════════
st.set_page_config(
    page_title="Conferência de Folha",
    page_icon="📋",
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
# CSS PERSONALIZADO — Layout Escuro (Idêntico ao Adiantamento)
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
    font-size: 3rem; 
    font-weight: 800; 
    line-height: 1.2; 
    margin-top: 10px;
    color: #ffffff;
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
.stDownloadButton > button {
    background: linear-gradient(135deg, #145a32, #1e8449) !important;
    color: white !important;
    border: none !important;
    border-radius: 10px !important;
    font-family: 'Syne', sans-serif !important;
    font-weight: 700 !important;
    width: 100% !important;
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
    
    st.markdown("**1. Upload: PLANILHA DE LANÇAMENTOS**")
    arq_lanc = st.file_uploader(
        "Upload Lançamentos", 
        type=["xlsx", "csv"], 
        key="lanc", 
        label_visibility="collapsed"
    )
    
    st.markdown("<br>", unsafe_allow_html=True)
    
    st.markdown("**2. Upload: PLANILHA DO SISTEMA**")
    arq_sist = st.file_uploader(
        "Upload Sistema", 
        type=["xlsx", "csv"], 
        key="sist", 
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
    st.markdown('<div class="titulo-central">CONFERÊNCIA<br>DE FOLHA</div>', unsafe_allow_html=True)

with col3:
    if logo_macaneiro:
        st.image(logo_macaneiro, use_column_width=True)
    else:
        st.warning("Logo Maçaneiro não encontrada.")

st.markdown("---")

# ════════════════════════════════════════════════════════════
# TELA INICIAL (SEM ARQUIVOS)
# ════════════════════════════════════════════════════════════
if not arq_lanc or not arq_sist:
    with st.expander("📖 Como extrair as planilhas do Sistema (Passo a Passo)"):
        st.write("1. Faça a exportação da planilha de lançamentos do mês.")
        st.write("2. Faça a exportação da planilha de eventos do sistema.")
        st.write("3. Anexe os dois arquivos no menu lateral esquerdo (📁 Importação de Dados).")
        
    st.info("👉 Por favor, anexe as DUAS planilhas no menu lateral para iniciar a conferência.")
    st.stop()

# ════════════════════════════════════════════════════════════
# COMPARAÇÃO E RESULTADOS
# ════════════════════════════════════════════════════════════
col_btn, _ = st.columns([1, 3])
with col_btn:
    btn_comparar = st.button("⚡  INICIAR COMPARAÇÃO")

if btn_comparar:
    with st.spinner("⚙️  A processar cruzamento de dados..."):
        try:
            resultados = executar_comparacao(
                BytesIO(arq_lanc.read()),
                BytesIO(arq_sist.read()),
            )
        except Exception as e:
            st.error(f"❌ Erro na comparação: {e}")
            st.stop()

    if not resultados:
        st.warning("⚠️  Nenhum evento foi comparado. Verifique os arquivos.")
        st.stop()

    df = pd.DataFrame(resultados)
    st.session_state["df"]         = df
    st.session_state["resultados"] = resultados
    st.success(f"✅ Comparação concluída com sucesso! ({len(df)} eventos processados)")

if "df" in st.session_state:
    df         = st.session_state["df"]
    resultados = st.session_state["resultados"]

    total   = len(df)
    ok_tot  = df["Status"].str.startswith("OK").sum()
    diverg  = (df["Status"] == "DIVERGENTE").sum()
    nao_enc = (df["Status"] == "NAO_ENCONTRADO").sum()

    st.markdown("<br>", unsafe_allow_html=True)
    st.markdown("### 📊 Resultado Geral")

    m1, m2, m3, m4 = st.columns(4)
    m1.metric("Eventos Comparados", total)
    m2.metric("Conferidos OK ✅", ok_tot)
    m3.metric("Divergências ❌", diverg)
    m4.metric("Não Encontrados ⚠️", nao_enc)

    st.markdown("---")

    st.markdown('### 🔍 Filtros')
    fc1, fc2, fc3 = st.columns([1.5, 2, 2])

    with fc1:
        opcoes_status = ["Todos"] + sorted(df["Status"].unique().tolist())
        filtro_status = st.selectbox("Status", opcoes_status)

    with fc2:
        filtro_func = st.text_input("Funcionário (Nome ou Matrícula)", placeholder="Ex: JOSE ou 276")

    with fc3:
        todos_eventos = sorted(df["Nome do Evento"].unique().tolist())
        filtro_evento = st.multiselect("Evento / Rubrica", todos_eventos, placeholder="Todos")

    df_filtrado = df.copy()
    if filtro_status != "Todos":
        df_filtrado = df_filtrado[df_filtrado["Status"] == filtro_status]
    if filtro_func.strip():
        mask = (
            df_filtrado["Funcionário"].str.contains(filtro_func.strip(), case=False, na=False)
            | df_filtrado["Matrícula"].astype(str).str.contains(filtro_func.strip(), na=False)
        )
        df_filtrado = df_filtrado[mask]
    if filtro_evento:
        df_filtrado = df_filtrado[df_filtrado["Nome do Evento"].isin(filtro_evento)]

    st.caption(f"A exibir {len(df_filtrado)} de {total} eventos")

    def colorir_linha(row):
        status = row["Status"]
        if "OK" in str(status):
            cor = "background-color: rgba(63, 185, 80, 0.15);"
        elif status == "DIVERGENTE":
            cor = "background-color: rgba(248, 81, 73, 0.15);"
        elif status == "NAO_ENCONTRADO":
            cor = "background-color: rgba(210, 153, 34, 0.15);"
        else:
            cor = ""
        return [cor] * len(row)

    st.dataframe(
        df_filtrado.style.apply(colorir_linha, axis=1),
        use_container_width=True,
        height=400,
        hide_index=True,
    )

    st.markdown("<br>", unsafe_allow_html=True)

    st.markdown('### ⬇️ Exportar Relatório')
    excel_bytes = gerar_excel(resultados)

    dcol1, dcol2, dcol3 = st.columns([1, 1, 2])
    with dcol1:
        st.download_button(
            label="📥  Baixar Excel Completo",
            data=excel_bytes,
            file_name="CONFERENCIA_FINAL.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )
    with dcol2:
        df_div = df[df["Status"] == "DIVERGENTE"]
        if not df_div.empty:
            csv_div = df_div.to_csv(index=False, sep=";", decimal=",").encode("utf-8-sig")
            st.download_button(
                label="⚠️  Baixar Divergências",
                data=csv_div,
                file_name="DIVERGENCIAS.csv",
                mime="text/csv",
                use_container_width=True
            )
