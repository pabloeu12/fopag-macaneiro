"""
app.py
──────
Interface Streamlit da Conferência de Folha — Maçaneiro e Bwise.
Execute com:  streamlit run app.py
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
# FUNÇÃO PARA ENCONTRAR AS LOGOS (PNG ou JPG)
# ════════════════════════════════════════════════════════════
def buscar_imagem(nome_base):
    """
    Busca o arquivo de imagem tentando diferentes extensões 
    e variações de nome (com espaço ou underline).
    """
    variacoes_nome = [nome_base, nome_base.replace(" ", "_"), nome_base.replace("_", " ")]
    extensoes = [".png", ".jpg", ".jpeg", ".PNG", ".JPG", ".JPEG"]
    
    for nome in variacoes_nome:
        for ext in extensoes:
            caminho = f"{nome}{ext}"
            if os.path.exists(caminho):
                return caminho
    return None

# Buscando as logos
logo_bwise = buscar_imagem("logo_bwise")
logo_macaneiro = buscar_imagem("logo_macaneiro")

# ════════════════════════════════════════════════════════════
# MENU LATERAL (SIDEBAR) - IDENTICO AO ADIANTAMENTO
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

# Coluna 1: Logo Bwise
with col1:
    if logo_bwise:
        st.image(logo_bwise, use_column_width=True)
    else:
        st.warning("Logo Bwise não encontrada\n(Verifique se o nome é 'logo bwise' .png ou .jpg).")

# Coluna 2: Título Central
with col2:
    st.markdown(
        """
        <h1 style='text-align: center; font-size: 3rem; font-weight: 800; line-height: 1.2; margin-top: 10px;'>
            CONFERÊNCIA<br>DE FOLHA
        </h1>
        """, 
        unsafe_allow_html=True
    )

# Coluna 3: Logo Maçaneiro
with col3:
    if logo_macaneiro:
        st.image(logo_macaneiro, use_column_width=True)
    else:
        st.warning("Logo Maçaneiro não encontrada\n(Verifique se o nome é 'logo macaneiro' .png ou .jpg).")

st.markdown("---")

# ════════════════════════════════════════════════════════════
# MENSAGEM INICIAL E TRAVA (Se não houver arquivos)
# ════════════════════════════════════════════════════════════
if not arq_lanc or not arq_sist:
    with st.expander("📖 Como extrair e preparar as planilhas (Passo a Passo)"):
        st.write("1. Faça a exportação da planilha de lançamentos do mês correspondente.")
        st.write("2. Faça a exportação da planilha de eventos do sistema.")
        st.write("3. Anexe os dois arquivos no menu lateral esquerdo (📁 Importação de Dados).")
        
    st.info("👉 Por favor, anexe as DUAS planilhas no menu lateral para iniciar a conferência.")
    st.stop() # Para a execução do app aqui até que os arquivos sejam enviados

# ════════════════════════════════════════════════════════════
# EXECUÇÃO DA COMPARAÇÃO (Aparece só quando os anexos estão OK)
# ════════════════════════════════════════════════════════════
col_btn, _ = st.columns([1, 3])
with col_btn:
    btn_comparar = st.button("⚡  INICIAR COMPARAÇÃO", use_container_width=True)

if btn_comparar:
    with st.spinner("⚙️  Processando cruzamento de dados..."):
        try:
            resultados = executar_comparacao(
                BytesIO(arq_lanc.read()),
                BytesIO(arq_sist.read()),
            )
        except Exception as e:
            st.error(f"❌ Erro durante a comparação: {e}")
            st.stop()

    if not resultados:
        st.warning("⚠️  Nenhum evento foi comparado. Verifique se os arquivos estão corretos.")
        st.stop()

    df = pd.DataFrame(resultados)
    st.session_state["df"]         = df
    st.session_state["resultados"] = resultados
    st.success(f"✅ Comparação concluída! {len(df)} eventos processados.")


# ════════════════════════════════════════════════════════════
# RESULTADOS E MÉTRICAS
# ════════════════════════════════════════════════════════════
if "df" in st.session_state:
    df         = st.session_state["df"]
    resultados = st.session_state["resultados"]

    total   = len(df)
    ok_tot  = df["Status"].str.startswith("OK").sum()
    diverg  = (df["Status"] == "DIVERGENTE").sum()
    nao_enc = (df["Status"] == "NAO_ENCONTRADO").sum()

    st.markdown("<br>", unsafe_allow_html=True)
    st.markdown("### 📊 Resultado Geral")

    # Métricas nativas do Streamlit
    m1, m2, m3, m4 = st.columns(4)
    m1.metric("Eventos Comparados", total)
    m2.metric("Conferidos OK ✅", ok_tot)
    m3.metric("Divergências ❌", diverg)
    m4.metric("Não Encontrados ⚠️", nao_enc)

    st.markdown("---")

    # ── Filtros ──────────────────────────────────────────────
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

    # Aplica filtros
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

    st.caption(f"Exibindo {len(df_filtrado)} de {total} eventos")

    # ── Tabela ───────────────────────────────────────────────
    def colorir_linha(row):
        status = row["Status"]
        if "OK" in str(status):
            cor = "background-color: rgba(63, 185, 80, 0.1);"
        elif status == "DIVERGENTE":
            cor = "background-color: rgba(248, 81, 73, 0.1);"
        elif status == "NAO_ENCONTRADO":
            cor = "background-color: rgba(210, 153, 34, 0.1);"
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

    # ── Download ─────────────────────────────────────────────
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
                label="⚠️  Baixar Divergências (CSV)",
                data=csv_div,
                file_name="DIVERGENCIAS.csv",
                mime="text/csv",
                use_container_width=True
            )
