"""
app.py
──────
Interface Streamlit da Conferência de Folha — Maçaneiro.
Execute com:  streamlit run app.py
"""

import pandas as pd
import streamlit as st
from io import BytesIO
from comparador import executar_comparacao, gerar_excel

# ════════════════════════════════════════════════════════════
# CONFIGURAÇÃO DA PÁGINA
# ════════════════════════════════════════════════════════════
st.set_page_config(
    page_title="Conferência de Folha — Maçaneiro",
    page_icon="📋",
    layout="wide",
    initial_sidebar_state="collapsed",
)

# ════════════════════════════════════════════════════════════
# CSS PERSONALIZADO — Design profissional
# ════════════════════════════════════════════════════════════
st.markdown("""
<style>
/* ── Importa fontes ── */
@import url('https://fonts.googleapis.com/css2?family=Syne:wght@400;600;700;800&family=DM+Sans:wght@300;400;500&display=swap');

/* ── Reset e base ── */
html, body, [class*="css"] {
    font-family: 'DM Sans', sans-serif;
}

/* ── Fundo geral ── */
.stApp {
    background: #0d1117;
    color: #e6edf3;
}

/* ── Bloco do cabeçalho principal ── */
.header-block {
    background: linear-gradient(135deg, #0f2942 0%, #1a3a5c 50%, #0f2942 100%);
    border: 1px solid #1f4e79;
    border-radius: 16px;
    padding: 40px 48px;
    margin-bottom: 32px;
    position: relative;
    overflow: hidden;
}
.header-block::before {
    content: '';
    position: absolute;
    top: -60px; right: -60px;
    width: 220px; height: 220px;
    background: radial-gradient(circle, rgba(0,120,212,0.18) 0%, transparent 70%);
    border-radius: 50%;
}
.header-title {
    font-family: 'Syne', sans-serif;
    font-size: 2.1rem;
    font-weight: 800;
    color: #ffffff;
    margin: 0 0 6px 0;
    letter-spacing: -0.5px;
    line-height: 1.15;
}
.header-sub {
    font-size: 0.95rem;
    color: #7d9bbd;
    margin: 0;
    font-weight: 300;
    letter-spacing: 0.3px;
}
.header-badge {
    display: inline-block;
    background: rgba(0,120,212,0.2);
    border: 1px solid rgba(0,120,212,0.4);
    color: #4ea8de;
    font-size: 0.72rem;
    font-weight: 600;
    letter-spacing: 1.5px;
    text-transform: uppercase;
    padding: 4px 12px;
    border-radius: 20px;
    margin-bottom: 14px;
}

/* ── Cards de upload ── */
.upload-card {
    background: #161b22;
    border: 1px solid #21262d;
    border-radius: 12px;
    padding: 24px;
    height: 100%;
    transition: border-color 0.2s;
}
.upload-card:hover {
    border-color: #1f4e79;
}
.upload-label {
    font-family: 'Syne', sans-serif;
    font-size: 0.8rem;
    font-weight: 700;
    letter-spacing: 1.2px;
    text-transform: uppercase;
    color: #4ea8de;
    margin-bottom: 4px;
}
.upload-desc {
    font-size: 0.82rem;
    color: #6e7681;
    margin-bottom: 16px;
}

/* ── Métricas customizadas ── */
.metric-grid {
    display: grid;
    grid-template-columns: repeat(4, 1fr);
    gap: 16px;
    margin-bottom: 28px;
}
.metric-card {
    background: #161b22;
    border: 1px solid #21262d;
    border-radius: 12px;
    padding: 22px 24px;
    position: relative;
    overflow: hidden;
}
.metric-card::after {
    content: '';
    position: absolute;
    bottom: 0; left: 0; right: 0;
    height: 3px;
}
.metric-card.total::after  { background: #4ea8de; }
.metric-card.ok::after     { background: #3fb950; }
.metric-card.diverg::after { background: #f85149; }
.metric-card.naoenc::after { background: #d29922; }
.metric-icon {
    font-size: 1.4rem;
    margin-bottom: 10px;
    display: block;
}
.metric-value {
    font-family: 'Syne', sans-serif;
    font-size: 2.4rem;
    font-weight: 800;
    line-height: 1;
    margin-bottom: 4px;
}
.metric-card.total  .metric-value { color: #4ea8de; }
.metric-card.ok     .metric-value { color: #3fb950; }
.metric-card.diverg .metric-value { color: #f85149; }
.metric-card.naoenc .metric-value { color: #d29922; }
.metric-label {
    font-size: 0.78rem;
    color: #6e7681;
    font-weight: 500;
    letter-spacing: 0.3px;
}
.metric-pct {
    position: absolute;
    top: 18px; right: 18px;
    font-size: 0.75rem;
    font-weight: 600;
    color: #30363d;
}

/* ── Barra de progresso de acerto ── */
.acerto-bar-wrap {
    background: #161b22;
    border: 1px solid #21262d;
    border-radius: 12px;
    padding: 20px 24px;
    margin-bottom: 28px;
}
.acerto-label {
    font-family: 'Syne', sans-serif;
    font-size: 0.78rem;
    font-weight: 700;
    letter-spacing: 1px;
    text-transform: uppercase;
    color: #6e7681;
    margin-bottom: 10px;
}
.acerto-row {
    display: flex;
    align-items: center;
    gap: 16px;
}
.acerto-track {
    flex: 1;
    height: 10px;
    background: #21262d;
    border-radius: 99px;
    overflow: hidden;
}
.acerto-fill {
    height: 100%;
    border-radius: 99px;
    background: linear-gradient(90deg, #1f4e79, #3fb950);
    transition: width 0.6s ease;
}
.acerto-pct {
    font-family: 'Syne', sans-serif;
    font-size: 1.5rem;
    font-weight: 800;
    color: #3fb950;
    min-width: 64px;
    text-align: right;
}

/* ── Seção de filtros ── */
.section-title {
    font-family: 'Syne', sans-serif;
    font-size: 0.78rem;
    font-weight: 700;
    letter-spacing: 1.2px;
    text-transform: uppercase;
    color: #4ea8de;
    margin-bottom: 14px;
    padding-bottom: 8px;
    border-bottom: 1px solid #21262d;
}

/* ── Tabela ── */
.stDataFrame {
    border: 1px solid #21262d !important;
    border-radius: 8px !important;
}

/* ── Botão principal ── */
.stButton > button {
    background: linear-gradient(135deg, #1f4e79, #0078d4) !important;
    color: white !important;
    border: none !important;
    border-radius: 10px !important;
    font-family: 'Syne', sans-serif !important;
    font-weight: 700 !important;
    font-size: 1rem !important;
    letter-spacing: 0.5px !important;
    padding: 14px 36px !important;
    width: 100% !important;
    transition: all 0.2s !important;
    box-shadow: 0 4px 24px rgba(0,120,212,0.25) !important;
}
.stButton > button:hover {
    background: linear-gradient(135deg, #0078d4, #106ebe) !important;
    transform: translateY(-1px) !important;
    box-shadow: 0 6px 28px rgba(0,120,212,0.35) !important;
}

/* ── Botão de download ── */
.stDownloadButton > button {
    background: linear-gradient(135deg, #145a32, #1e8449) !important;
    color: white !important;
    border: none !important;
    border-radius: 10px !important;
    font-family: 'Syne', sans-serif !important;
    font-weight: 700 !important;
    letter-spacing: 0.5px !important;
    padding: 12px 28px !important;
    width: 100% !important;
    box-shadow: 0 4px 20px rgba(46,160,67,0.25) !important;
}
.stDownloadButton > button:hover {
    background: linear-gradient(135deg, #1e8449, #27ae60) !important;
    transform: translateY(-1px) !important;
}

/* ── Uploader ── */
[data-testid="stFileUploader"] {
    background: #0d1117 !important;
    border: 1.5px dashed #30363d !important;
    border-radius: 10px !important;
}
[data-testid="stFileUploader"]:hover {
    border-color: #4ea8de !important;
}

/* ── Selectbox e multiselect ── */
[data-testid="stSelectbox"] > div > div,
[data-testid="stMultiSelect"] > div > div {
    background: #161b22 !important;
    border-color: #30363d !important;
    color: #e6edf3 !important;
}

/* ── Alerta de sucesso/erro ── */
.stAlert {
    border-radius: 10px !important;
}

/* ── Divider ── */
hr {
    border-color: #21262d !important;
}

/* ── Scrollbar ── */
::-webkit-scrollbar { width: 6px; height: 6px; }
::-webkit-scrollbar-track { background: #0d1117; }
::-webkit-scrollbar-thumb { background: #30363d; border-radius: 3px; }
::-webkit-scrollbar-thumb:hover { background: #4ea8de; }

/* ── Status chips na tabela ── */
.chip-ok      { background:#1a3a28; color:#3fb950; padding:2px 10px; border-radius:99px; font-size:0.75rem; font-weight:600; }
.chip-div     { background:#3d1a1a; color:#f85149; padding:2px 10px; border-radius:99px; font-size:0.75rem; font-weight:600; }
.chip-nao     { background:#3a2d10; color:#d29922; padding:2px 10px; border-radius:99px; font-size:0.75rem; font-weight:600; }

/* ── Tag de arquivo carregado ── */
.file-ok {
    display: inline-flex;
    align-items: center;
    gap: 6px;
    background: rgba(63,185,80,0.1);
    border: 1px solid rgba(63,185,80,0.3);
    color: #3fb950;
    font-size: 0.8rem;
    font-weight: 500;
    padding: 4px 12px;
    border-radius: 99px;
    margin-top: 8px;
}
</style>
""", unsafe_allow_html=True)


# ════════════════════════════════════════════════════════════
# CABEÇALHO
# ════════════════════════════════════════════════════════════
st.markdown("""
<div class="header-block">
    <div class="header-badge">Departamento Pessoal</div>
    <p class="header-title">Conferência de Folha<br>Maçaneiro</p>
    <p class="header-sub">Compare lançamentos manuais com eventos do sistema · Identifique divergências automaticamente · Gere relatório em Excel</p>
</div>
""", unsafe_allow_html=True)


# ════════════════════════════════════════════════════════════
# UPLOAD DOS ARQUIVOS
# ════════════════════════════════════════════════════════════
st.markdown('<p class="section-title">📂 Arquivos do mês</p>', unsafe_allow_html=True)

col_l, col_r = st.columns(2, gap="large")

with col_l:
    st.markdown("""
    <div class="upload-card">
        <div class="upload-label">01 — Planilha de Lançamentos</div>
        <div class="upload-desc">Formato horizontal · Colunas por rubrica</div>
    </div>
    """, unsafe_allow_html=True)
    arq_lanc = st.file_uploader(
        "Lançamentos",
        type=["xlsx"],
        key="lanc",
        label_visibility="collapsed",
    )
    if arq_lanc:
        st.markdown(f'<span class="file-ok">✓ {arq_lanc.name}</span>', unsafe_allow_html=True)

with col_r:
    st.markdown("""
    <div class="upload-card">
        <div class="upload-label">02 — Planilha do Sistema</div>
        <div class="upload-desc">Formato vertical · Uma linha por evento</div>
    </div>
    """, unsafe_allow_html=True)
    arq_sist = st.file_uploader(
        "Sistema",
        type=["xlsx"],
        key="sist",
        label_visibility="collapsed",
    )
    if arq_sist:
        st.markdown(f'<span class="file-ok">✓ {arq_sist.name}</span>', unsafe_allow_html=True)

st.markdown("<br>", unsafe_allow_html=True)


# ════════════════════════════════════════════════════════════
# BOTÃO COMPARAR
# ════════════════════════════════════════════════════════════
col_btn, col_vazio = st.columns([1, 2])
with col_btn:
    btn_comparar = st.button("⚡  COMPARAR", disabled=(not arq_lanc or not arq_sist))

if not arq_lanc or not arq_sist:
    st.markdown(
        "<p style='color:#6e7681; font-size:0.85rem; margin-top:8px;'>"
        "⬆️  Faça o upload das duas planilhas para habilitar a comparação."
        "</p>",
        unsafe_allow_html=True,
    )


# ════════════════════════════════════════════════════════════
# EXECUÇÃO
# ════════════════════════════════════════════════════════════
if btn_comparar:
    with st.spinner("⚙️  Processando comparação..."):
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
# RESULTADOS (só aparece após comparação)
# ════════════════════════════════════════════════════════════
if "df" in st.session_state:
    df         = st.session_state["df"]
    resultados = st.session_state["resultados"]

    total   = len(df)
    ok_tot  = df["Status"].str.startswith("OK").sum()
    diverg  = (df["Status"] == "DIVERGENTE").sum()
    nao_enc = (df["Status"] == "NAO_ENCONTRADO").sum()
    perc    = round(ok_tot / total * 100, 1) if total else 0

    st.markdown("<br>", unsafe_allow_html=True)
    st.markdown('<p class="section-title">📊 Resultado geral</p>', unsafe_allow_html=True)

    # ── Métricas ─────────────────────────────────────────────
    st.markdown(f"""
    <div class="metric-grid">
        <div class="metric-card total">
            <span class="metric-icon">📋</span>
            <div class="metric-value">{total}</div>
            <div class="metric-label">Eventos comparados</div>
        </div>
        <div class="metric-card ok">
            <span class="metric-icon">✅</span>
            <div class="metric-value">{ok_tot}</div>
            <div class="metric-label">Conferidos OK</div>
            <span class="metric-pct">{round(ok_tot/total*100,1) if total else 0}%</span>
        </div>
        <div class="metric-card diverg">
            <span class="metric-icon">❌</span>
            <div class="metric-value">{diverg}</div>
            <div class="metric-label">Divergências</div>
            <span class="metric-pct">{round(diverg/total*100,1) if total else 0}%</span>
        </div>
        <div class="metric-card naoenc">
            <span class="metric-icon">⚠️</span>
            <div class="metric-value">{nao_enc}</div>
            <div class="metric-label">Não encontrados</div>
            <span class="metric-pct">{round(nao_enc/total*100,1) if total else 0}%</span>
        </div>
    </div>
    """, unsafe_allow_html=True)

    # ── Barra de acerto ──────────────────────────────────────
    cor_pct = "#3fb950" if perc >= 90 else "#d29922" if perc >= 70 else "#f85149"
    st.markdown(f"""
    <div class="acerto-bar-wrap">
        <div class="acerto-label">Percentual de acerto geral</div>
        <div class="acerto-row">
            <div class="acerto-track">
                <div class="acerto-fill" style="width:{perc}%"></div>
            </div>
            <div class="acerto-pct" style="color:{cor_pct}">{perc}%</div>
        </div>
    </div>
    """, unsafe_allow_html=True)

    st.markdown("<hr>", unsafe_allow_html=True)

    # ── Filtros ──────────────────────────────────────────────
    st.markdown('<p class="section-title">🔍 Filtros</p>', unsafe_allow_html=True)
    fc1, fc2, fc3 = st.columns([1.5, 2, 2])

    with fc1:
        opcoes_status = ["Todos"] + sorted(df["Status"].unique().tolist())
        filtro_status = st.selectbox("Status", opcoes_status)

    with fc2:
        filtro_func = st.text_input("Funcionário (busca por nome ou matrícula)", placeholder="Ex: JOSE ou 276")

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

    st.markdown(f"<p style='color:#6e7681; font-size:0.83rem; margin-bottom:8px;'>Exibindo {len(df_filtrado)} de {total} eventos</p>", unsafe_allow_html=True)

    # ── Tabela colorida ──────────────────────────────────────
    def colorir_linha(row):
        status = row["Status"]
        if "OK" in str(status):
            cor = "background-color:#0d2818; color:#e6edf3"
        elif status == "DIVERGENTE":
            cor = "background-color:#2d1317; color:#e6edf3"
        elif status == "NAO_ENCONTRADO":
            cor = "background-color:#271f0d; color:#e6edf3"
        else:
            cor = ""
        return [cor] * len(row)

    st.dataframe(
        df_filtrado.style.apply(colorir_linha, axis=1),
        use_container_width=True,
        height=420,
        hide_index=True,
    )

    st.markdown("<br>", unsafe_allow_html=True)

    # ── Download ─────────────────────────────────────────────
    st.markdown('<p class="section-title">⬇️  Exportar relatório</p>', unsafe_allow_html=True)

    excel_bytes = gerar_excel(resultados)

    dcol1, dcol2, dcol3 = st.columns([1, 1, 2])
    with dcol1:
        st.download_button(
            label="📥  Baixar Excel Completo",
            data=excel_bytes,
            file_name="CONFERENCIA_FINAL.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
    with dcol2:
        # CSV dos divergentes apenas
        df_div = df[df["Status"] == "DIVERGENTE"]
        if not df_div.empty:
            csv_div = df_div.to_csv(index=False, sep=";", decimal=",").encode("utf-8-sig")
            st.download_button(
                label="⚠️  Baixar só Divergências (CSV)",
                data=csv_div,
                file_name="DIVERGENCIAS.csv",
                mime="text/csv",
            )

    st.markdown("<br><br>", unsafe_allow_html=True)
    st.markdown(
        "<p style='color:#30363d; font-size:0.78rem; text-align:center;'>"
        "Conferência de Folha Maçaneiro · Bwise · Departamento Pessoal"
        "</p>",
        unsafe_allow_html=True,
    )
