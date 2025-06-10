# importar as bibliotecas
import streamlit as st
import pandas as pd
from PIL import Image
import base64
import io
from datetime import datetime
from streamlit_autorefresh import st_autorefresh

# Configurações da página
st.set_page_config(
    page_title="FiscAI",
    layout="wide",
    page_icon="💻",
    initial_sidebar_state="collapsed"
)

# Cache da imagem convertida
@st.cache_data
def carregar_banner_base64():
    image = Image.open("fiscai_banner.png")
    buffered = io.BytesIO()
    image.save(buffered, format="PNG")
    return base64.b64encode(buffered.getvalue()).decode()

# CSS da sidebar
st.markdown("""
    <style>
        section[data-testid="stSidebar"] {
            background-color: rgba(0, 0, 0, 0.0);
        }
        section[data-testid="stSidebar"] > div:first-child {
            box-shadow: none;
        }
    </style>
""", unsafe_allow_html=True)

# Sidebar - Navegação unificada
st.sidebar.markdown("##  Menu")

menu = st.sidebar.radio("Escolha uma opção:", [
    "🏠 Início",
    "📄 Leitor PDF | Energia Elétrica",
    "📂 Leitor XML | Regime Tributário",
    "📊 Resumo     | Natureza da Receita" 
])

# Linha separadora visual
st.sidebar.markdown(
    "<hr style='margin: 10px 5px; border: none; height: 1px; background-color: #00e0ff;'>",
    unsafe_allow_html=True
)

# Exibir conteúdo com base na opção escolhida
if menu == "🏠 Início":
    st_autorefresh(interval=300000, key="relogio_reforma")  # Atualiza a cada 5 min

    img_base64 = carregar_banner_base64()
    st.markdown(
        f"""
        <div style='text-align: center; margin-bottom: 1rem;'>
            <img src="data:image/png;base64,{img_base64}" width="900">
        </div>
        """,
        unsafe_allow_html=True
    )

    st.markdown("""
    <style>
        .clock-container {
            display: flex;
            justify-content: center;
            align-items: center;
            margin-top: 1rem;
            flex-direction: column;
        }
        .clock-label {
            color: #ffffff;
            font-size: 1.1rem;
            margin-bottom: 0.3rem;
        }
        .clock {
            font-family: 'Courier New', monospace;
            font-size: 1.3rem;
            color: #00ffcc;
            background-color: #001f2e;
            padding: 8px 18px;
            border-radius: 10px;
            box-shadow: 0 0 10px #00e0ff55;
            letter-spacing: 3px;
        }
    </style>
    """, unsafe_allow_html=True)

    agora = datetime.now()
    data_reforma = datetime(2026, 1, 1)
    tempo_restante = data_reforma - agora
    dias = tempo_restante.days
    horas = tempo_restante.seconds // 3600

    st.markdown(f"""
    <div class="clock-container">
        <div class="clock-label">Contagem Regressiva para a Reforma Tributária</div>
        <div class="clock">
            ⏳ {dias}d : {horas:02d}h
        </div>
    </div>
    """, unsafe_allow_html=True)

elif menu == "📄 Leitor PDF | Energia Elétrica":
    from ferramentas.leitor_pdf_nf3e import app as leitor_pdf_nf3e_app
    leitor_pdf_nf3e_app()

elif menu == "📂 Leitor XML | Regime Tributário":
    from ferramentas.leitor_rt import app as leitor_rt_app
    leitor_rt_app()

elif menu == "📊 Resumo     | Natureza da Receita":
    from ferramentas.resumo_nat_receita import app as resumo_app
    resumo_app()

