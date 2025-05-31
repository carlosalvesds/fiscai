# importar as bibliotecas
import streamlit as st
import pandas as pd
from PIL import Image
import base64  #  Adicione isto
import io      #  Necessário para converter imagem para base64

# criar as funcões de carregamento de dados
# verificar etapas ferramentas do lado esquerdo talvez pedir ajuda para o GPT


# preparar as visualizações










# INTERFACE OK

# Cache da imagem convertida
@st.cache_data
def carregar_banner_base64():
    image = Image.open("fiscai_banner.png")
    buffered = io.BytesIO()
    image.save(buffered, format="PNG")
    return base64.b64encode(buffered.getvalue()).decode()
# Configurações da página

st.set_page_config(
    page_title="FiscAI",
    layout="wide",
    page_icon="💻",
    initial_sidebar_state="collapsed"
)

# CSS personalizado para transparência da sidebar
st.markdown("""
    <style>
        /* Tornar a sidebar transparente */
        section[data-testid="stSidebar"] {
            background-color: rgba(0, 0, 0, 0.0);  /* transparente */
        }

        /* Remover sombra e bordas da sidebar */
        section[data-testid="stSidebar"] > div:first-child {
            box-shadow: none;
        }
    </style>
""", unsafe_allow_html=True)

# Menu lateral
st.sidebar.header("🛠️ Ferramentas")
opcao = st.sidebar.radio("Escolha uma opção:", [
    "🏠 Início",
    "📂 Leitor XML | Regime Tributário"
])


# Home com imagem
if opcao == "🏠 Início":
    import time
    from datetime import datetime

    # Mostrar imagem no topo
    image = Image.open("fiscai_banner.png")
    buffered = io.BytesIO()
    image.save(buffered, format="PNG")
    img_base64 = base64.b64encode(buffered.getvalue()).decode()

    st.markdown(
        f"""
        <div style='text-align: center; margin-bottom: 1rem;'>
            <img src="data:image/png;base64,{img_base64}" width="900">
        </div>
        """,
        unsafe_allow_html=True
    )

    # CSS para relógio centralizado e discreto
    st.markdown("""
    <style>
        .clock-container {
            display: flex;
            justify-content: center;
            align-items: center;
            margin-top: 1rem;
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

    # Espaço reservado para o relógio
    clock_placeholder = st.empty()

    # Loop para atualização ao vivo por 3 minutos
    for _ in range(200):
        agora = datetime.now()
        data_reforma = datetime(2026, 1, 1)
        tempo_restante = data_reforma - agora
        dias = tempo_restante.days
        horas, resto = divmod(tempo_restante.seconds, 3600)
        minutos, segundos = divmod(resto, 60)

        with clock_placeholder.container():
            st.markdown(f"""
                
    <div class="clock-container" style="flex-direction: column;">
        <div style="color: #ffffff; font-size: 1.1rem; margin-bottom: 0.3rem;">
            Contagem Regressiva para a Reforma Tributária
        </div>
        <div class="clock">
            ⏳ {dias}d : {horas:02d}h : {minutos:02d}m : {segundos:02d}s
        </div>
    </div>
    """, unsafe_allow_html=True)

# Ferramenta: Leitor de XML
elif opcao == "📂 Leitor XML | Regime Tributário":
    from ferramentas.leitor_rt import app as leitor_rt_app
    leitor_rt_app()



