import streamlit as st
import google.generativeai as genai
import os

# Configuração da página e identidade visual CazéTV
st.set_page_config(
    page_title="CazéTV — Decupador de VT",
    page_icon="⚡",
    layout="wide"
)

st.markdown("""
<style>
    .stApp { background-color: #0b0f19; color: #f1f5f9; }
    .main-header { font-size: 2.2rem; font-weight: 800; color: #e11d48; margin-bottom: 0.2rem; text-transform: uppercase; }
    .sub-header { color: #94a3b8; margin-bottom: 1.5rem; font-size: 0.95rem; }
    .stTextArea textarea { background-color: #1e293b; color: #ffffff; border: 1px solid #334155; border-radius: 8px; font-size: 14px; line-height: 1.5; }
    .stButton button { background-color: #e11d48; color: white; font-weight: 700; font-size: 1rem; border-radius: 8px; padding: 0.7rem 1.5rem; border: none; width: 100%; transition: all 0.2s; }
    .stButton button:hover { background-color: #be123c; color: white; transform: translateY(-1px); }
</style>
""", unsafe_allow_html=True)

st.markdown('<div class="main-header">⚡ CAZÉTV — DECUPADOR DE VT</div>', unsafe_allow_html=True)
st.markdown('<div class="sub-header">Cole a locução (OFF) abaixo para gerar a decupagem oficial fundamentada no histórico da CazéTV.</div>', unsafe_allow_html=True)

# Recupera chave da API
api_key = st.secrets.get("GEMINI_API_KEY", "")
if not api_key:
    with st.sidebar:
        st.subheader("⚙️ Configurações")
        api_key = st.text_input("Gemini API Key:", type="password")

# Lê a base de conhecimento real dos seus documentos
base_conhecimento = ""
if os.path.exists("exemplos_cazetv.txt"):
    with open("exemplos_cazetv.txt", "r", encoding="utf-8", errors="ignore") as f:
        base_conhecimento = f.read()

off_text = st.text_area("LOCUÇÃO / TEXTO DO OFF:", height=220, placeholder="Cole o texto corrido do OFF aqui...")

if off_text:
    palavras = len(off_text.split())
    segundos = int((palavras / 135) * 60)
    st.caption(f"📊 **{palavras} palavras** | Estimativa de locução: **~{segundos} segundos**")

if st.button("🎬 GERAR DECUPAGEM COM BASE NO HISTÓRICO CAZÉTV"):
    if not api_key:
        st.error("⚠️ Chave de API não configurada.")
    elif not off_text.strip():
        st.warning("⚠️ Insira o texto do OFF antes de gerar a decupagem.")
    else:
        with st.spinner("Estudando a base histórica de roteiros CazéTV, criando propostas e buscando links no YouTube..."):
            try:
                genai.configure(api_key=api_key)
                
                system_prompt = f"""
                VOCÊ É O DIRETOR CHEFE DE VTS E DECUPADOR DA CAZÉTV.
                Sua única função é receber uma locução (OFF) e produzir a DECUPAGEM TÉCNICA OFICIAL DA CAZÉTV.

                Abaixo está a sua BASE DE CONHECIMENTO COMPLETA — os roteiros e decupagens reais da CazéTV que servem de ensinamento e modelo de excelência.
                Você deve estudar rigorosamente o vocabulário, a métrica de corte, a linguagem de direção de imagem, as instruções sonoras (sobe som), os letterings estilizados e o formato da tabela de 3 colunas presentes nestes documentos:

                --- INÍCIO DA BASE DE CONHECIMENTO CAZÉTV ---
                {base_conhecimento}
                --- FIM DA BASE DE CONHECIMENTO CAZÉTV ---

                DIRETRIZES RÍGIDAS DE SAÍDA:
                1. Apresente o texto narrado no topo com suas divisões de ritmo.
                2. Em seguida, gere EXCLUSIVAMENTE a TABELA DE DECUPAGEM DE 3 COLUNAS no padrão exato dos documentos acima:
                   | PROPOSTA (IMAGEM / VÍDEO / ÁUDIO) | OFF | LINKS / OBSERVAÇÕES E TIMECODES |
                3. Na coluna LINKS: Forneça links diretos reais no YouTube com o marcador de segundo exato no player (formato: https://www.youtube.com/watch?v=...&t=XXs), com observações de edição fidedignas (crops, closes, sobe som).
                """
                
                # Lista de modelos modernos suportados em ordem de preferência
                modelos_disponiveis = ["gemini-2.5-flash", "gemini-2.0-flash", "gemini-flash-latest"]
                response = None
                
                for modelo in modelos_disponiveis:
                    try:
                        model = genai.GenerativeModel(
                            model_name=modelo,
                            generation_config={"temperature": 0.4},
                            system_instruction=system_prompt,
                            tools="google_search_retrieval"
                        )
                        response = model.generate_content(f"Decupe este novo OFF espelhando rigorosamente a qualidade e estilo dos exemplos da base de conhecimento:\n\n{off_text}")
                        if response:
                            break
                    except Exception:
                        continue

                if response:
                    st.success("✅ Decupagem concluída com sucesso!")
                    st.markdown("---")
                    st.markdown(response.text)
                    
                    st.download_button(
                        label="📥 Baixar Decupagem (.txt)",
                        data=response.text,
                        file_name="decupagem_cazetv.txt",
                        mime="text/plain"
                    )
                else:
                    st.error("Não foi possível processar com os modelos disponíveis. Verifique sua chave de API.")
                    
            except Exception as e:
                st.error(f"Erro ao processar: {e}")
