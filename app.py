import streamlit as st
import google.generativeai as genai
import os
import re
import io
import json
from youtube_search import YoutubeSearch
import docx
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls

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
st.markdown('<div class="sub-header">Padrão Oficial: Links reais e canais oficiais verificados no YouTube, arquivos de fotos e download formatado para Google Docs (.docx).</div>', unsafe_allow_html=True)

api_key = st.secrets.get("GEMINI_API_KEY", "")
if not api_key:
    with st.sidebar:
        st.subheader("⚙️ Configurações")
        api_key = st.text_input("Gemini API Key:", type="password")

base_conhecimento = ""
if os.path.exists("exemplos_cazetv.txt"):
    with open("exemplos_cazetv.txt", "r", encoding="utf-8", errors="ignore") as f:
        base_conhecimento = f.read()

# Adiciona Hiperlink Real no Word / Google Docs
def add_hyperlink(paragraph, url, text, color="E11D48"):
    try:
        part = paragraph.part
        r_id = part.relate_to(url, docx.opc.constants.RELATIONSHIP_TYPE.HYPERLINK, is_external=True)
        hyperlink = parse_xml(
            f'<w:hyperlink xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" r:id="{r_id}" w:history="1">'
            f'<w:r>'
            f'<w:rPr><w:color w:val="{color}"/><w:u w:val="single"/><w:b/></w:rPr>'
            f'<w:t>{text}</w:t>'
            f'</w:r>'
            f'</w:hyperlink>'
        )
        paragraph._p.append(hyperlink)
    except Exception:
        r = paragraph.add_run(f"{text} ({url})")
        r.font.color.rgb = RGBColor(225, 29, 72)
        r.font.underline = True

CANAIS_OFICIAIS = [
    "uefa", "fifa", "barcelona", "barça", "real madrid", "psg", 
    "paris saint-germain", "ligue 1", "movistar", "espn", "tnt sports",
    "losc", "marseille", "wec", "fia", "rfeif", "olympic"
]

def buscar_video_oficial_youtube(termo_busca, timecode_str):
    try:
        segundos = 0
        match_tc = re.search(r'(\d{1,2}):(\d{2})', timecode_str)
        if match_tc:
            minutos = int(match_tc.group(1))
            segs = int(match_tc.group(2))
            segundos = (minutos * 60) + segs

        results = YoutubeSearch(termo_busca, max_results=5).to_dict()
        if results:
            video_escolhido = results[0]
            # Prioriza canal oficial se encontrado nos 5 primeiros resultados
            for v in results:
                canal = v.get("channel", "").lower()
                titulo = v.get("title", "").lower()
                if any(oficial in canal or oficial in titulo for oficial in CANAIS_OFICIAIS):
                    video_escolhido = v
                    break

            url = f"https://www.yout
