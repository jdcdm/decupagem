import streamlit as st
import google.generativeai as genai
import os
import re
import io
from youtube_search import YoutubeSearch
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
st.markdown('<div class="sub-header">Cole a locução (OFF) abaixo para gerar a decupagem oficial com download para Google Docs (.docx).</div>', unsafe_allow_html=True)

# Chave de API
api_key = st.secrets.get("GEMINI_API_KEY", "")
if not api_key:
    with st.sidebar:
        st.subheader("⚙️ Configurações")
        api_key = st.text_input("Gemini API Key:", type="password")

# Lê a base de conhecimento real
base_conhecimento = ""
if os.path.exists("exemplos_cazetv.txt"):
    with open("exemplos_cazetv.txt", "r", encoding="utf-8", errors="ignore") as f:
        base_conhecimento = f.read()

# Função que busca links reais no YouTube
def buscar_link_real_youtube(termo_busca, timecode_str):
    try:
        segundos = 0
        match_tc = re.search(r'(\d{1,2}):(\d{2})', timecode_str)
        if match_tc:
            minutos = int(match_tc.group(1))
            segs = int(match_tc.group(2))
            segundos = (minutos * 60) + segs

        results = YoutubeSearch(termo_busca, max_results=1).to_dict()
        if results:
            video = results[0]
            url = f"https://www.youtube.com{video['url_suffix']}"
            if segundos > 0:
                url += f"&t={segundos}s"
            titulo = video.get('title', termo_busca)
            return f"[{titulo}]({url})"
    except Exception:
        pass
    return f"**{termo_busca}**"

# Pós-processador para substituir tags de busca por links reais
def injetar_links_reais(texto_tabela):
    linhas = texto_tabela.split("\n")
    novas_linhas = []
    for linha in linhas:
        if "|" in linha and "BUSCA:" in linha:
            padrao = r'BUSCA:\s*\[(.*?)\]'
            match = re.search(padrao, linha)
            if match:
                termo = match.group(1)
                tc_match = re.search(r'\((\d{1,2}:\d{2}\s*-\s*\d{1,2}:\d{2})\)', linha)
                tc_str = tc_match.group(1) if tc_match else ""
                
                link_formatado = buscar_link_real_youtube(termo, tc_str)
                linha = re.sub(padrao, link_formatado, linha)
        novas_linhas.append(linha)
    return "\n".join(novas_linhas)

# Gerador oficial de documento DOCX / Google Docs
def gerar_documento_docx(off_completo, texto_resultado):
    doc = Document()
    
    # Margens do documento
    sections = doc.sections
    for section in sections:
        section.top_margin = Inches(0.8)
        section.bottom_margin = Inches(0.8)
        section.left_margin = Inches(0.8)
        section.right_margin = Inches(0.8)

    # Cabeçalho CazéTV
    h1 = doc.add_paragraph()
    r1 = h1.add_run("CAZÉTV — ROTEIRO E DECUPAGEM DE VT")
    r1.font.name = "Arial"
    r1.font.size = Pt(18)
    r1.font.bold = True
    r1.font.color.rgb = RGBColor(225, 29, 72) # Cor CazéTV
    h1.alignment = WD_ALIGN_PARAGRAPH.CENTER

    doc.add_paragraph("─" * 45).alignment = WD_ALIGN_PARAGRAPH.CENTER

    # Seção 1: Texto Corrido do OFF
    h2 = doc.add_paragraph()
    r2 = h2.add_run("1. LOCUÇÃO COMPLETA (OFF)")
    r2.font.name = "Arial"
    r2.font.size = Pt(13)
    r2.font.bold = True

    p_off = doc.add_paragraph()
    r_off = p_off.add_run(off_completo)
    r_off.font.name = "Georgia"
    r_off.font.size = Pt(11)
    r_off.font.italic = True
    p_off.paragraph_format.line_spacing = 1.3

    doc.add_paragraph()

    # Seção 2: Tabela de Decupagem
    h3 = doc.add_paragraph()
    r3 = h3.add_run("2. TABELA DE DECUPAGEM TÉCNICA")
    r3.font.name = "Arial"
    r3.font.size = Pt(13)
    r3.font.bold = True

    # Extrai as linhas da tabela markdown gerada
    linhas = texto_resultado.split("\n")
    dados_tabela = []
    
    for l in linhas:
        l_strip = l.strip()
        if l_strip.startswith("|") and not re.search(r'\|\s*:?---', l_strip):
            colunas = [c.strip() for c in l_strip.split("|")[1:-1]]
            if len(colunas) == 3:
                dados_tabela.append(colunas)

    if dados_tabela:
        # Cria a tabela no Word
        tabela = doc.add_table(rows=len(dados_tabela), cols=3)
        tabela.autofit = False

        # Larguras ideais das 3 colunas
        larguras = [Inches(3.0), Inches(1.8), Inches(2.7)]

        for i, row in enumerate(tabela.rows):
            colunas_texto = dados_tabela[i]
            for j, cell in enumerate(row.cells):
                cell.width = larguras[j]
                cell.paragraphs[0].text = colunas_texto[j]
                p = cell.paragraphs[0]
                p.paragraph_format.line_spacing = 1.15
                
                # Cabeçalho da tabela (Linha 0)
                if i == 0:
                    shading = parse_xml(r'<w:shd {} w:fill="111827"/>'.format(nsdecls('w')))
                    cell._tc.get_or_add_tcPr().append(shading)
                    for r in p.runs:
                        r.font.name = "Arial"
                        r.font.bold = True
                        r.font.size = Pt(10)
                        r.font.color.rgb = RGBColor(255, 255, 255)
                else:
                    # Linhas de dados
                    for r in p.runs:
                        r.font.name = "Arial"
                        r.font.size = Pt(9.5)

        # Adiciona bordas elegantes na tabela
        for row in tabela.rows:
            for cell in row.cells:
                tcPr = cell._tc.get_or_add_tcPr()
                tcBorders = parse_xml(
                    r'<w:tcBorders {}><w:top w:val="single" w:sz="4" w:space="0" w:color="D1D5DB"/>'
                    r'<w:bottom w:val="single" w:sz="4" w:space="0" w:color="D1D5DB"/>'
                    r'<w:left w:val="none"/><w:right w:val="none"/></w:tcBorders>'.format(nsdecls('w'))
                )
                tcPr.append(tcBorders)

    # Salva em memória para download imediato
    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer

off_text = st.text_area("LOCUÇÃO / TEXTO DO OFF:", height=220, placeholder="Cole o texto corrido do OFF aqui...")

if off_text:
    palavras = len(off_text.split())
    segundos = int((palavras / 135) * 60)
    st.caption(f"📊 **{palavras} palavras** | Estimativa de locução: **~{segundos} segundos**")

if st.button("🎬 GERAR DECUPAGEM COM LINKS REAIS"):
    if not api_key:
        st.error("⚠️ Chave de API não configurada.")
    elif not off_text.strip():
        st.warning("⚠️ Insira o texto do OFF antes de gerar a decupagem.")
    else:
        with st.spinner("Decupando roteiro, buscando vídeos reais e formatando para o Google Docs..."):
            try:
                genai.configure(api_key=api_key)
                
                modelos_validos = []
                try:
                    for m in genai.list_models():
                        if 'generateContent' in m.supported_generation_methods:
                            modelos_validos.append(m.name)
                except Exception:
                    pass

                if not modelos_validos:
                    modelos_validos = ["models/gemini-2.5-flash", "models/gemini-3.8-flash", "gemini-2.5-flash"]

                system_prompt = f"""
                VOCÊ É O DIRETOR CHEFE DE VTS E DECUPADOR DA CAZÉTV.
                Sua função é gerar a decupagem oficial da CazéTV a partir do OFF.

                Base de Conhecimento CazéTV:
                {base_conhecimento}

                REGRAS OBRIGATÓRIAS:
                1. Apresente o texto narrado no topo.
                2. Gere a TABELA DE 3 COLUNAS: | PROPOSTA | OFF | LINKS / OBSERVAÇÕES E TIMECODES |
                3. Na coluna de LINKS:
                   - Para imagens de arquivo ou YouTube, use OBRIGATORIAMENTE o formato:
                     BUSCA: [Título exato ou termos cirúrgicos do vídeo no YouTube] (MM:SS - MM:SS) - Observações de edição.
                   - NUNCA invente URLs com watch?v=. Use sempre a tag BUSCA: [nome do lance] que o nosso sistema buscará o link real.
                """

                prompt_final = f"Decupe este novo OFF:\n\n{off_text}"

                response = None
                for modelo in modelos_validos:
                    try:
                        model = genai.GenerativeModel(
                            model_name=modelo,
                            generation_config={"temperature": 0.4},
                            system_instruction=system_prompt
                        )
                        response = model.generate_content(prompt_final)
                        if response and response.text:
                            break
                    except Exception:
                        continue

                if response and response.text:
                    tabela_com_links_reais = injetar_links_reais(response.text)
                    
                    st.success("✅ Decupagem concluída com sucesso!")
                    st.markdown("---")
                    st.markdown(tabela_com_links_reais)
                    
                    # Gera o arquivo .docx (Google Docs / Word) em memória
                    docx_buffer = gerar_documento_docx(off_text, tabela_com_links_reais)
                    
                    st.download_button(
                        label="📄 BAIXAR EM FORMATO GOOGLE DOCS / WORD (.DOCX)",
                        data=docx_buffer,
                        file_name="decupagem_cazetv.docx",
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                    )
                else:
                    st.error("Não foi possível gerar com os modelos da sua conta.")
                    
            except Exception as e:
                st.error(f"Erro geral: {e}")
