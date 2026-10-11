import io
import json
import streamlit as st
from google import genai
from google.genai import types
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.shapes import MSO_SHAPE
from pptx.dml.color import RGBColor

st.set_page_config(page_title="DuoDeck", layout="centered", page_icon="🪄")

# --- CSS ÉPURÉ & MODERNE ---
st.markdown("""
<style>
    .block-container {
        padding-top: 2rem;
        padding-bottom: 3rem;
        max-width: 820px;
    }
    .main-header {
        text-align: center;
        margin-bottom: 1.8rem;
    }
    .main-title {
        font-size: 2.3rem;
        font-weight: 800;
        margin-bottom: 0.2rem;
    }
    .main-subtitle {
        font-size: 1rem;
        opacity: 0.8;
    }
    .card-preview {
        background-color: rgba(125, 125, 125, 0.08);
        border-radius: 12px;
        padding: 1.2rem;
        margin-bottom: 1rem;
        border: 1px solid rgba(125, 125, 125, 0.15);
    }
    .card-preview h4 {
        margin-top: 0;
        margin-bottom: 0.6rem;
    }
    div.stButton > button {
        border-radius: 10px;
        font-weight: bold;
        height: 3rem;
    }
</style>
""", unsafe_allow_html=True)

st.markdown("""
<div class="main-header">
    <div class="main-title">🪄 DuoDeck</div>
    <div class="main-subtitle">Générez des diapositives pédagogiques structurées en cartes visuelles.</div>
</div>
""", unsafe_allow_html=True)

api_key = st.secrets.get("GEMINI_API_KEY", None)

# --- 1. MODÈLE & STYLE ---
st.subheader("🎨 Modèle & Ambiance visuelle")
template_file = st.file_uploader(
    "Support d'inspiration (.pptx, optionnel)",
    type=["pptx"],
    help="DuoDeck analysera le style pédagogique et la formulation de ce modèle."
)

col_opt1, col_opt2 = st.columns(2)
with col_opt1:
    use_style_instruction = st.checkbox("Imiter le style d'écriture du modèle", value=True)
with col_opt2:
    color_scheme = st.selectbox(
        "Palette graphique",
        [
            "🪄 Couleurs créatives adaptées au sujet (par l'IA)",
            "Bleu Éducatif (Moderne & Sobre)",
            "Tons Chauds (Terracotta & Ambre)",
            "Vert Forêt (Calme & Nature)",
            "Violet / Indigo (Créatif)"
        ],
        index=0
    )

st.divider()

# --- 2. PARAMÈTRES DU COURS ---
st.subheader("📝 Paramètres de la leçon")
custom_title = st.text_input("Titre du cours (optionnel)", placeholder="Ex : Le cycle de l'eau dans la nature")
topic = st.text_area(
    "Sujet et notions clés à enseigner *",
    placeholder="Ex : L'évaporation, la condensation, les précipitations et le ruissellement. Préciser le rôle du soleil.",
    height=100
)

col_s1, col_s2 = st.columns([1, 1])
with col_s1:
    nb_slides = st.slider("Nombre de diapositives de contenu", min_value=2, max_value=8, value=3)
with col_s2:
    pedagogy_notes = st.text_input(
        "Consignes spécifiques",
        placeholder="Ex : 2 points clés par carte, émojis concrets, pas de jargon..."
    )

def extract_writing_style(pptx_bytes):
    try:
        prs = Presentation(io.BytesIO(pptx_bytes))
        samples = []
        for slide in prs.slides:
            for shape in slide.shapes:
                if shape.has_text_frame:
                    for p in shape.text_frame.paragraphs:
                        txt = p.text.strip()
                        if txt and len(txt) > 3:
                            samples.append(txt)
        return "\n".join(samples[:25])
    except Exception:
        return ""

# Schéma JSON incluant la génération automatique de palette de couleurs par Gemini
response_schema = {
    "type": "OBJECT",
    "properties": {
        "title": {"type": "STRING"},
        "palette": {
            "type": "OBJECT",
            "properties": {
                "primary": {"type": "STRING", "description": "Code HEX pour les titres (ex: #1E3A8A)"},
                "accent": {"type": "STRING", "description": "Code HEX vif pour les bordures et badges (ex: #3B82F6)"},
                "card_bg": {"type": "STRING", "description": "Code HEX très clair pour le fond des cartes (ex: #F1F5F9)"},
                "card_border": {"type": "STRING", "description": "Code HEX doux pour le contour (ex: #CBD5E1)"}
            },
            "required": ["primary", "accent", "card_bg", "card_border"]
        },
        "slides": {
            "type": "ARRAY",
            "items": {
                "type": "OBJECT",
                "properties": {
                    "slide_title": {"type": "STRING"},
                    "subtitle": {"type": "STRING"},
                    "cards": {
                        "type": "ARRAY",
                        "items": {
                            "type": "OBJECT",
                            "properties": {
                                "badge": {"type": "STRING"},
                                "card_title": {"type": "STRING"},
                                "points": {
                                    "type": "ARRAY",
                                    "items": {"type": "STRING"}
                                }
                            },
                            "required": ["badge", "card_title", "points"]
                        }
                    }
                },
                "required": ["slide_title", "subtitle", "cards"]
            }
        }
    },
    "required": ["title", "palette", "slides"]
}

# Palettes manuelles de secours
PRESET_PALETTES = {
    "Bleu Éducatif (Moderne & Sobre)": {
        "primary": RGBColor(30, 58, 138),
        "accent": RGBColor(59, 130, 246),
        "card_bg": RGBColor(241, 245, 249),
        "card_border": RGBColor(203, 213, 225),
        "text_dark": RGBColor(15, 23, 42),
        "text_muted": RGBColor(71, 85, 105)
    },
    "Tons Chauds (Terracotta & Ambre)": {
        "primary": RGBColor(154, 52, 18),
        "accent": RGBColor(217, 119, 6),
        "card_bg": RGBColor(254, 243, 199),
        "card_border": RGBColor(251, 191, 36),
        "text_dark": RGBColor(69, 26, 3),
        "text_muted": RGBColor(120, 53, 15)
    },
    "Vert Forêt (Calme & Nature)": {
        "primary": RGBColor(20, 83, 45),
        "accent": RGBColor(34, 197, 94),
        "card_bg": RGBColor(240, 253, 244),
        "card_border": RGBColor(187, 247, 208),
        "text_dark": RGBColor(5, 46, 22),
        "text_muted": RGBColor(22, 101, 52)
    },
    "Violet / Indigo (Créatif)": {
        "primary": RGBColor(76, 29, 149),
        "accent": RGBColor(139, 92, 246),
        "card_bg": RGBColor(245, 243, 255),
        "card_border": RGBColor(221, 214, 254),
        "text_dark": RGBColor(46, 16, 101),
        "text_muted": RGBColor(107, 33, 168)
    }
}

def hex_to_rgb(hex_code, fallback):
    try:
        clean = hex_code.lstrip("#")
        if len(clean) == 6:
            return RGBColor.from_string(clean)
    except Exception:
        pass
    return fallback

def resolve_palette(data, choice):
    if choice in PRESET_PALETTES:
        return PRESET_PALETTES[choice]
    
    # Palette créative générée par Gemini
    ai_pal = data.get("palette", {})
    return {
        "primary": hex_to_rgb(ai_pal.get("primary", ""), RGBColor(30, 58, 138)),
        "accent": hex_to_rgb(ai_pal.get("accent", ""), RGBColor(59, 130, 246)),
        "card_bg": hex_to_rgb(ai_pal.get("card_bg", ""), RGBColor(241, 245, 249)),
        "card_border": hex_to_rgb(ai_pal.get("card_border", ""), RGBColor(203, 213, 225)),
        "text_dark": RGBColor(15, 23, 42),
        "text_muted": RGBColor(71, 85, 105)
    }

# --- CRÉATION POWERPOINT ---
def build_card_presentation(data, level_name, palette_choice):
    pal = resolve_palette(data, palette_choice)

    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    blank_layout = prs.slide_layouts[6]

    # Slide 1 : Couverture
    s_cover = prs.slides.add_slide(blank_layout)
    banner = s_cover.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, prs.slide_width, Inches(0.4))
    banner.fill.solid()
    banner.fill.fore_color.rgb = pal["accent"]
    banner.line.color.rgb = pal["accent"]

    center_card = s_cover.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE,
        Inches(1.5), Inches(1.5), Inches(10.33), Inches(4.8)
    )
    center_card.fill.solid()
    center_card.fill.fore_color.rgb = pal["card_bg"]
    center_card.line.color.rgb = pal["card_border"]
    center_card.line.width = Pt(2)

    tb_cov = s_cover.shapes.add_textbox(Inches(2.0), Inches(2.2), Inches(9.33), Inches(3.5))
    tf_cov = tb_cov.text_frame
    tf_cov.word_wrap = True

    p_badge = tf_cov.paragraphs[0]
    p_badge.text = f"SUPPORT PÉDAGOGIQUE • {level_name.upper()}"
    p_badge.font.size = Pt(14)
    p_badge.font.bold = True
    p_badge.font.color.rgb = pal["accent"]
    p_badge.space_after = Pt(16)

    p_title = tf_cov.add_paragraph()
    p_title.text = data["title"]
    p_title.font.size = Pt(40)
    p_title.font.bold = True
    p_title.font.color.rgb = pal["primary"]
    p_title.space_after = Pt(16)

    p_desc = tf_cov.add_paragraph()
    p_desc.text = "Créé automatiquement avec DuoDeck"
    p_desc.font.size = Pt(18)
    p_desc.font.color.rgb = pal["text_muted"]

    # Slides de contenu
    for slide_data in data["slides"]:
        s = prs.slides.add_slide(blank_layout)

        top_bar = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, prs.slide_width, Inches(0.18))
        top_bar.fill.solid()
        top_bar.fill.fore_color.rgb = pal["accent"]
        top_bar.line.color.rgb = pal["accent"]

        header_box = s.shapes.add_textbox(Inches(0.8), Inches(0.4), Inches(11.73), Inches(1.2))
        tf_h = header_box.text_frame
        tf_h.word_wrap = True

        p_st = tf_h.paragraphs[0]
        p_st.text = slide_data["slide_title"]
        p_st.font.size = Pt(30)
        p_st.font.bold = True
        p_st.font.color.rgb = pal["primary"]

        if slide_data.get("subtitle"):
            p_sub = tf_h.add_paragraph()
            p_sub.text = slide_data["subtitle"]
            p_sub.font.size = Pt(16)
            p_sub.font.color.rgb = pal["text_muted"]
            p_sub.space_before = Pt(4)

        cards = slide_data.get("cards", [])
        nb_cards = max(1, len(cards))
        margin_x = Inches(0.8)
        gap_x = Inches(0.4)
        total_available_w = Inches(11.733) - (gap_x * (nb_cards - 1))
        card_w = total_available_w / nb_cards
        card_top = Inches(1.9)
        card_h = Inches(4.9)

        for i, card_info in enumerate(cards):
            card_left = margin_x + i * (card_w + gap_x)

            card_shape = s.shapes.add_shape(
                MSO_SHAPE.ROUNDED_RECTANGLE,
                card_left, card_top, card_w, card_h
            )
            card_shape.fill.solid()
            card_shape.fill.fore_color.rgb = pal["card_bg"]
            card_shape.line.color.rgb = pal["card_border"]
            card_shape.line.width = Pt(1.5)

            pad = Inches(0.25)
            content_box = s.shapes.add_textbox(
                card_left + pad, card_top + pad, card_w - (pad * 2), card_h - (pad * 2)
            )
            tf_c = content_box.text_frame
            tf_c.word_wrap = True

            p_c_badge = tf_c.paragraphs[0]
            p_c_badge.text = card_info.get("badge", "").upper()
            p_c_badge.font.size = Pt(12)
            p_c_badge.font.bold = True
            p_c_badge.font.color.rgb = pal["accent"]
            p_c_badge.space_after = Pt(8)

            p_c_title = tf_c.add_paragraph()
            p_c_title.text = card_info.get("card_title", "")
            p_c_title.font.size = Pt(20)
            p_c_title.font.bold = True
            p_c_title.font.color.rgb = pal["primary"]
            p_c_title.space_after = Pt(12)

            for point in card_info.get("points", []):
                p_c_pt = tf_c.add_paragraph()
                p_c_pt.text = f"•  {point}"
                p_c_pt.font.size = Pt(16)
                p_c_pt.font.color.rgb = pal["text_dark"]
                p_c_pt.space_after = Pt(8)

    output = io.BytesIO()
    prs.save(output)
    output.seek(0)
    return output

# --- GÉNÉRATION ---
def run_generation(level_type):
    if not api_key:
        st.error("⚠️ Clé GEMINI_API_KEY introuvable dans les Secrets de Streamlit.")
        return
    if not topic.strip():
        st.warning("Veuillez renseigner un sujet de cours.")
        return

    progress_bar = st.progress(0)
    status_text = st.empty()

    try:
        status_text.markdown("*Étape 1/3 : Analyse du sujet et détermination de la palette adaptée...*")
        progress_bar.progress(20)

        t_bytes = template_file.getvalue() if template_file else None
        style_guide = ""
        if t_bytes and use_style_instruction:
            samples = extract_writing_style(t_bytes)
            if samples:
                style_guide = (
                    "\n\nVoici des exemples du style d'écriture de l'enseignant. "
                    "Imite son ton et sa clarté sans reprendre ses anciens sujets :\n"
                    f"\"\"\"\n{samples}\n\"\"\""
                )

        is_adapted = (level_type == "adapted")
        level_label = "Niveau Adapté / Découverte" if is_adapted else "Niveau Approfondi"

        pedagogy_target = (
            "POUR CETTE VERSION ADAPTÉE : Syntaxe simple, phrases courtes, repères très concrets avec émojis. "
            "Divise le contenu en 2 cartes très lisibles par diapositive, chaque carte ayant 1 ou 2 phrases claires."
            if is_adapted else
            "POUR CETTE VERSION APPROFONDIE : Contenu enrichi, explications complètes, vocabulaire précis. "
            "Divise le contenu en 2 ou 3 cartes détaillées par diapositive, chacune contenant 2 à 3 points riches."
        )

        title_rule = f"Titre imposé : {custom_title}" if custom_title.strip() else "Trouve un titre pédagogique clair et motivant."
        needs_rule = f"Consignes pédagogiques spécifiques : {pedagogy_notes}" if pedagogy_notes.strip() else ""

        color_instruction = (
            "Pour le champ 'palette', choisis 4 couleurs HEX adaptées au thème abordé "
            "(par exemple des tons ambrés/rouges pour le feu ou l'histoire, des verts pour la nature/SVT, des bleus pour l'eau ou la science, etc.). "
            "Garantis un excellent contraste : 'primary' foncé, 'accent' vif, 'card_bg' très clair pastel, 'card_border' douce."
        )

        sys_prompt = (
            f"Tu es un expert en différenciation pédagogique et en design graphique.\n"
            f"{pedagogy_target}\n"
            f"{color_instruction}\n"
            f"{style_guide}"
        )

        prompt = (
            f"{title_rule}\n"
            f"Sujet du cours : {topic}\n"
            f"{needs_rule}\n"
            f"Génère exactement {nb_slides} diapositives avec leur palette et leurs cartes structurées."
        )

        status_text.markdown("*Étape 2/3 : Rédaction et composition visuelle par Gemini...*")
        progress_bar.progress(60)

        client = genai.Client(api_key=api_key)
        response = None
        last_err = None

        for mdl in ["gemini-2.5-flash", "gemini-2.0-flash"]:
            try:
                response = client.models.generate_content(
                    model=mdl,
                    contents=prompt,
                    config=types.GenerateContentConfig(
                        system_instruction=sys_prompt,
                        response_mime_type="application/json",
                        response_schema=response_schema,
                        temperature=0.3
                    )
                )
                if response and response.text:
                    break
            except Exception as err:
                last_err = err
                continue

        if not response or not response.text:
            raise last_err

        status_text.markdown("*Étape 3/3 : Création d'un PowerPoint neuf avec les couleurs thématiques...*")
        progress_bar.progress(90)

        data = json.loads(response.text)
        pptx_bytes = build_card_presentation(data, level_label, color_scheme)

        st.session_state[f"deck_{level_type}"] = {
            "data": data,
            "pptx": pptx_bytes,
            "level_label": level_label
        }

        progress_bar.progress(100)
        status_text.markdown("*Terminé !*")
        st.success(f"🎉 Présentation {level_label} générée avec succès !")

    except Exception as e:
        progress_bar.empty()
        status_text.empty()
        st.error(f"Une erreur est survenue lors de la génération : {e}")

# --- BOUTONS ---
col_btn1, col_btn2 = st.columns(2)
with col_btn1:
    if st.button("🟢 Générer la version Adaptée", type="primary", use_container_width=True):
        run_generation("adapted")

with col_btn2:
    if st.button("🔵 Générer la version Approfondie", type="secondary", use_container_width=True):
        run_generation("advanced")

# --- PRÉVISUALISATION & TÉLÉCHARGEMENT ---
has_adapted = "deck_adapted" in st.session_state
has_advanced = "deck_advanced" in st.session_state

if has_adapted or has_advanced:
    st.divider()
    tabs_labels = []
    if has_adapted:
        tabs_labels.append("🟢 Version Adaptée")
    if has_advanced:
        tabs_labels.append("🔵 Version Approfondie")

    tabs = st.tabs(tabs_labels)
    tab_i = 0

    if has_adapted:
        with tabs[tab_i]:
            deck = st.session_state["deck_adapted"]
            st.markdown(f"### 📖 {deck['data']['title']}")
            for idx, s in enumerate(deck["data"]["slides"], 1):
                cards_html = ""
                for c in s.get("cards", []):
                    pts = "".join([f"<li>{p}</li>" for p in c.get("points", [])])
                    cards_html += f"""
                    <div style="background: rgba(16, 185, 129, 0.08); border-left: 3px solid #10b981; padding: 10px 14px; margin-top: 8px; border-radius: 6px;">
                        <small style="color: #10b981; font-weight: bold;">{c.get('badge', '')}</small>
                        <div style="font-weight: bold; margin: 4px 0;">{c.get('card_title', '')}</div>
                        <ul style="margin: 0; padding-left: 20px;">{pts}</ul>
                    </div>
                    """
                st.markdown(f"""
                <div class="card-preview">
                    <h4>Diapositive {idx} : {s['slide_title']}</h4>
                    <p style="opacity: 0.8; margin-bottom: 6px;"><em>{s.get('subtitle', '')}</em></p>
                    {cards_html}
                </div>
                """, unsafe_allow_html=True)

            st.download_button(
                "📥 Télécharger la présentation Adaptée (.pptx)",
                deck["pptx"],
                file_name="cours_adapte.pptx",
                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                key="dl_adapted"
            )
        tab_i += 1

    if has_advanced:
        with tabs[tab_i]:
            deck = st.session_state["deck_advanced"]
            st.markdown(f"### 📖 {deck['data']['title']}")
            for idx, s in enumerate(deck["data"]["slides"], 1):
                cards_html = ""
                for c in s.get("cards", []):
                    pts = "".join([f"<li>{p}</li>" for p in c.get("points", [])])
                    cards_html += f"""
                    <div style="background: rgba(59, 130, 246, 0.08); border-left: 3px solid #3b82f6; padding: 10px 14px; margin-top: 8px; border-radius: 6px;">
                        <small style="color: #3b82f6; font-weight: bold;">{c.get('badge', '')}</small>
                        <div style="font-weight: bold; margin: 4px 0;">{c.get('card_title', '')}</div>
                        <ul style="margin: 0; padding-left: 20px;">{pts}</ul>
                    </div>
                    """
                st.markdown(f"""
                <div class="card-preview">
                    <h4>Diapositive {idx} : {s['slide_title']}</h4>
                    <p style="opacity: 0.8; margin-bottom: 6px;"><em>{s.get('subtitle', '')}</em></p>
                    {cards_html}
                </div>
                """, unsafe_allow_html=True)

            st.download_button(
                "📥 Télécharger la présentation Approfondie (.pptx)",
                deck["pptx"],
                file_name="cours_avance.pptx",
                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                key="dl_advanced"
            )
