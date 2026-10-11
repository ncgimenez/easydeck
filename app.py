import io
import json
import copy
import streamlit as st
from google import genai
from google.genai import types
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor

st.set_page_config(page_title="DuoDeck", layout="centered", page_icon="🪄")

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
    div.stButton > button:first-child {
        width: 100%;
        border-radius: 12px;
        height: 3.2rem;
        font-size: 1.1rem;
        font-weight: bold;
        margin-top: 1rem;
    }
</style>
""", unsafe_allow_html=True)

st.markdown("""
<div class="main-header">
    <div class="main-title">🪄 DuoDeck</div>
    <div class="main-subtitle">Générez vos diapositives adaptées et approfondies à partir de vos propres modèles.</div>
</div>
""", unsafe_allow_html=True)

api_key = st.secrets.get("GEMINI_API_KEY", None)

# --- 1. MODÈLE MAÎTRE ---
st.subheader("🎨 Votre modèle PowerPoint (.pptx)")
template_file = st.file_uploader(
    "Déposez votre présentation PowerPoint modèle",
    type=["pptx"],
    help="DuoDeck conservera vos polices, vos couleurs, vos arrières-plans et votre disposition."
)

col_opt1, col_opt2 = st.columns(2)
with col_opt1:
    use_style_instruction = st.checkbox("Imiter le style d'écriture de ce modèle", value=True)
with col_opt2:
    color_scheme = st.selectbox(
        "Nuance des titres & textes",
        ["Conserver les couleurs d'origine", "Bleu nuit & Ardoise", "Tons chauds (Terracotta)", "Vert sauge (Apaisant)"]
    )

st.divider()

# --- 2. LE NOUVEAU COURS ---
st.subheader("📝 Paramètres de la leçon")
custom_title = st.text_input("Titre du cours (optionnel)", placeholder="Ex : Les étapes du cycle de l'eau")
topic = st.text_area(
    "Sujet et notions clés à enseigner *",
    placeholder="Ex : Évaporation, condensation, précipitations, ruissellement. Expliquer le rôle du soleil.",
    height=100
)

col_s1, col_s2 = st.columns([1, 1])
with col_s1:
    nb_slides = st.slider("Nombre de diapositives", min_value=2, max_value=8, value=4)
with col_s2:
    pedagogy_notes = st.text_input(
        "Consignes spécifiques de structure",
        placeholder="Ex : 2 phrases max par slide, émojis explicites, vocabulaire simple..."
    )

# --- EXTRACTION DU STYLE RÉDACTIONNEL DU MODÈLE ---
def extract_writing_style(pptx_bytes):
    try:
        prs = Presentation(io.BytesIO(pptx_bytes))
        sample_texts = []
        for slide in prs.slides:
            for shape in slide.shapes:
                if shape.has_text_frame:
                    for p in shape.text_frame.paragraphs:
                        txt = p.text.strip()
                        if txt and len(txt) > 3:
                            sample_texts.append(txt)
        return "\n".join(sample_texts[:25])
    except Exception:
        return ""

# Schéma JSON
response_schema = {
    "type": "OBJECT",
    "properties": {
        "title": {"type": "STRING"},
        "slides": {
            "type": "ARRAY",
            "items": {
                "type": "OBJECT",
                "properties": {
                    "slide_title": {"type": "STRING"},
                    "adapted_content": {
                        "type": "ARRAY",
                        "items": {"type": "STRING"},
                        "description": "Phrases très courtes, simples, repères clairs avec émojis pour besoins spécifiques."
                    },
                    "advanced_content": {
                        "type": "ARRAY",
                        "items": {"type": "STRING"},
                        "description": "Notions plus riches, phrases complètes et vocabulaire approfondi."
                    }
                },
                "required": ["slide_title", "adapted_content", "advanced_content"]
            }
        }
    },
    "required": ["title", "slides"]
}

# --- GÉNÉRATION EN RÉUTILISANT LE TEMPLATE PPTX ---
def apply_color_override(paragraph, choice):
    if choice == "Bleu nuit & Ardoise":
        paragraph.font.color.rgb = RGBColor(30, 41, 59)
    elif choice == "Tons chauds (Terracotta)":
        paragraph.font.color.rgb = RGBColor(154, 52, 18)
    elif choice == "Vert sauge (Apaisant)":
        paragraph.font.color.rgb = RGBColor(47, 79, 79)

def generate_from_template(template_bytes, data, level_key, color_choice):
    # Si l'utilisateur n'a pas mis de modèle, on crée une présentation vierge 16:9
    if template_bytes is None:
        prs = Presentation()
        prs.slide_width = Inches(13.333)
        prs.slide_height = Inches(7.5)
        blank = prs.slide_layouts[6]
        
        # Titre
        s_title = prs.slides.add_slide(blank)
        tb = s_title.shapes.add_textbox(Inches(1.5), Inches(2.2), Inches(10.33), Inches(3))
        p = tb.text_frame.paragraphs[0]
        p.text = data["title"]
        p.font.size = Pt(44)
        p.font.bold = True
        apply_color_override(p, color_choice)
        
        # Diapositives de cours
        for item in data["slides"]:
            s = prs.slides.add_slide(blank)
            box = s.shapes.add_textbox(Inches(1.2), Inches(1.2), Inches(11), Inches(5))
            tf = box.text_frame
            tf.word_wrap = True
            
            p_title = tf.paragraphs[0]
            p_title.text = item["slide_title"]
            p_title.font.size = Pt(32)
            p_title.font.bold = True
            p_title.space_after = Pt(16)
            apply_color_override(p_title, color_choice)

            for b in item[level_key]:
                pb = tf.add_paragraph()
                pb.text = b
                pb.font.size = Pt(22)
                pb.space_after = Pt(12)
                if color_choice != "Conserver les couleurs d'origine":
                    pb.font.color.rgb = RGBColor(71, 85, 105)
    else:
        # Réutilisation du fichier modèle fourni
        prs = Presentation(io.BytesIO(template_bytes))
        
        # On repère les slides types du modèle (slide 1 = couverture, slide 2 = contenu si elle existe)
        title_layout = prs.slides[0].slide_layout if len(prs.slides) > 0 else prs.slide_layouts[0]
        content_layout = prs.slides[1].slide_layout if len(prs.slides) > 1 else title_layout

        # On injecte le titre sur la slide 1 existante ou nouvelle
        if len(prs.slides) > 0:
            slide_1 = prs.slides[0]
            for shape in slide_1.shapes:
                if shape.has_text_frame and shape.text_frame.text.strip():
                    shape.text_frame.text = data["title"]
                    if color_choice != "Conserver les couleurs d'origine":
                        for p in shape.text_frame.paragraphs:
                            apply_color_override(p, color_choice)
                    break
        
        # Ajout des diapositives en respectant le template
        for item in data["slides"]:
            new_slide = prs.slides.add_slide(content_layout)
            
            # Recherche des zones de titre et corps préexistantes dans le layout
            title_set = False
            content_set = False
            
            for shape in new_slide.shapes:
                if shape.has_text_frame:
                    tf = shape.text_frame
                    if not title_set:
                        tf.text = item["slide_title"]
                        title_set = True
                        if color_choice != "Conserver les couleurs d'origine":
                            for p in tf.paragraphs:
                                apply_color_override(p, color_choice)
                    elif not content_set:
                        tf.text = ""
                        for idx, bullet_text in enumerate(item[level_key]):
                            p = tf.add_paragraph() if idx > 0 else tf.paragraphs[0]
                            p.text = bullet_text
                            p.font.size = Pt(22)
                            p.space_after = Pt(10)
                        content_set = True
                        break
            
            # S'il n'y avait pas de placeholder texte, on pose un bloc propre
            if not title_set or not content_set:
                box = new_slide.shapes.add_textbox(Inches(1.2), Inches(1.2), Inches(11), Inches(5))
                tf = box.text_frame
                tf.word_wrap = True
                p_t = tf.paragraphs[0]
                p_t.text = item["slide_title"]
                p_t.font.size = Pt(32)
                p_t.font.bold = True
                p_t.space_after = Pt(14)
                apply_color_override(p_t, color_choice)
                
                for bullet_text in item[level_key]:
                    pb = tf.add_paragraph()
                    pb.text = bullet_text
                    pb.font.size = Pt(22)
                    pb.space_after = Pt(10)

    out = io.BytesIO()
    prs.save(out)
    out.seek(0)
    return out

# --- GÉNÉRATION ---

if st.button("✨ Générer mes deux présentations", type="primary"):
    if not api_key:
        st.error("⚠️ Clé GEMINI_API_KEY introuvable dans les Secrets de Streamlit.")
    elif not topic.strip():
        st.warning("Veuillez indiquer un sujet de cours.")
    else:
        try:
            with st.spinner("Conception pédagogique basée sur votre modèle..."):
                t_bytes = template_file.getvalue() if template_file else None
                st.session_state["raw_template"] = t_bytes
                
                style_guide = ""
                if t_bytes and use_style_instruction:
                    extracted_samples = extract_writing_style(t_bytes)
                    if extracted_samples:
                        style_guide = (
                            "\n\nVoici des exemples tirés du modèle de l'enseignant. "
                            "Imite UNIQUEMENT la tournure de phrases, le niveau de langage et la structure (ne reprends pas les anciens sujets) :\n"
                            f"\"\"\"\n{extracted_samples}\n\"\"\""
                        )

                title_rule = f"Titre imposé : {custom_title}" if custom_title.strip() else "Trouve un titre pédagogique adapté."
                needs_rule = f"Consignes pédagogiques spécifiques : {pedagogy_notes}" if pedagogy_notes.strip() else ""

                sys_prompt = (
                    "Tu es un expert en pédagogie et en différenciation d'enseignement. "
                    "Tu produis deux niveaux pour chaque slide :\n"
                    "1. 'adapted_content' : syntaxe directe, phrases très courtes, repères et émojis concrets, charge cognitive allégée.\n"
                    "2. 'advanced_content' : notions complètes, phrases élaborées, vocabulaire approfondi.\n"
                    f"{style_guide}"
                )

                prompt = f"{title_rule}\nSujet du nouveau cours : {topic}\n{needs_rule}\nGénère {nb_slides} diapositives."

                client = genai.Client(api_key=api_key)
                response = client.models.generate_content(
                    model="gemini-3.8-flash",                    
                    contents=prompt,
                    config=types.GenerateContentConfig(
                        system_instruction=sys_prompt,
                        response_mime_type="application/json",
                        response_schema=response_schema,
                        temperature=0.3
                    )
                )

                st.session_state["course_data"] = json.loads(response.text)
                st.success("Présentations créées à partir de votre modèle !")
        except Exception as e:
            st.error(f"Une erreur est survenue : {e}")

# --- AFFICHAGE & TÉLÉCHARGEMENT ---
if "course_data" in st.session_state:
    data = st.session_state["course_data"]
    saved_template = st.session_state.get("raw_template", None)

    st.divider()
    st.subheader(f"📖 {data['title']}")

    t1, t2 = st.tabs(["🟢 Version Adaptée", "🔵 Version Avancée"])

    with t1:
        for idx, s in enumerate(data["slides"], 1):
            st.markdown(f"**Slide {idx} : {s['slide_title']}**")
            for pt in s["adapted_content"]:
                st.write(f"- {pt}")
        pptx_out1 = generate_from_template(saved_template, data, "adapted_content", color_scheme)
        st.download_button(
            "📥 Télécharger la présentation Adaptée (.pptx)",
            pptx_out1,
            file_name="cours_adapte.pptx",
            key="dl_1"
        )

    with t2:
        for idx, s in enumerate(data["slides"], 1):
            st.markdown(f"**Slide {idx} : {s['slide_title']}**")
            for pt in s["advanced_content"]:
                st.write(f"- {pt}")
        pptx_out2 = generate_from_template(saved_template, data, "advanced_content", color_scheme)
        st.download_button(
            "📥 Télécharger la présentation Avancée (.pptx)",
            pptx_out2,
            file_name="cours_avance.pptx",
            key="dl_2"
        )
