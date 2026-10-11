import io
import json
import streamlit as st
from google import genai
from google.genai import types
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor

st.set_page_config(page_title="DuoDeck", layout="wide")

st.title("🪄 DuoDeck - Générateur de Slides Différenciées")
st.markdown("Créez deux versions de votre support en un clic avec votre propre arrière-plan.")

# Barre latérale pour la clé et le fond
with st.sidebar:
    st.header("⚙️ Réglages")
    # Récupération automatique de la clé si enregistrée dans Streamlit, sinon champ libre
    api_key_env = st.secrets.get("GEMINI_API_KEY", "") if hasattr(st, "secrets") else ""
    api_key = st.text_input("Clé API Gemini", value=api_key_env, type="password", help="Obtenez une clé gratuite sur https://aistudio.google.com/")
    bg_file = st.file_uploader("Image d'arrière-plan (PNG ou JPG)", type=["png", "jpg", "jpeg"])

# Saisie du sujet
col1, col2 = st.columns([2, 1])
with col1:
    topic = st.text_area("Sujet / Notions à aborder", placeholder="Ex : Les différentes saisons de l'année...", height=120)
with col2:
    nb_slides = st.slider("Nombre de diapositives", min_value=3, max_value=8, value=4)

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
                        "description": "Phrases très courtes, simples, repères clairs avec émojis."
                    },
                    "advanced_content": {
                        "type": "ARRAY",
                        "items": {"type": "STRING"},
                        "description": "Phrases complètes, explications et vocabulaire détaillé."
                    }
                },
                "required": ["slide_title", "adapted_content", "advanced_content"]
            }
        }
    },
    "required": ["title", "slides"]
}

def create_presentation(data, level_key, background_file=None):
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    blank_layout = prs.slide_layouts[6]

    bg_bytes = None
    if background_file is not None:
        bg_bytes = io.BytesIO(background_file.getvalue())

    # Slide 1 : Titre
    slide = prs.slides.add_slide(blank_layout)
    if bg_bytes:
        bg_bytes.seek(0)
        slide.shapes.add_picture(bg_bytes, 0, 0, width=prs.slide_width, height=prs.slide_height)

    tx_box = slide.shapes.add_textbox(Inches(1.5), Inches(2.3), Inches(10.33), Inches(3))
    tf = tx_box.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = data["title"]
    p.font.size = Pt(44)
    p.font.bold = True
    p.font.color.rgb = RGBColor(30, 41, 59)

    p2 = tf.add_paragraph()
    level_label = "Niveau Adapté / Découverte" if level_key == "adapted_content" else "Niveau Approfondi"
    p2.text = f"— {level_label} —"
    p2.font.size = Pt(24)
    p2.font.color.rgb = RGBColor(100, 116, 139)

    # Slides suivantes
    for item in data["slides"]:
        s = prs.slides.add_slide(blank_layout)
        if bg_bytes:
            bg_bytes.seek(0)
            s.shapes.add_picture(bg_bytes, 0, 0, width=prs.slide_width, height=prs.slide_height)

        box = s.shapes.add_textbox(Inches(1.2), Inches(1.2), Inches(11), Inches(5))
        tf_s = box.text_frame
        tf_s.word_wrap = True

        p_title = tf_s.paragraphs[0]
        p_title.text = item["slide_title"]
        p_title.font.size = Pt(34)
        p_title.font.bold = True
        p_title.font.color.rgb = RGBColor(15, 23, 42)
        p_title.space_after = Pt(18)

        for bullet in item[level_key]:
            p_bullet = tf_s.add_paragraph()
            p_bullet.text = bullet
            p_bullet.font.size = Pt(22)
            p_bullet.font.color.rgb = RGBColor(51, 65, 85)
            p_bullet.space_after = Pt(12)

    out_stream = io.BytesIO()
    prs.save(out_stream)
    out_stream.seek(0)
    return out_stream

if st.button("✨ Générer les présentations", type="primary"):
    if not api_key:
        st.error("Ajoutez votre clé API Gemini dans le menu de gauche.")
    elif not topic.strip():
        st.warning("Veuillez indiquer un sujet.")
    else:
        try:
            with st.spinner("Création des deux versions..."):
                client = genai.Client(api_key=api_key)
                sys_instruct = (
                    "Tu es un expert pédagogique. Tu dois produire deux versions de chaque diapositive :"
                    "1. 'adapted_content' : très simple, phrases courtes, puces directes avec émojis explicites."
                    "2. 'advanced_content' : explications détaillées, enrichissement du vocabulaire."
                )
                prompt = f"Crée une présentation de {nb_slides} slides sur le thème : {topic}"
                response = client.models.generate_content(
                    model="gemini-2.5-flash",
                    contents=prompt,
                    config=types.GenerateContentConfig(
                        system_instruction=sys_instruct,
                        response_mime_type="application/json",
                        response_schema=response_schema,
                        temperature=0.3
                    )
                )
                st.session_state["course_data"] = json.loads(response.text)
                st.success("Présentation prête !")
        except Exception as e:
            st.error(f"Une erreur est survenue : {e}")

if "course_data" in st.session_state:
    data = st.session_state["course_data"]
    st.subheader(f"📖 {data['title']}")
    tab1, tab2 = st.tabs(["🟢 Version Adaptée", "🔵 Version Avancée"])

    with tab1:
        for idx, s in enumerate(data["slides"], 1):
            st.markdown(f"**Diapositive {idx} : {s['slide_title']}**")
            for pt in s["adapted_content"]:
                st.write(f"- {pt}")
        pptx_1 = create_presentation(data, "adapted_content", bg_file)
        st.download_button(
            label="📥 Télécharger la version Adaptée (.pptx)",
            data=pptx_1,
            file_name="slides_adaptees.pptx",
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
            key="dl_1"
        )

    with tab2:
        for idx, s in enumerate(data["slides"], 1):
            st.markdown(f"**Diapositive {idx} : {s['slide_title']}**")
            for pt in s["advanced_content"]:
                st.write(f"- {pt}")
        pptx_2 = create_presentation(data, "advanced_content", bg_file)
        st.download_button(
            label="📥 Télécharger la version Avancée (.pptx)",
            data=pptx_2,
            file_name="slides_avancees.pptx",
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
            key="dl_2"
        )
