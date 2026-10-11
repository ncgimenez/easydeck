import io
import json
import streamlit as st
from google import genai
from google.genai import types
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor

st.set_page_config(page_title="DuoDeck", layout="wide", page_icon="🪄")

st.title("🪄 DuoDeck - Générateur de Slides Différenciées")
st.markdown("Créez automatiquement vos supports en deux niveaux (adapté & avancé), modifiables dans PowerPoint, Slides ou Canva.")

# Récupération sécurisée et invisible de la clé API
api_key = st.secrets.get("GEMINI_API_KEY", None)

with st.sidebar:
    st.header("🎨 Personnalisation visuelle")
    bg_file = st.file_uploader("Image d'arrière-plan (optionnelle)", type=["png", "jpg", "jpeg"])
    st.info("💡 Sans image importée, un fond clair standard sera appliqué.")

# Formulaire principal
col1, col2 = st.columns([2, 1])
with col1:
    topic = st.text_area(
        "Sujet ou notions à aborder",
        placeholder="Ex : Les 4 saisons de l'année, distinguer les pièces de monnaie, le cycle de l'eau...",
        height=130
    )
with col2:
    nb_slides = st.slider("Nombre de diapositives de cours", min_value=3, max_value=8, value=4)

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
                        "description": "2 à 3 points concrets, phrases très courtes, repères visuels clairs avec émojis."
                    },
                    "advanced_content": {
                        "type": "ARRAY",
                        "items": {"type": "STRING"},
                        "description": "3 à 4 points détaillés avec explications et vocabulaire approfondi."
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
    prs.slide_width = Inches(13.333)  # Format 16:9
    prs.slide_height = Inches(7.5)
    blank_layout = prs.slide_layouts[6]

    bg_bytes = io.BytesIO(background_file.getvalue()) if background_file else None

    # Slide de titre
    slide = prs.slides.add_slide(blank_layout)
    if bg_bytes:
        bg_bytes.seek(0)
        slide.shapes.add_picture(bg_bytes, 0, 0, width=prs.slide_width, height=prs.slide_height)

    tx_box = slide.shapes.add_textbox(Inches(1.5), Inches(2.2), Inches(10.33), Inches(3.2))
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

    # Diapositives de contenu
    for item in data["slides"]:
        s = prs.slides.add_slide(blank_layout)
        if bg_bytes:
            bg_bytes.seek(0)
            s.shapes.add_picture(bg_bytes, 0, 0, width=prs.slide_width, height=prs.slide_height)

        box = s.shapes.add_textbox(Inches(1.2), Inches(1.2), Inches(11), Inches(5.2))
        tf_s = box.text_frame
        tf_s.word_wrap = True

        p_title = tf_s.paragraphs[0]
        p_title.text = item["slide_title"]
        p_title.font.size = Pt(32)
        p_title.font.bold = True
        p_title.font.color.rgb = RGBColor(15, 23, 42)
        p_title.space_after = Pt(20)

        for bullet in item[level_key]:
            p_bullet = tf_s.add_paragraph()
            p_bullet.text = bullet
            p_bullet.font.size = Pt(22)
            p_bullet.font.color.rgb = RGBColor(51, 65, 85)
            p_bullet.space_after = Pt(14)

    out = io.BytesIO()
    prs.save(out)
    out.seek(0)
    return out

# Génération
if st.button("✨ Générer les deux présentations", type="primary"):
    if not api_key:
        st.error("⚠️ La clé GEMINI_API_KEY est introuvable dans les Secrets de Streamlit.")
    elif not topic.strip():
        st.warning("Veuillez renseigner un sujet.")
    else:
        try:
            with st.spinner("Conception des deux séries pédagogiques en cours..."):
                client = genai.Client(api_key=api_key)
                sys_prompt = (
                    "Tu es un spécialiste de la différenciation pédagogique. Tu dois concevoir deux approches pour chaque diapositive :"
                    "1. 'adapted_content' : syntaxe claire, puces directes avec repères et émojis concrets, charge cognitive allégée."
                    "2. 'advanced_content' : notions complètes, phrases rédigées, vocabulaire technique approfondi."
                )
                prompt = f"Génère une présentation de {nb_slides} diapositives sur : {topic}"
                
                response = client.models.generate_content(
                    model="gemini-2.5-flash",
                    contents=prompt,
                    config=types.GenerateContentConfig(
                        system_instruction=sys_prompt,
                        response_mime_type="application/json",
                        response_schema=response_schema,
                        temperature=0.3
                    )
                )
                st.session_state["course_data"] = json.loads(response.text)
                st.success("Diapositives prêtes !")
        except Exception as e:
            st.error(f"Une erreur est survenue lors de la création : {e}")

# Affichage & Téléchargement
if "course_data" in st.session_state:
    data = st.session_state["course_data"]
    st.divider()
    st.subheader(f"📖 {data['title']}")
    
    tab1, tab2 = st.tabs(["🟢 Version Adaptée (Spécialisée)", "🔵 Version Avancée"])

    with tab1:
        for idx, s in enumerate(data["slides"], 1):
            st.markdown(f"**Slide {idx} : {s['slide_title']}**")
            for pt in s["adapted_content"]:
                st.write(f"- {pt}")
        pptx_1 = create_presentation(data, "adapted_content", bg_file)
        st.download_button(
            label="📥 Télécharger la version Adaptée (.pptx)",
            data=pptx_1,
            file_name=f"cours_{topic[:15].strip()}_adapte.pptx",
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
            key="dl_1"
        )

    with tab2:
        for idx, s in enumerate(data["slides"], 1):
            st.markdown(f"**Slide {idx} : {s['slide_title']}**")
            for pt in s["advanced_content"]:
                st.write(f"- {pt}")
        pptx_2 = create_presentation(data, "advanced_content", bg_file)
        st.download_button(
            label="📥 Télécharger la version Avancée (.pptx)",
            data=pptx_2,
            file_name=f"cours_{topic[:15].strip()}_avance.pptx",
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
            key="dl_2"
        )
