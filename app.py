import io
import json
import streamlit as st
from google import genai
from google.genai import types
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pypdf import PdfReader

st.set_page_config(page_title="DuoDeck", layout="centered", page_icon="🪄")

# CSS adapté au mode clair et sombre
st.markdown("""
<style>
    .block-container {
        padding-top: 2rem;
        padding-bottom: 3rem;
        max-width: 820px;
    }
    .main-header {
        text-align: center;
        margin-bottom: 2rem;
    }
    .main-title {
        font-size: 2.3rem;
        font-weight: 800;
        margin-bottom: 0.3rem;
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
    <div class="main-subtitle">Générez deux versions différenciées (adaptée & avancée) en un clic.</div>
</div>
""", unsafe_allow_html=True)

api_key = st.secrets.get("GEMINI_API_KEY", None)

with st.expander("🎨 Personnalisation visuelle (Arrière-plan)", expanded=False):
    bg_file = st.file_uploader("Image de fond pour les diapositives", type=["png", "jpg", "jpeg"])

with st.expander("📚 S'inspirer d'un ancien cours (Optionnel)", expanded=False):
    old_course_file = st.file_uploader("Importer un ancien support (.pptx, .pdf, .txt)", type=["pptx", "pdf", "txt"])
    old_course_text_input = st.text_area("Ou collez du texte d'exemple", placeholder="Collez ici des éléments d'un ancien cours...", height=80)

st.subheader("📝 Paramètres de la leçon")
custom_title = st.text_input("Titre du cours (optionnel)", placeholder="Ex : La découverte des 4 saisons")
topic = st.text_area("Sujet et notions clés *", placeholder="Ex : Le passage de l'automne à l'hiver, le froid, la neige...", height=110)

c1, c2 = st.columns([1, 1])
with c1:
    nb_slides = st.slider("Nombre de diapositives", min_value=3, max_value=8, value=4)
with c2:
    pedagogy_notes = st.text_input("Besoins spécifiques / Consignes", placeholder="Ex : Mots simples, 2 puces max, émojis...")

def extract_text_from_file(uploaded_file):
    if uploaded_file is None:
        return ""
    try:
        if uploaded_file.name.endswith(".txt"):
            return uploaded_file.getvalue().decode("utf-8")
        elif uploaded_file.name.endswith(".pdf"):
            reader = PdfReader(uploaded_file)
            return "\n".join([page.extract_text() or "" for page in reader.pages])
        elif uploaded_file.name.endswith(".pptx"):
            prs = Presentation(uploaded_file)
            runs = []
            for slide in prs.slides:
                for shape in slide.shapes:
                    if shape.has_text_frame:
                        for p in shape.text_frame.paragraphs:
                            runs.append(p.text)
            return "\n".join(runs)
    except Exception:
        return ""
    return ""

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
                    "adapted_content": {"type": "ARRAY", "items": {"type": "STRING"}},
                    "advanced_content": {"type": "ARRAY", "items": {"type": "STRING"}}
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
    bg_bytes = io.BytesIO(background_file.getvalue()) if background_file else None

    # Titre
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
    p2.text = f"— {'Niveau Découverte / Adapté' if level_key == 'adapted_content' else 'Niveau Approfondi'} —"
    p2.font.size = Pt(24)
    p2.font.color.rgb = RGBColor(100, 116, 139)

    # Contenu
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
            pb = tf_s.add_paragraph()
            pb.text = bullet
            pb.font.size = Pt(22)
            pb.font.color.rgb = RGBColor(51, 65, 85)
            pb.space_after = Pt(14)

    out = io.BytesIO()
    prs.save(out)
    out.seek(0)
    return out

if st.button("✨ Créer mes deux présentations", type="primary"):
    if not api_key:
        st.error("⚠️ Clé GEMINI_API_KEY introuvable dans les Secrets de Streamlit.")
    elif not topic.strip():
        st.warning("Veuillez saisir un sujet.")
    else:
        try:
            with st.spinner("Génération pédagogique en cours..."):
                extracted = extract_text_from_file(old_course_file)
                combined_ref = f"{extracted}\n{old_course_text_input}".strip()
                ref_context = f"\n\nExemple de modèle :\n\"\"\"\n{combined_ref[:4000]}\n\"\"\"" if combined_ref else ""
                
                title_rule = f"Titre : {custom_title}" if custom_title.strip() else "Génère un titre adapté."
                needs_rule = f"Besoins spécifiques : {pedagogy_notes}" if pedagogy_notes.strip() else ""

                sys_prompt = (
                    "Tu es un expert en pédagogie adaptée. Génère 2 approches par slide :\n"
                    "1. 'adapted_content' : très simple, émojis concrets, repères clairs.\n"
                    "2. 'advanced_content' : phrases complètes, vocabulaire riche.\n"
                    f"{ref_context}"
                )
                prompt = f"{title_rule}\nSujet : {topic}\n{needs_rule}\n{nb_slides} diapositives."

                client = genai.Client(api_key=api_key)
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
                st.success("Terminé !")
        except Exception as e:
            st.error(f"Erreur : {e}")

if "course_data" in st.session_state:
    data = st.session_state["course_data"]
    st.divider()
    st.subheader(f"📖 {data['title']}")
    t1, t2 = st.tabs(["🟢 Version Adaptée", "🔵 Version Avancée"])
    with t1:
        for idx, s in enumerate(data["slides"], 1):
            st.markdown(f"**Slide {idx} : {s['slide_title']}**")
            for pt in s["adapted_content"]:
                st.write(f"- {pt}")
        st.download_button("📥 Télécharger (.pptx)", create_presentation(data, "adapted_content", bg_file), "cours_adapte.pptx", key="dl_1")
    with t2:
        for idx, s in enumerate(data["slides"], 1):
            st.markdown(f"**Slide {idx} : {s['slide_title']}**")
            for pt in s["advanced_content"]:
                st.write(f"- {pt}")
        st.download_button("📥 Télécharger (.pptx)", create_presentation(data, "advanced_content", bg_file), "cours_avance.pptx", key="dl_2")
