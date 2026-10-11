import io
import json
import streamlit as st
from google import genai
from google.genai import types
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pypdf import PdfReader

st.set_page_config(page_title="DuoDeck", layout="wide", page_icon="🪄")

st.title("🪄 DuoDeck - Générateur de Slides Différenciées")
st.markdown("Créez des présentations adaptées et différenciées, guidées par vos anciens cours et vos consignes.")

# Clé API stockée de façon sécurisée
api_key = st.secrets.get("GEMINI_API_KEY", None)

# --- MENU LATÉRAL : VISUEL & EXEMPLE ANCIEN COURS ---
with st.sidebar:
    st.header("🎨 Visuel")
    bg_file = st.file_uploader("Image d'arrière-plan (PNG ou JPG)", type=["png", "jpg", "jpeg"])
    
    st.divider()
    st.header("📚 S'inspirer d'un ancien cours")
    st.caption("Donnez à l'IA un modèle de votre façon d'enseigner.")
    old_course_file = st.file_uploader(
        "Importer un ancien support (.pptx, .pdf, .txt)", 
        type=["pptx", "pdf", "txt"],
        help="L'IA analysera le vocabulaire, la taille des phrases et le style de cet ancien cours."
    )
    old_course_text_input = st.text_area(
        "Ou collez des notes d'un ancien cours",
        placeholder="Exemple de slide précédente : \n- L'eau bout à 100°C 🌡️\n- La vapeur monte dans l'air...",
        height=100
    )

# --- FORMULAIRE PRINCIPAL ---
col_main, col_settings = st.columns([2, 1])

with col_main:
    custom_title = st.text_input("Titre du cours (optionnel)", placeholder="Ex : La découverte des 4 saisons")
    topic = st.text_area(
        "Sujet et notions clés à aborder *",
        placeholder="Ex : Expliquer le passage de l'automne à l'hiver, la chute des feuilles, les vêtements chauds...",
        height=120
    )

with col_settings:
    nb_slides = st.slider("Nombre de diapositives", min_value=3, max_value=8, value=4)
    pedagogy_notes = st.text_area(
        "Besoins spécifiques & Profil de la classe",
        placeholder="Ex : Classe spécialisée type 4, phrases très courtes, maximum 2 puces par slide, privilégier des repères météo concrets...",
        height=100
    )

# Fonction d'extraction de texte de l'ancien fichier
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
            text_runs = []
            for slide in prs.slides:
                for shape in slide.shapes:
                    if shape.has_text_frame:
                        for paragraph in shape.text_frame.paragraphs:
                            text_runs.append(paragraph.text)
            return "\n".join(text_runs)
    except Exception as e:
        st.warning(f"Impossible de lire le fichier exemple : {e}")
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
                    "adapted_content": {
                        "type": "ARRAY",
                        "items": {"type": "STRING"},
                        "description": "Phrases très courtes, simples, repères concrets avec émojis pour besoins spécifiques."
                    },
                    "advanced_content": {
                        "type": "ARRAY",
                        "items": {"type": "STRING"},
                        "description": "Notions plus poussées, phrases complètes et vocabulaire enrichi."
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

    bg_bytes = io.BytesIO(background_file.getvalue()) if background_file else None

    # Slide 1 : Titre
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

    # Slides de contenu
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

# --- ACTION DE GÉNÉRATION ---
if st.button("✨ Générer les deux présentations", type="primary"):
    if not api_key:
        st.error("⚠️ La clé GEMINI_API_KEY est introuvable dans les Secrets de Streamlit.")
    elif not topic.strip():
        st.warning("Veuillez indiquer le sujet du cours.")
    else:
        try:
            with st.spinner("Analyse des consignes et création des slides..."):
                # Récupération de l'ancien contenu éventuel
                extracted_reference = extract_text_from_file(old_course_file)
                combined_reference = f"{extracted_reference}\n{old_course_text_input}".strip()

                reference_context = ""
                if combined_reference:
                    reference_context = f"\n\nExemple de cours précédent à utiliser comme modèle de style et structure :\n\"\"\"\n{combined_reference[:4000]}\n\"\"\""

                title_instruction = f"Titre imposé : {custom_title}" if custom_title.strip() else "Génère un titre clair et adapté."
                specific_needs = f"Besoins spécifiques de la classe : {pedagogy_notes}" if pedagogy_notes.strip() else ""

                system_prompt = (
                    "Tu es un expert en pédagogie spécialisée et en différenciation d'enseignement. "
                    "Tu dois produire deux versions distinctes pour chaque diapositive :\n"
                    "1. 'adapted_content' : syntaxe directe, phrases très courtes, repères et émojis concrets, charge cognitive minimale.\n"
                    "2. 'advanced_content' : notions complètes, phrases élaborées, vocabulaire approfondi.\n"
                    f"{reference_context}"
                )

                prompt = (
                    f"{title_instruction}\n"
                    f"Sujet : {topic}\n"
                    f"{specific_needs}\n"
                    f"Génère exactement {nb_slides} diapositives de contenu."
                )

                client = genai.Client(api_key=api_key)
                response = client.models.generate_content(
                    model="gemini-2.5-flash",
                    contents=prompt,
                    config=types.GenerateContentConfig(
                        system_instruction=system_prompt,
                        response_mime_type="application/json",
                        response_schema=response_schema,
                        temperature=0.3
                    )
                )

                st.session_state["course_data"] = json.loads(response.text)
                st.success("Présentations générées avec succès !")
        except Exception as e:
            st.error(f"Une erreur est survenue : {e}")

# --- AFFICHAGE ET TÉLÉCHARGEMENT ---
if "course_data" in st.session_state:
    data = st.session_state["course_data"]
    st.divider()
    st.subheader(f"📖 {data['title']}")
    
    tab1, tab2 = st.tabs(["🟢 Version Adaptée (Spécialisée)", "🔵 Version Approfondie"])

    with tab1:
        for idx, s in enumerate(data["slides"], 1):
            st.markdown(f"**Slide {idx} : {s['slide_title']}**")
            for pt in s["adapted_content"]:
                st.write(f"- {pt}")
        pptx_1 = create_presentation(data, "adapted_content", bg_file)
        st.download_button(
            label="📥 Télécharger la version Adaptée (.pptx)",
            data=pptx_1,
            file_name=f"cours_adapte.pptx",
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
            label="📥 Télécharger la version Approfondie (.pptx)",
            data=pptx_2,
            file_name=f"cours_avance.pptx",
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
            key="dl_2"
        )
