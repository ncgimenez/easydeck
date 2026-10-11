import io
import json
import streamlit as st
from google import genai
from google.genai import types
from pptx import Presentation
from pptx.util import Inches, Pt
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
    .slide-preview-card {
        background-color: rgba(125, 125, 125, 0.08);
        border-radius: 10px;
        padding: 1rem 1.25rem;
        margin-bottom: 1rem;
    }
    .slide-title-tag {
        font-weight: 700;
        font-size: 1.05rem;
        margin-bottom: 0.4rem;
    }
    div.stButton > button {
        border-radius: 10px;
        font-weight: bold;
    }
</style>
""", unsafe_allow_html=True)

st.markdown("""
<div class="main-header">
    <div class="main-title">🪄 DuoDeck</div>
    <div class="main-subtitle">Générez vos diapositives adaptées ou approfondies de façon ciblée.</div>
</div>
""", unsafe_allow_html=True)

api_key = st.secrets.get("GEMINI_API_KEY", None)

# --- 1. MODÈLE & AMBIANCE ---
st.subheader("🎨 Modèle & Ambiance visuelle")
template_file = st.file_uploader(
    "Déposez votre présentation modèle (.pptx, optionnel)",
    type=["pptx"],
    help="DuoDeck conservera vos polices, formes et arrières-plans."
)

col_opt1, col_opt2 = st.columns(2)
with col_opt1:
    use_style_instruction = st.checkbox("Imiter le style d'écriture du modèle", value=True)
with col_opt2:
    color_scheme = st.selectbox(
        "Nuance des couleurs",
        [
            "Harmonie créative adaptée au sujet (recommandé)",
            "Conserver les couleurs d'origine du modèle",
            "Bleu nuit & Ardoise",
            "Tons chauds (Terracotta)",
            "Vert sauge (Apaisant)"
        ],
        index=0
    )

st.divider()

# --- 2. PARAMÈTRES DU COURS ---
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
        "Consignes spécifiques",
        placeholder="Ex : 2 phrases max, émojis concrets, pas de vocabulaire abstrait..."
    )

# --- OUTILS & SCHÉMAS ---
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
                    "bullets": {
                        "type": "ARRAY",
                        "items": {"type": "STRING"}
                    }
                },
                "required": ["slide_title", "bullets"]
            }
        }
    },
    "required": ["title", "slides"]
}

def get_color_rgb(choice, is_title=True):
    if choice == "Bleu nuit & Ardoise":
        return RGBColor(30, 41, 59) if is_title else RGBColor(71, 85, 105)
    elif choice == "Tons chauds (Terracotta)":
        return RGBColor(154, 52, 18) if is_title else RGBColor(120, 53, 15)
    elif choice == "Vert sauge (Apaisant)":
        return RGBColor(47, 79, 79) if is_title else RGBColor(64, 93, 84)
    # Mode créatif automatique par défaut : palette moderne bleu-indigo contrastée
    return RGBColor(37, 99, 235) if is_title else RGBColor(51, 65, 85)

def build_presentation(template_bytes, data, level_name, color_choice):
    if template_bytes is None:
        prs = Presentation()
        prs.slide_width, prs.slide_height = Inches(13.333), Inches(7.5)
        blank = prs.slide_layouts[6]
        
        # Titre
        s_title = prs.slides.add_slide(blank)
        tb = s_title.shapes.add_textbox(Inches(1.5), Inches(2.2), Inches(10.33), Inches(3))
        p = tb.text_frame.paragraphs[0]
        p.text = data["title"]
        p.font.size = Pt(44)
        p.font.bold = True
        p.font.color.rgb = get_color_rgb(color_choice, is_title=True)

        p2 = tb.text_frame.add_paragraph()
        p2.text = f"— {level_name} —"
        p2.font.size = Pt(22)
        p2.font.color.rgb = RGBColor(100, 116, 139)

        # Diapositives
        for item in data["slides"]:
            s = prs.slides.add_slide(blank)
            box = s.shapes.add_textbox(Inches(1.2), Inches(1.2), Inches(11), Inches(5))
            tf = box.text_frame
            tf.word_wrap = True
            
            pt = tf.paragraphs[0]
            pt.text = item["slide_title"]
            pt.font.size = Pt(32)
            pt.font.bold = True
            pt.font.color.rgb = get_color_rgb(color_choice, is_title=True)
            pt.space_after = Pt(16)

            for b in item["bullets"]:
                pb = tf.add_paragraph()
                pb.text = b
                pb.font.size = Pt(22)
                pb.font.color.rgb = get_color_rgb(color_choice, is_title=False)
                pb.space_after = Pt(12)
    else:
        prs = Presentation(io.BytesIO(template_bytes))
        content_layout = prs.slides[1].slide_layout if len(prs.slides) > 1 else (prs.slides[0].slide_layout if len(prs.slides) > 0 else prs.slide_layouts[0])

        if len(prs.slides) > 0:
            for s in prs.slides[0].shapes:
                if s.has_text_frame and s.text_frame.text.strip():
                    s.text_frame.text = data["title"]
                    break

        for item in data["slides"]:
            ns = prs.slides.add_slide(content_layout)
            placed_title, placed_content = False, False

            for shape in ns.shapes:
                if shape.has_text_frame:
                    tf = shape.text_frame
                    if not placed_title:
                        tf.text = item["slide_title"]
                        placed_title = True
                        if color_choice != "Conserver les couleurs d'origine du modèle":
                            for p in tf.paragraphs:
                                p.font.color.rgb = get_color_rgb(color_choice, is_title=True)
                    elif not placed_content:
                        tf.text = ""
                        for idx, b in enumerate(item["bullets"]):
                            p = tf.add_paragraph() if idx > 0 else tf.paragraphs[0]
                            p.text = b
                            p.font.size = Pt(22)
                            p.space_after = Pt(10)
                            if color_choice != "Conserver les couleurs d'origine du modèle":
                                p.font.color.rgb = get_color_rgb(color_choice, is_title=False)
                        placed_content = True
                        break

            if not placed_title or not placed_content:
                box = ns.shapes.add_textbox(Inches(1.2), Inches(1.2), Inches(11), Inches(5))
                tf = box.text_frame
                tf.word_wrap = True
                pt = tf.paragraphs[0]
                pt.text = item["slide_title"]
                pt.font.size = Pt(32)
                pt.font.bold = True
                pt.space_after = Pt(14)
                if color_choice != "Conserver les couleurs d'origine du modèle":
                    pt.font.color.rgb = get_color_rgb(color_choice, is_title=True)

                for b in item["bullets"]:
                    pb = tf.add_paragraph()
                    pb.text = b
                    pb.font.size = Pt(22)
                    pb.space_after = Pt(10)
                    if color_choice != "Conserver les couleurs d'origine du modèle":
                        pb.font.color.rgb = get_color_rgb(color_choice, is_title=False)

    out = io.BytesIO()
    prs.save(out)
    out.seek(0)
    return out

def run_generation(level_type):
    if not api_key:
        st.error("⚠️ Clé GEMINI_API_KEY manquante dans les Secrets Streamlit.")
        return
    if not topic.strip():
        st.warning("Veuillez renseigner un sujet de cours.")
        return

    progress_bar = st.progress(0)
    status_text = st.empty()

    try:
        status_text.markdown("*Étape 1/3 : Analyse du contexte et des consignes... (environ 10 secondes)*")
        progress_bar.progress(20)

        t_bytes = template_file.getvalue() if template_file else None
        style_guide = ""
        if t_bytes and use_style_instruction:
            samples = extract_writing_style(t_bytes)
            if samples:
                style_guide = f"\n\nModèle d'écriture à imiter strictement en termes de structure et de ton (sans réutiliser les sujets) :\n\"\"\"\n{samples}\n\"\"\""

        is_adapted = (level_type == "adapted")
        pedagogy_target = (
            "Syntaxe directe et épurée, repères visuels clairs avec émojis concrets, charge cognitive minimale."
            if is_adapted else
            "Notions approfondies, phrases rédigées et vocabulaire technique enrichi."
        )

        title_rule = f"Titre imposé : {custom_title}" if custom_title.strip() else "Trouve un titre pédagogique accrocheur."
        needs_rule = f"Consignes pédagogiques spécifiques : {pedagogy_notes}" if pedagogy_notes.strip() else ""

        sys_prompt = (
            f"Tu es un expert en pédagogie spécialisée. Rédige les diapositives pour la version : '{level_type}'.\n"
            f"Objectif : {pedagogy_target}\n"
            f"{style_guide}"
        )
        prompt = f"{title_rule}\nSujet : {topic}\n{needs_rule}\nGénère exactement {nb_slides} diapositives avec le champ 'bullets' pour chaque diapositive."

        status_text.markdown("*Étape 2/3 : Rédaction de la présentation par l'IA...*")
        progress_bar.progress(55)

        client = genai.Client(api_key=api_key)
        response = None
        last_err = None

        
        for mdl in ["gemini-2.0-flash", "gemini-2.0-flash-lite"]:
        
                    
            
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
            except Exception as e:
                last_err = e
                continue

        if not response or not response.text:
            raise last_err

        status_text.markdown("*Étape 3/3 : Finalisation et assemblage du support PowerPoint...*")
        progress_bar.progress(90)

        data = json.loads(response.text)
        session_key = f"deck_{level_type}"
        st.session_state[session_key] = {
            "data": data,
            "pptx": build_presentation(t_bytes, data, "Niveau Adapté" if is_adapted else "Niveau Approfondi", color_scheme)
        }

        progress_bar.progress(100)
        status_text.markdown("*Terminé !*")
        st.success(f"🎉 Présentation {'Adaptée' if is_adapted else 'Avancée'} prête à être téléchargée !")

    except Exception as e:
        progress_bar.empty()
        status_text.empty()
        st.error(f"Une erreur est survenue : {e}")

# --- 3. LES DEUX BOUTONS DE CRÉATION SÉPARÉS ---
col_btn1, col_btn2 = st.columns(2)
with col_btn1:
    if st.button("🟢 Générer la version Adaptée", type="primary", use_container_width=True):
        run_generation("adapted")

with col_btn2:
    if st.button("🔵 Générer la version Avancée", type="secondary", use_container_width=True):
        run_generation("advanced")

# --- 4. AFFICHAGE ET TÉLÉCHARGEMENT ---
has_adapted = "deck_adapted" in st.session_state
has_advanced = "deck_advanced" in st.session_state

if has_adapted or has_advanced:
    st.divider()
    tabs_to_show = []
    if has_adapted:
        tabs_to_show.append("🟢 Version Adaptée")
    if has_advanced:
        tabs_to_show.append("🔵 Version Approfondie")

    tab_objs = st.tabs(tabs_to_show)

    current_idx = 0
    if has_adapted:
        with tab_objs[current_idx]:
            deck = st.session_state["deck_adapted"]
            st.markdown(f"### 📖 {deck['data']['title']}")
            for idx, s in enumerate(deck["data"]["slides"], 1):
                st.markdown(f"""
                <div class="slide-preview-card" style="border-left: 4px solid #10b981;">
                    <div class="slide-title-tag">Diapositive {idx} : {s['slide_title']}</div>
                    {''.join([f'<div style=\"margin-bottom: 3px;\">• {b}</div>' for b in s['bullets']])}
                </div>
                """, unsafe_allow_html=True)
            st.download_button(
                "📥 Télécharger la version Adaptée (.pptx)",
                deck["pptx"],
                file_name="cours_adapte.pptx",
                key="dl_adapted"
            )
        current_idx += 1

    if has_advanced:
        with tab_objs[current_idx]:
            deck = st.session_state["deck_advanced"]
            st.markdown(f"### 📖 {deck['data']['title']}")
            for idx, s in enumerate(deck["data"]["slides"], 1):
                st.markdown(f"""
                <div class="slide-preview-card" style="border-left: 4px solid #3b82f6;">
                    <div class="slide-title-tag">Diapositive {idx} : {s['slide_title']}</div>
                    {''.join([f'<div style=\"margin-bottom: 3px;\">• {b}</div>' for b in s['bullets']])}
                </div>
                """, unsafe_allow_html=True)
            st.download_button(
                "📥 Télécharger la version Approfondie (.pptx)",
                deck["pptx"],
                file_name="cours_avance.pptx",
                key="dl_advanced"
            )
