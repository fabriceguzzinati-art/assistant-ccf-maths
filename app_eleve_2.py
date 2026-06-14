# ============================================================
# app_eleve.py — Interface élève
# Toute la logique métier est dans utils.py
# ============================================================

import random
import streamlit as st
from utils import (
    # Constantes
    THEMES, NIVEAUX_CATEGORIES, MATIERES, LISTE_FILIERES, CONTEXTES_FILIERES,
    NIVEAUX_DIFFICULTE, ORDRE_NIVEAUX_DIFF, MASCOTTES, MESSAGES_VICTOIRE,
    BADGES, XP_PAR_ACTION, RANGS, SEUIL_VALIDATION,
    # Helpers chapitres
    get_chapitres, build_contexte_filiere,
    # Prompts
    build_prompt_exercices, build_prompt_ccf_entrainement,
    build_prompt_ccf_officiel, build_prompt_boss,
    # Gemini
    call_gemini,
    # Grist
    envoyer_grist, envoyer_proposition_grist,
    lire_progression_grist, lire_classement_grist,
    # Fonctions métier
    calculer_xp, calculer_streak, calculer_progression,
    calculer_objectif, calculer_badges, niveau_suivant,
    # Banque
    charger_banque, charger_banque_interactif,
    # Export Word
    markdown_to_docx, generate_ccf_officiel_docx,
)

# ============================================================
# 1. CONFIG — doit être le PREMIER appel Streamlit
# ============================================================
st.set_page_config(
    page_title="Entraînement Bac Pro — Maths",
    page_icon="📚",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ============================================================
# 2. THÈME — sélecteur + CSS
# ============================================================
theme_keys = list(THEMES.keys())
if "theme_pref" not in st.session_state or st.session_state.theme_pref not in theme_keys:
    st.session_state.theme_pref = theme_keys[0]

theme_nom = st.sidebar.selectbox(
    "🎨 Style de l'interface",
    theme_keys,
    index=theme_keys.index(st.session_state.theme_pref),
)
st.session_state.theme_pref = theme_nom
t = THEMES[theme_nom]

st.markdown(f"""<style>
@import url('https://fonts.googleapis.com/css2?family=Outfit:wght@400;600;700;800&display=swap');
.stApp {{ background:{t['bg_app']}; color:{t['text']}; font-family:'Outfit',sans-serif; }}
.stApp > header {{ background:transparent !important; }}
[data-testid="stSidebar"] {{ background:{t['bg_side']} !important; border-right:1px solid {t['border']}; }}
[data-testid="stSidebar"] * {{ color:{t['text']} !important; }}
.stWidgetLabel p, div[data-testid="stMarkdownContainer"] p {{ color:{t['text']} !important; font-weight:600; }}
div[data-testid="stRadio"] div[role="radiogroup"] div[aria-checked="true"] > div {{ background-color:{t['primary']} !important; border-color:{t['primary']} !important; }}
.stTabs [data-baseweb="tab-list"] {{ background:{t['bg_side']}; border-radius:12px; padding:4px; gap:4px; border:1px solid {t['border']}; }}
.stTabs [data-baseweb="tab"] {{ background:transparent; color:{t['text']}; opacity:.7; border-radius:8px; font-weight:600; }}
.stTabs [aria-selected="true"] {{ background:linear-gradient(135deg,{t['primary']},{t['secondary']}) !important; color:white !important; box-shadow:0 0 16px {t['primary']}80; }}
.stButton > button[kind="primary"] {{ background:linear-gradient(135deg,{t['primary']},{t['secondary']}); border:none; border-radius:12px; font-weight:700; color:white; box-shadow:0 0 20px {t['primary']}66; transition:all .25s; }}
.stButton > button[kind="primary"]:hover {{ transform:translateY(-2px); box-shadow:0 0 30px {t['primary']}B3; }}
.stSelectbox > div > div, .stTextInput > div > div > input {{ background:{t['bg_side']} !important; border:1px solid {t['border']} !important; color:{t['text']} !important; border-radius:10px !important; }}
[data-testid="stMetric"] {{ background:{t['bg_side']}; border:1px solid {t['border']}; border-radius:12px; padding:16px; }}
[data-testid="stMetricValue"] {{ color:{t['accent']} !important; font-weight:800; }}
.xp-banner {{ background:linear-gradient(135deg,{t['bg_side']},{t['bg_app']}); border:1px solid {t['border']}; border-radius:14px; padding:14px 20px; margin-bottom:16px; display:flex; align-items:center; gap:16px; }}
.xp-value {{ font-family:'Outfit',sans-serif; font-weight:800; font-size:1.4rem; color:{t['primary']}; }}
.boss-banner {{ background:linear-gradient(135deg,{t['bg_side']},{t['secondary']}44); border:2px solid {t['primary']}; border-radius:16px; padding:24px; text-align:center; margin:16px 0; box-shadow:0 0 40px {t['primary']}33; }}
::-webkit-scrollbar {{ width:6px; }}
::-webkit-scrollbar-track {{ background:{t['bg_app']}; }}
::-webkit-scrollbar-thumb {{ background:{t['border']}; border-radius:3px; }}
::-webkit-scrollbar-thumb:hover {{ background:{t['primary']}; }}
.stMarkdown h1,.stMarkdown h2,.stMarkdown h3 {{ color:{t['text']}; }}
.stMarkdown p,.stMarkdown li {{ color:{t['text']}; opacity:.9; }}
.stCheckbox > label {{ color:#94a3b8 !important; }}
.stRadio > label {{ color:#94a3b8 !important; }}
.stExpander {{ background:#13162a !important; border:1px solid #2d3561 !important; border-radius:10px !important; }}
.info-box {{ background:#f0f4ff; border-left:4px solid #4a6cf7; padding:12px 16px; border-radius:4px; margin:8px 0; font-size:.9rem; }}
.warn-box {{ background:#fff8e1; border-left:4px solid #f59e0b; padding:12px 16px; border-radius:4px; margin:8px 0; font-size:.9rem; }}
.ok-box   {{ background:#f0fdf4; border-left:4px solid #22c55e; padding:12px 16px; border-radius:4px; margin:8px 0; font-size:.9rem; }}
.badge    {{ display:inline-block; background:#e0e7ff; color:#3730a3; border-radius:6px; padding:3px 10px; font-size:.8rem; margin-right:6px; margin-bottom:4px; }}
</style>""", unsafe_allow_html=True)

# ============================================================
# 3. PRÉFÉRENCES IDENTITÉ
# ============================================================
if "genre_pref"  not in st.session_state: st.session_state.genre_pref  = "Neutre (Aventurier)"
if "avatar_pref" not in st.session_state: st.session_state.avatar_pref = "Robotique 🤖"

col1, col2 = st.columns(2)
with col1:
    genre = st.radio("Comment souhaites-tu que l'on s'adresse à toi ?",
                     ["Neutre (Aventurier)", "Féminin (Aventurière)", "Masculin (Aventurier)"],
                     key="radio_genre")
with col2:
    avatar_style = st.selectbox("Choisis ton style d'avatar",
                                ["Robotique 🤖", "Mage 🧙‍♂️", "Guerrier/ère 🛡️", "Animalier 🐾"],
                                key="select_avatar")
st.session_state.genre_pref  = genre
st.session_state.avatar_pref = avatar_style

# ============================================================
# 4. SESSION STATE
# ============================================================
for key in ["generated_md", "generated_ccf_md", "meta_gen", "meta_ccf",
            "eval_gen_done", "eval_ccf_done", "progression_cache",
            "boss_actif", "boss_niveau", "boss_chapitre", "boss_md", "eval_boss_done",
            "streak_cache", "interactif_sujet", "interactif_idx",
            "interactif_reponses", "interactif_termine", "interactif_eval_done",
            "interactif_done", "classement_cache"]:
    if key not in st.session_state:
        st.session_state[key] = None

# ============================================================
# 5. SIDEBAR
# ============================================================
with st.sidebar:
    st.markdown("""
    <div style="text-align:center;padding:16px 0 8px">
        <div style="font-size:2.5rem">🎓</div>
        <div style="font-family:'Outfit',sans-serif;font-weight:800;font-size:1.1rem;
                    color:#a5b4fc;letter-spacing:.5px">MATHS BAC PRO</div>
        <div style="font-size:.75rem;color:#475569;margin-top:2px">Entraînement IA</div>
    </div>""", unsafe_allow_html=True)
    st.divider()

    st.markdown('<div style="font-family:Outfit,sans-serif;font-weight:700;color:#94a3b8;font-size:.8rem;letter-spacing:1px;margin-bottom:6px">🎮 TON IDENTIFIANT</div>', unsafe_allow_html=True)
    code_eleve = st.text_input("Code élève", placeholder="Ex : ASSP-03", key="code_eleve",
                               label_visibility="collapsed",
                               help="Code distribué par ton professeur.").strip().upper()

    if code_eleve:
        if not st.session_state.streak_cache:
            st.session_state.streak_cache = calculer_streak(lire_progression_grist(code_eleve))
        s      = st.session_state.streak_cache or {}
        streak = s.get("streak", 0)
        record = s.get("record", 0)

        from datetime import date
        derniere = s.get("derniere_date")
        if derniere == date.today(): salut = "Bonne continuation"
        elif streak > 0:             salut = "Bon retour"
        else:                        salut = "Bienvenue"

        if streak >= 7:   feu, cb = "🔥🔥🔥", "#ef4444"
        elif streak >= 3: feu, cb = "🔥🔥",   "#f97316"
        elif streak >= 1: feu, cb = "🔥",     "#f59e0b"
        else:             feu, cb = "👋",     "#6366f1"

        streak_txt = (f"{feu} {streak} jour{'s' if streak > 1 else ''} d'affilée !"
                      if streak > 0 else "Lance ton premier streak aujourd'hui !")
        st.markdown(
            f'<div style="background:linear-gradient(135deg,#13162a,#1e2235);'
            f'border:1px solid {cb}44;border-radius:10px;padding:12px 14px;margin:6px 0">'
            f'<div style="font-size:.82rem;color:#94a3b8">{salut} 👋</div>'
            f'<div style="font-family:Outfit,sans-serif;font-weight:800;color:#e2e8f0;font-size:1rem;margin:2px 0">{code_eleve}</div>'
            f'<div style="font-size:.82rem;color:{cb};margin-top:4px">{streak_txt}</div>'
            + (f'<div style="font-size:.72rem;color:#475569;margin-top:2px">Record : {record} jour{"s" if record > 1 else ""}</div>' if record > 1 else "")
            + '</div>', unsafe_allow_html=True)
    else:
        st.markdown('<div class="warn-box">⚠️ Entre ton code pour sauvegarder ta progression.</div>', unsafe_allow_html=True)

    st.divider()
    st.markdown('<div style="font-family:Outfit,sans-serif;font-weight:700;color:#94a3b8;font-size:.8rem;letter-spacing:1px;margin-bottom:6px">🔑 CLÉ GEMINI</div>', unsafe_allow_html=True)
    cle_api = st.text_input("Clé API", type="password", key="gemini_key", label_visibility="collapsed")
    if cle_api:
        st.markdown('<div class="ok-box">✅ Prêt à générer !</div>', unsafe_allow_html=True)
    else:
        st.markdown('<div class="warn-box">⚠️ Clé API manquante.</div>', unsafe_allow_html=True)
    with st.expander("📋 Obtenir une clé gratuite (2 min)"):
        st.markdown("""**1.** → [aistudio.google.com](https://aistudio.google.com/app/apikey)
**2.** Connecte-toi avec Google  
**3.** Clique **"Create API Key"**  
**4.** Copie et colle la clé ici  
💡 **C'est gratuit** — pas de carte bancaire requise.""")
    st.divider()
    st.caption("Entraînement Bac Pro · Gemini 2.5 Flash")

# ============================================================
# 6. TITRE
# ============================================================
st.markdown("""
<div style="padding:24px 0 8px">
    <div style="font-family:'Outfit',sans-serif;font-weight:800;font-size:2rem;
                background:linear-gradient(90deg,#a5b4fc,#c084fc,#67e8f9);
                -webkit-background-clip:text;-webkit-text-fill-color:transparent;
                background-clip:text;line-height:1.2">📚 Entraînement Bac Pro</div>
    <div style="color:#475569;font-size:.9rem;margin-top:4px">Exercices · CCF · Progression — propulsé par l'IA</div>
</div>""", unsafe_allow_html=True)

# ============================================================
# 7. ONGLETS
# ============================================================
tab_gen, tab_ccf, tab_graphique, tab_progression = st.tabs([
    "📝 Exercices d'entraînement",
    "🎯 Sujets CCF",
    "📈 Graphique GeoGebra",
    "📊 Ma progression",
])

# ─────────────────────────────────────────────────────────────
# ONGLET 1 — GÉNÉRATEUR
# ─────────────────────────────────────────────────────────────
with tab_gen:
    # Objectif du jour
    if code_eleve and st.session_state.progression_cache:
        chap_actif  = (st.session_state.meta_gen or {}).get("chapitre", "")
        diff_active = (st.session_state.meta_gen or {}).get("difficulte", "")
        obj = calculer_objectif(st.session_state.progression_cache, chap_actif, diff_active)
        st.markdown(
            f'<div style="background:#0d0f1a;border:1px solid {obj["couleur"]}55;'
            f'border-left:4px solid {obj["couleur"]};border-radius:0 10px 10px 0;'
            f'padding:10px 16px;margin-bottom:12px">'
            f'<span style="font-size:1.1rem">{obj["emoji"]}</span> '
            f'<span style="color:#e2e8f0;font-size:.88rem">{obj["message"]}</span>'
            f'</div>', unsafe_allow_html=True)

    st.subheader("📝 Générateur de sujets et exercices")
    col1, col2 = st.columns(2)
    with col1:
        cat     = st.selectbox("Type d'établissement", list(NIVEAUX_CATEGORIES.keys()), key="gen_cat")
        niv     = st.selectbox("Classe", NIVEAUX_CATEGORIES[cat], key="gen_niv")
        filiere = ""
        if cat in ["Bac Pro", "CAP"]:
            fil = st.selectbox("Filière", LISTE_FILIERES, key="gen_fil")
            if fil == "Autre (Préciser ci-dessous)":
                filiere = st.text_input("Filière libre", key="gen_spec")
            else:
                filiere = fil
                if fil in CONTEXTES_FILIERES:
                    st.info(f"📌 {CONTEXTES_FILIERES[fil]['nom_complet']}")
    with col2:
        mat  = st.selectbox("Matière", MATIERES, key="gen_mat")
        chap = st.selectbox("Chapitre (BO)", get_chapitres(mat, niv, cat), key="gen_chap")

    consignes  = st.text_area("Consignes particulières (optionnel)", height=80, key="gen_consignes",
                               placeholder="Ex : 3 exercices, niveau accessible…")
    st.markdown("**🎯 Niveau de difficulté**")
    difficulte = st.select_slider("Niveau", options=list(NIVEAUX_DIFFICULTE.keys()),
                                  value="🟡 Moyen", key="gen_diff", label_visibility="collapsed")
    st.markdown(f'<div class="info-box">{difficulte} — {NIVEAUX_DIFFICULTE[difficulte]}</div>',
                unsafe_allow_html=True)

    sujets_banque      = charger_banque(niv, filiere, chap, difficulte)
    sujets_interactifs = charger_banque_interactif(niv, filiere, chap, difficulte)
    nb_banque          = len(sujets_banque)
    nb_interactifs     = len(sujets_interactifs)

    col_btn1, col_btn2, col_btn3 = st.columns(3)
    with col_btn1:
        btn_interactif = st.button(f"⚡ Mode interactif ({nb_interactifs} dispo)",
                                   type="primary", use_container_width=True,
                                   disabled=(nb_interactifs == 0), key="btn_interactif")
        if nb_interactifs == 0: st.caption("Aucun sujet interactif en banque.")
    with col_btn2:
        btn_banque = st.button(f"📚 Sujet classique ({nb_banque} dispo)",
                               type="primary", use_container_width=True,
                               disabled=(nb_banque == 0), key="btn_banque")
        if nb_banque == 0: st.caption("Aucun sujet classique en banque.")
    with col_btn3:
        btn_gemini = st.button("✨ Générer (Gemini)", type="primary",
                               use_container_width=True, key="btn_gemini")

    # ── Bouton interactif ─────────────────────────────────────
    if btn_interactif and nb_interactifs > 0:
        combo_key = f"{niv}|{filiere}|{chap}|{difficulte}"
        if not st.session_state.interactif_done:
            st.session_state.interactif_done = {}
        done_set        = set(st.session_state.interactif_done.get(combo_key, []))
        indices_dispo   = [i for i in range(nb_interactifs) if i not in done_set]
        if indices_dispo:
            idx = random.choice(indices_dispo)
            done_set.add(idx)
            st.session_state.interactif_done[combo_key] = list(done_set)
            st.session_state.interactif_sujet    = sujets_interactifs[idx]
            st.session_state.interactif_idx      = 0
            st.session_state.interactif_reponses = {}
            st.session_state.interactif_termine  = False
            st.session_state.interactif_eval_done = None
            st.session_state.generated_md         = None
            st.session_state.eval_gen_done        = None
            st.session_state.meta_gen = {"niveau": niv, "matiere": mat, "chapitre": chap,
                                         "filiere": filiere, "difficulte": difficulte,
                                         "source": "Banque interactif"}
            st.rerun()
        else:
            st.warning(f"🎉 Tu as déjà fait les {nb_interactifs} exercice(s) disponibles !")
            c1, c2 = st.columns(2)
            with c1:
                if st.button("🔄 Recommencer depuis le début", use_container_width=True,
                             key="btn_reset_done"):
                    st.session_state.interactif_done[combo_key] = []
                    st.rerun()
            with c2:
                if st.button("✨ Générer via Gemini", use_container_width=True,
                             key="btn_gen_from_done"):
                    if not cle_api:
                        st.error("🔑 Clé API manquante !")
                    else:
                        with st.spinner("⏳ Génération…"):
                            try:
                                res = call_gemini(cle_api, build_prompt_exercices(
                                    niv, cat, mat, chap, "", filiere, difficulte))
                                st.session_state.generated_md     = res
                                st.session_state.interactif_sujet = None
                                st.session_state.eval_gen_done    = None
                                st.session_state.meta_gen = {"niveau": niv, "matiere": mat,
                                    "chapitre": chap, "filiere": filiere,
                                    "difficulte": difficulte, "source": "Gemini"}
                                st.rerun()
                            except Exception as e:
                                st.error(f"Erreur : {e}")

    # ── Bouton banque classique ────────────────────────────────
    if btn_banque and nb_banque > 0:
        sujet = random.choice(sujets_banque)
        st.session_state.generated_md         = sujet["contenu"]
        st.session_state.interactif_sujet     = None
        st.session_state.eval_gen_done        = None
        st.session_state.meta_gen = {"niveau": niv, "matiere": mat, "chapitre": chap,
                                     "filiere": filiere, "difficulte": difficulte,
                                     "source": "Banque"}
        st.rerun()

    # ── Bouton Gemini ─────────────────────────────────────────
    if btn_gemini:
        if not cle_api:
            st.error("🔑 Renseigne ta clé API dans le panneau gauche !")
        else:
            with st.spinner("⏳ Génération en cours…"):
                try:
                    res = call_gemini(cle_api, build_prompt_exercices(
                        niv, cat, mat, chap, consignes, filiere, difficulte))
                    st.session_state.generated_md     = res
                    st.session_state.interactif_sujet = None
                    st.session_state.eval_gen_done    = None
                    st.session_state.meta_gen = {"niveau": niv, "matiere": mat,
                        "chapitre": chap, "filiere": filiere,
                        "difficulte": difficulte, "source": "Gemini"}
                    envoyer_proposition_grist(code_eleve or "anonyme",
                                              st.session_state.meta_gen, res)
                    st.success("✅ Sujet généré !")
                except Exception as e:
                    if "429" in str(e):
                        st.error("⏱️ Quota dépassé. Attends 1 minute ou utilise la banque.")
                    else:
                        st.error(f"Erreur API : {e}")

    # ── MODE INTERACTIF ───────────────────────────────────────
    if st.session_state.get("interactif_sujet"):
        sujet_i   = st.session_state.interactif_sujet
        questions = sujet_i.get("questions", [])
        idx       = st.session_state.get("interactif_idx", 0)
        reponses  = st.session_state.get("interactif_reponses", {})
        termine   = st.session_state.get("interactif_termine", False)

        if sujet_i.get("contexte"):
            st.markdown(f'<div class="info-box">🏢 <strong>Mise en situation</strong><br>{sujet_i["contexte"]}</div>',
                        unsafe_allow_html=True)
        if sujet_i.get("rappel_cours"):
            st.markdown(f'<div style="background:#1a2040;border:1px solid #3d4480;border-radius:10px;'
                        f'padding:12px 16px;margin:8px 0;font-size:.88rem;color:#a5b4fc">'
                        f'📐 <strong>Rappel</strong> : {sujet_i["rappel_cours"]}</div>',
                        unsafe_allow_html=True)

        if not termine:
            st.markdown(f'<div style="margin:12px 0 4px;font-size:.82rem;color:#64748b">'
                        f'Question {idx + 1} / {len(questions)}</div>', unsafe_allow_html=True)
            st.progress(idx / len(questions))
            q = questions[idx]
            st.markdown(f'<div style="background:#13162a;border:1px solid #3d4480;border-radius:12px;'
                        f'padding:18px 20px;margin:12px 0">'
                        f'<div style="font-size:.8rem;color:#6366f1;font-weight:700;margin-bottom:6px">'
                        f'QUESTION {q["numero"]}</div>'
                        f'<div style="color:#e2e8f0;font-size:.95rem">{q["enonce"]}</div>'
                        f'</div>', unsafe_allow_html=True)

            col_rep, col_ind = st.columns([3, 1])
            with col_rep:
                unite      = q.get("unite", "")
                rep_saisie = st.number_input(f"Ta réponse {f'({unite})' if unite else ''}",
                                             value=0.0, step=0.01, format="%.2f",
                                             key=f"rep_q_{idx}")
            with col_ind:
                if q.get("indice") and st.button("💡 Indice", key=f"ind_{idx}",
                                                  use_container_width=True):
                    st.info(q["indice"])

            col_v, col_p = st.columns(2)
            with col_v:
                if st.button("✅ Valider", type="primary", use_container_width=True,
                             key=f"val_{idx}"):
                    tolerance = float(q.get("tolerance", 0.5))
                    bonne     = abs(rep_saisie - float(q["reponse"])) <= tolerance
                    reponses[idx] = {"saisie": rep_saisie, "bonne": bonne,
                                     "reponse": q["reponse"], "unite": unite,
                                     "corrige": q.get("corrige", "")}
                    st.session_state.interactif_reponses = reponses
                    if idx + 1 < len(questions):
                        st.session_state.interactif_idx = idx + 1
                    else:
                        st.session_state.interactif_termine = True
                    st.rerun()
            with col_p:
                if idx > 0 and st.button("⬅️ Précédente", use_container_width=True,
                                          key=f"prev_{idx}"):
                    st.session_state.interactif_idx = idx - 1
                    st.rerun()
        else:
            # ── RÉSULTATS FINAUX ──────────────────────────────
            nb_bonnes_i = sum(1 for r in reponses.values() if r["bonne"])
            nb_total_i  = len(questions)
            score_pct   = int(nb_bonnes_i / nb_total_i * 100) if nb_total_i else 0

            if score_pct == 100: st.balloons(); verdict, couleur = "🏆 Parfait !",               "#065f46"
            elif score_pct >= 70:               verdict, couleur = "👍 Bien joué !",              "#1e3a5f"
            elif score_pct >= 50:               verdict, couleur = "😐 Peut mieux faire",         "#3f2a00"
            else:                               verdict, couleur = "💪 Continue à t'entraîner !", "#3b0000"

            st.markdown(f'<div style="background:{couleur};border-radius:14px;padding:20px;'
                        f'text-align:center;margin:12px 0">'
                        f'<div style="font-size:2rem">{verdict}</div>'
                        f'<div style="font-size:1.5rem;font-weight:800;color:white;margin:8px 0">'
                        f'{nb_bonnes_i} / {nb_total_i} — {score_pct}%</div>'
                        f'</div>', unsafe_allow_html=True)

            # Étape 1 : Ressenti AVANT la correction
            if not st.session_state.interactif_eval_done:
                st.divider()
                st.markdown("**🎯 Et ton ressenti avant de voir la correction ?**")
                col_e1, col_e2, col_e3, col_e4 = st.columns(4)
                eval_choix_i = None
                with col_e1:
                    if st.button("😕 Difficile", use_container_width=True, key="eval_i_1"): eval_choix_i = "😕 Difficile"
                with col_e2:
                    if st.button("😐 Moyen",     use_container_width=True, key="eval_i_2"): eval_choix_i = "😐 Moyen"
                with col_e3:
                    if st.button("😊 Bien",      use_container_width=True, key="eval_i_3"): eval_choix_i = "😊 Bien"
                with col_e4:
                    if st.button("🌟 Très bien", use_container_width=True, key="eval_i_4"): eval_choix_i = "🌟 Très bien"
                if eval_choix_i:
                    m_i = st.session_state.meta_gen or {}
                    envoyer_grist(code_eleve or "anonyme",
                                  f"Exercice-interactif ({score_pct}%)",
                                  {**m_i, "score_auto": f"{nb_bonnes_i}/{nb_total_i}"},
                                  eval_choix_i,
                                  genre=st.session_state.genre_pref,
                                  avatar=st.session_state.avatar_pref)
                    st.session_state.interactif_eval_done = eval_choix_i
                    st.session_state.streak_cache         = None
                    st.rerun()
            else:
                # Étape 2 : Correction + bouton suivant
                st.markdown(f'<div class="ok-box">✅ Ressenti enregistré : '
                            f'<strong>{st.session_state.interactif_eval_done}</strong></div>',
                            unsafe_allow_html=True)
                st.markdown("### 📋 Correction détaillée")
                for i, q in enumerate(questions):
                    r     = reponses.get(i, {})
                    icone = "✅" if r.get("bonne") else "❌"
                    unite = r.get("unite", "")
                    st.markdown(
                        f'<div style="background:#13162a;border-left:3px solid '
                        f'{"#22c55e" if r.get("bonne") else "#ef4444"};'
                        f'border-radius:0 10px 10px 0;padding:12px 16px;margin:6px 0">'
                        f'<div style="font-size:.8rem;color:#64748b">Q{q["numero"]}</div>'
                        f'<div style="color:#e2e8f0;margin:4px 0">{q["enonce"]}</div>'
                        f'<div style="font-size:.88rem;margin-top:6px">'
                        f'{icone} Ta réponse : <strong>{r.get("saisie","?"):.2f} {unite}</strong> — '
                        f'Bonne réponse : <strong>{float(q["reponse"]):.2f} {unite}</strong></div>'
                        f'<div style="font-size:.82rem;color:#94a3b8;margin-top:4px">'
                        f'💡 {r.get("corrige","")}</div></div>', unsafe_allow_html=True)

                st.divider()
                m_i       = st.session_state.meta_gen or {}
                combo_key = f"{m_i.get('niveau')}|{m_i.get('filiere')}|{m_i.get('chapitre')}|{m_i.get('difficulte')}"
                if not st.session_state.interactif_done:
                    st.session_state.interactif_done = {}
                done_set         = set(st.session_state.interactif_done.get(combo_key, []))
                indices_restants = [i for i in range(nb_interactifs) if i not in done_set]
                nb_restants      = len(indices_restants)

                col_next, col_reset = st.columns(2)
                with col_next:
                    label_next = (f"➡️ Exercice suivant ({nb_restants} restant{'s' if nb_restants > 1 else ''})"
                                  if nb_restants > 0 else "🔄 Recommencer depuis le début")
                    if st.button(label_next, type="primary", use_container_width=True,
                                 key="btn_suivant_i"):
                        if nb_restants > 0:
                            idx = random.choice(indices_restants)
                            done_set.add(idx)
                            st.session_state.interactif_done[combo_key] = list(done_set)
                            sujet_suivant = sujets_interactifs[idx]
                        else:
                            st.session_state.interactif_done[combo_key] = []
                            sujet_suivant = random.choice(sujets_interactifs)
                        st.session_state.interactif_sujet     = sujet_suivant
                        st.session_state.interactif_idx       = 0
                        st.session_state.interactif_reponses  = {}
                        st.session_state.interactif_termine   = False
                        st.session_state.interactif_eval_done = None
                        st.rerun()
                with col_reset:
                    if st.button("🏠 Retour au menu", use_container_width=True, key="btn_reset_i"):
                        st.session_state.interactif_sujet     = None
                        st.session_state.interactif_reponses  = {}
                        st.session_state.interactif_termine   = False
                        st.session_state.interactif_eval_done = None
                        st.rerun()

    # ── SUJET CLASSIQUE ───────────────────────────────────────
    if st.session_state.generated_md:
        m = st.session_state.meta_gen or {}
        st.divider()
        badges_html = (
            f'<span class="badge">📚 {m.get("matiere","")}</span>'
            f'<span class="badge">🎓 {m.get("niveau","")}</span>'
            f'<span class="badge">📖 {m.get("chapitre","")}</span>'
        )
        if m.get("filiere"):   badges_html += f'<span class="badge">🏢 {m["filiere"]}</span>'
        if m.get("difficulte"):badges_html += f'<span class="badge">{m["difficulte"]}</span>'
        st.markdown(badges_html, unsafe_allow_html=True)
        st.markdown(st.session_state.generated_md)

        st.divider()
        st.markdown("**🎯 Comment tu t'en es sorti ?**")
        col_e1, col_e2, col_e3, col_e4 = st.columns(4)
        eval_choix = None
        with col_e1:
            if st.button("😕 Difficile", use_container_width=True, key="eval_gen_1"): eval_choix = "😕 Difficile"
        with col_e2:
            if st.button("😐 Moyen",     use_container_width=True, key="eval_gen_2"): eval_choix = "😐 Moyen"
        with col_e3:
            if st.button("😊 Bien",      use_container_width=True, key="eval_gen_3"): eval_choix = "😊 Bien"
        with col_e4:
            if st.button("🌟 Très bien", use_container_width=True, key="eval_gen_4"): eval_choix = "🌟 Très bien"

        if eval_choix:
            envoyer_grist(code_eleve or "anonyme", "Exercice", m, eval_choix,
                          genre=st.session_state.genre_pref,
                          avatar=st.session_state.avatar_pref)
            if m.get("source") == "Gemini":
                if envoyer_proposition_grist(code_eleve or "anonyme", m,
                                             st.session_state.generated_md, eval_choix):
                    st.success("📤 Ton sujet Gemini proposé au prof ! 👏")
            st.session_state.eval_gen_done = eval_choix
            st.session_state.streak_cache  = None
            if eval_choix in ("😊 Bien", "🌟 Très bien") and code_eleve:
                records_check = lire_progression_grist(code_eleve)
                records_check.append({"chapitre": m.get("chapitre",""),
                                      "niveau_difficulte": m.get("difficulte",""),
                                      "auto_evaluation": eval_choix,
                                      "type_activite": "Exercice"})
                prog        = calculer_progression(records_check, m.get("chapitre",""))
                diff_actuel = m.get("difficulte","")
                if (prog.get(diff_actuel, {}).get("valide") and
                        not prog.get(diff_actuel, {}).get("boss_vaincu") and
                        diff_actuel in ORDRE_NIVEAUX_DIFF):
                    st.session_state.boss_actif    = True
                    st.session_state.boss_niveau   = diff_actuel
                    st.session_state.boss_chapitre = m.get("chapitre","")
            st.rerun()

        if st.session_state.eval_gen_done:
            st.markdown(f'<div class="ok-box">✅ Auto-évaluation enregistrée : '
                        f'<strong>{st.session_state.eval_gen_done}</strong> — continue comme ça !</div>',
                        unsafe_allow_html=True)
            if st.session_state.streak_cache is None:
                s_data = calculer_streak(lire_progression_grist(code_eleve or ""))
                st.session_state.streak_cache = s_data
                if s_data.get("nouveau_record") and s_data.get("streak", 0) >= 3:
                    st.snow()

        # ── BOSS DÉBLOQUÉ ─────────────────────────────────────
        if st.session_state.boss_actif and st.session_state.boss_chapitre == m.get("chapitre",""):
            diff_boss = st.session_state.boss_niveau
            mascotte  = MASCOTTES.get(diff_boss, {})
            st.markdown("---")
            st.markdown(
                f'<div class="boss-banner">'
                f'<div style="font-size:4rem;margin-bottom:8px">{mascotte.get("animal","⚔️")}</div>'
                f'<div style="font-family:Outfit,sans-serif;font-size:1.5rem;font-weight:800;color:#e9d5ff">BOSS DÉBLOQUÉ !</div>'
                f'<div style="font-size:1.1rem;font-weight:700;color:#c084fc;margin:6px 0">{mascotte.get("nom","Boss").upper()} t\'attend…</div>'
                f'<div style="font-size:.9rem;color:#94a3b8;max-width:400px;margin:0 auto">'
                f'Tu as validé le niveau <strong style="color:#a5b4fc">{diff_boss.split(" ",1)[-1]}</strong> sur '
                f'<strong style="color:#a5b4fc">{m.get("chapitre","")[:45]}</strong>.<br>'
                f'Bats ce boss pour débloquer le niveau suivant !</div></div>',
                unsafe_allow_html=True)
            col_boss1, col_boss2 = st.columns(2)
            with col_boss1:
                if st.button(f"⚔️ Affronter le {mascotte.get('nom','Boss')} !",
                             type="primary", use_container_width=True, key="btn_boss"):
                    if not cle_api:
                        st.error("🔑 Clé API manquante !")
                    else:
                        with st.spinner(f"⚔️ Le {mascotte.get('nom','Boss')} se prépare…"):
                            try:
                                res_boss = call_gemini(cle_api, build_prompt_boss(
                                    niv, filiere,
                                    st.session_state.boss_chapitre,
                                    st.session_state.boss_niveau))
                                st.session_state.boss_md        = res_boss
                                st.session_state.eval_boss_done = None
                                st.session_state.boss_actif     = False
                                st.rerun()
                            except Exception as e:
                                st.error(f"Erreur : {e}")
            with col_boss2:
                if st.button("⏭️ Plus tard", use_container_width=True, key="btn_boss_skip"):
                    st.session_state.boss_actif = False
                    st.rerun()

        # ── SUJET BOSS AFFICHÉ ────────────────────────────────
        if st.session_state.boss_md:
            diff_boss = st.session_state.boss_niveau or "🟢 Débutant"
            mascotte  = MASCOTTES.get(diff_boss, {})
            st.markdown(f'<div style="background:#1e1b4b;color:white;padding:12px 16px;'
                        f'border-radius:8px;margin:8px 0;font-size:1.1rem;font-weight:bold">'
                        f'{mascotte.get("animal","⚔️")} DÉFI BOSS — {mascotte.get("nom","Boss").upper()}'
                        f'</div>', unsafe_allow_html=True)
            st.markdown(st.session_state.boss_md)
            st.divider()
            st.markdown("**⚔️ As-tu vaincu le Boss ?**")
            col_b1, col_b2 = st.columns(2)
            with col_b1:
                if st.button("😤 Pas encore…", use_container_width=True, key="boss_fail"):
                    envoyer_grist(code_eleve or "anonyme", f"Boss-{diff_boss}",
                                  {"chapitre": st.session_state.boss_chapitre or "",
                                   "niveau": niv, "filiere": filiere, "matiere": mat},
                                  "😕 Difficile",
                                  genre=st.session_state.genre_pref,
                                  avatar=st.session_state.avatar_pref)
                    st.session_state.eval_boss_done = "fail"
                    st.session_state.boss_md        = None
                    st.rerun()
            with col_b2:
                if st.button("🏆 Boss vaincu !", use_container_width=True, key="boss_win"):
                    envoyer_grist(code_eleve or "anonyme", f"Boss-{diff_boss}",
                                  {"chapitre": st.session_state.boss_chapitre or "",
                                   "niveau": niv, "filiere": filiere, "matiere": mat},
                                  "🌟 Très bien",
                                  genre=st.session_state.genre_pref,
                                  avatar=st.session_state.avatar_pref)
                    st.session_state.eval_boss_done    = "win"
                    st.session_state.boss_md           = None
                    st.session_state.progression_cache = None
                    st.rerun()

        if st.session_state.eval_boss_done == "win":
            st.balloons()
            msg = MESSAGES_VICTOIRE.get(st.session_state.boss_niveau or "", "🏆 Félicitations !")
            st.markdown(f'<div style="background:linear-gradient(135deg,#052e16,#14532d);'
                        f'border:2px solid #22c55e;color:white;padding:20px;border-radius:14px;'
                        f'text-align:center;box-shadow:0 0 30px rgba(34,197,94,.3);margin:8px 0">'
                        f'<div style="font-size:2.5rem;margin-bottom:8px">🏆</div>'
                        f'<div style="font-size:1.2rem;font-weight:800;color:#86efac">{msg}</div>'
                        f'</div>', unsafe_allow_html=True)
        elif st.session_state.eval_boss_done == "fail":
            st.markdown('<div class="info-box">💪 Pas grave ! Continue à t\'entraîner et reviens affronter le Boss.</div>',
                        unsafe_allow_html=True)

        # ── Téléchargements ───────────────────────────────────
        titre_doc = f"Sujet — {m.get('matiere','')} {m.get('niveau','')} {m.get('difficulte','')}"
        c1, c2, c3 = st.columns(3)
        with c1:
            st.download_button("📥 .md",  st.session_state.generated_md,
                               file_name="sujet.md",  mime="text/markdown", key="dl1_md")
        with c2:
            st.download_button("📄 .txt", st.session_state.generated_md,
                               file_name="sujet.txt", mime="text/plain",    key="dl1_txt")
        with c3:
            st.download_button("📝 .docx", markdown_to_docx(st.session_state.generated_md, titre_doc),
                               file_name="sujet.docx",
                               mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                               key="dl1_docx")
                              
# ============================================================
# app_eleve.py — Interface élève
# Toute la logique métier est dans utils.py
# ============================================================

import random
import streamlit as st
from utils import (
    # Constantes
    THEMES, NIVEAUX_CATEGORIES, MATIERES, LISTE_FILIERES, CONTEXTES_FILIERES,
    NIVEAUX_DIFFICULTE, ORDRE_NIVEAUX_DIFF, MASCOTTES, MESSAGES_VICTOIRE,
    BADGES, XP_PAR_ACTION, RANGS, SEUIL_VALIDATION,
    # Helpers chapitres
    get_chapitres, build_contexte_filiere,
    # Prompts
    build_prompt_exercices, build_prompt_ccf_entrainement,
    build_prompt_ccf_officiel, build_prompt_boss,
    # Gemini
    call_gemini,
    # Grist
    envoyer_grist, envoyer_proposition_grist,
    lire_progression_grist, lire_classement_grist,
    # Fonctions métier
    calculer_xp, calculer_streak, calculer_progression,
    calculer_objectif, calculer_badges, niveau_suivant,
    # Banque
    charger_banque, charger_banque_interactif,
    # Export Word
    markdown_to_docx, generate_ccf_officiel_docx,
)

# ============================================================
# 1. CONFIG — doit être le PREMIER appel Streamlit
# ============================================================
st.set_page_config(
    page_title="Entraînement Bac Pro — Maths",
    page_icon="📚",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ============================================================
# 2. THÈME — sélecteur + CSS
# ============================================================
if "theme_pref" not in st.session_state:
    st.session_state.theme_pref = "🌌 Nébuleuse"

theme_nom = st.sidebar.selectbox(
    "🎨 Style de l'interface",
    list(THEMES.keys()),
    index=list(THEMES.keys()).index(st.session_state.theme_pref)
)
st.session_state.theme_pref = theme_nom
t = THEMES[theme_nom]

st.markdown(f"""<style>
@import url('https://fonts.googleapis.com/css2?family=Outfit:wght@400;600;700;800&display=swap');
.stApp {{ background:{t['bg_app']}; color:{t['text']}; font-family:'Outfit',sans-serif; }}
.stApp > header {{ background:transparent !important; }}
[data-testid="stSidebar"] {{ background:{t['bg_side']} !important; border-right:1px solid {t['border']}; }}
[data-testid="stSidebar"] * {{ color:{t['text']} !important; }}
.stWidgetLabel p, div[data-testid="stMarkdownContainer"] p {{ color:{t['text']} !important; font-weight:600; }}
div[data-testid="stRadio"] div[role="radiogroup"] div[aria-checked="true"] > div {{ background-color:{t['primary']} !important; border-color:{t['primary']} !important; }}
.stTabs [data-baseweb="tab-list"] {{ background:{t['bg_side']}; border-radius:12px; padding:4px; gap:4px; border:1px solid {t['border']}; }}
.stTabs [data-baseweb="tab"] {{ background:transparent; color:{t['text']}; opacity:.7; border-radius:8px; font-weight:600; }}
.stTabs [aria-selected="true"] {{ background:linear-gradient(135deg,{t['primary']},{t['secondary']}) !important; color:white !important; box-shadow:0 0 16px {t['primary']}80; }}
.stButton > button[kind="primary"] {{ background:linear-gradient(135deg,{t['primary']},{t['secondary']}); border:none; border-radius:12px; font-weight:700; color:white; box-shadow:0 0 20px {t['primary']}66; transition:all .25s; }}
.stButton > button[kind="primary"]:hover {{ transform:translateY(-2px); box-shadow:0 0 30px {t['primary']}B3; }}
.stSelectbox > div > div, .stTextInput > div > div > input {{ background:{t['bg_side']} !important; border:1px solid {t['border']} !important; color:{t['text']} !important; border-radius:10px !important; }}
[data-testid="stMetric"] {{ background:{t['bg_side']}; border:1px solid {t['border']}; border-radius:12px; padding:16px; }}
[data-testid="stMetricValue"] {{ color:{t['accent']} !important; font-weight:800; }}
.xp-banner {{ background:linear-gradient(135deg,{t['bg_side']},{t['bg_app']}); border:1px solid {t['border']}; border-radius:14px; padding:14px 20px; margin-bottom:16px; display:flex; align-items:center; gap:16px; }}
.xp-value {{ font-family:'Outfit',sans-serif; font-weight:800; font-size:1.4rem; color:{t['primary']}; }}
.boss-banner {{ background:linear-gradient(135deg,{t['bg_side']},{t['secondary']}44); border:2px solid {t['primary']}; border-radius:16px; padding:24px; text-align:center; margin:16px 0; box-shadow:0 0 40px {t['primary']}33; }}
::-webkit-scrollbar {{ width:6px; }}
::-webkit-scrollbar-track {{ background:{t['bg_app']}; }}
::-webkit-scrollbar-thumb {{ background:{t['border']}; border-radius:3px; }}
::-webkit-scrollbar-thumb:hover {{ background:{t['primary']}; }}
.stMarkdown h1,.stMarkdown h2,.stMarkdown h3 {{ color:{t['text']}; }}
.stMarkdown p,.stMarkdown li {{ color:{t['text']}; opacity:.9; }}
.stCheckbox > label {{ color:#94a3b8 !important; }}
.stRadio > label {{ color:#94a3b8 !important; }}
.stExpander {{ background:#13162a !important; border:1px solid #2d3561 !important; border-radius:10px !important; }}
.info-box {{ background:#f0f4ff; border-left:4px solid #4a6cf7; padding:12px 16px; border-radius:4px; margin:8px 0; font-size:.9rem; }}
.warn-box {{ background:#fff8e1; border-left:4px solid #f59e0b; padding:12px 16px; border-radius:4px; margin:8px 0; font-size:.9rem; }}
.ok-box   {{ background:#f0fdf4; border-left:4px solid #22c55e; padding:12px 16px; border-radius:4px; margin:8px 0; font-size:.9rem; }}
.badge    {{ display:inline-block; background:#e0e7ff; color:#3730a3; border-radius:6px; padding:3px 10px; font-size:.8rem; margin-right:6px; margin-bottom:4px; }}
</style>""", unsafe_allow_html=True)

# ============================================================
# 3. PRÉFÉRENCES IDENTITÉ
# ============================================================
if "genre_pref"  not in st.session_state: st.session_state.genre_pref  = "Neutre (Aventurier)"
if "avatar_pref" not in st.session_state: st.session_state.avatar_pref = "Robotique 🤖"

col1, col2 = st.columns(2)
with col1:
    genre = st.radio("Comment souhaites-tu que l'on s'adresse à toi ?",
                     ["Neutre (Aventurier)", "Féminin (Aventurière)", "Masculin (Aventurier)"],
                     key="radio_genre")
with col2:
    avatar_style = st.selectbox("Choisis ton style d'avatar",
                                ["Robotique 🤖", "Mage 🧙‍♂️", "Guerrier/ère 🛡️", "Animalier 🐾"],
                                key="select_avatar")
st.session_state.genre_pref  = genre
st.session_state.avatar_pref = avatar_style

# ============================================================
# 4. SESSION STATE
# ============================================================
for key in ["generated_md", "generated_ccf_md", "meta_gen", "meta_ccf",
            "eval_gen_done", "eval_ccf_done", "progression_cache",
            "boss_actif", "boss_niveau", "boss_chapitre", "boss_md", "eval_boss_done",
            "streak_cache", "interactif_sujet", "interactif_idx",
            "interactif_reponses", "interactif_termine", "interactif_eval_done",
            "interactif_done", "classement_cache"]:
    if key not in st.session_state:
        st.session_state[key] = None

# ============================================================
# 5. SIDEBAR
# ============================================================
with st.sidebar:
    st.markdown("""
    <div style="text-align:center;padding:16px 0 8px">
        <div style="font-size:2.5rem">🎓</div>
        <div style="font-family:'Outfit',sans-serif;font-weight:800;font-size:1.1rem;
                    color:#a5b4fc;letter-spacing:.5px">MATHS BAC PRO</div>
        <div style="font-size:.75rem;color:#475569;margin-top:2px">Entraînement IA</div>
    </div>""", unsafe_allow_html=True)
    st.divider()

    st.markdown('<div style="font-family:Outfit,sans-serif;font-weight:700;color:#94a3b8;font-size:.8rem;letter-spacing:1px;margin-bottom:6px">🎮 TON IDENTIFIANT</div>', unsafe_allow_html=True)
    code_eleve = st.text_input("Code élève", placeholder="Ex : ASSP-03", key="code_eleve",
                               label_visibility="collapsed",
                               help="Code distribué par ton professeur.").strip().upper()

    if code_eleve:
        if not st.session_state.streak_cache:
            st.session_state.streak_cache = calculer_streak(lire_progression_grist(code_eleve))
        s      = st.session_state.streak_cache or {}
        streak = s.get("streak", 0)
        record = s.get("record", 0)

        from datetime import date
        derniere = s.get("derniere_date")
        if derniere == date.today(): salut = "Bonne continuation"
        elif streak > 0:             salut = "Bon retour"
        else:                        salut = "Bienvenue"

        if streak >= 7:   feu, cb = "🔥🔥🔥", "#ef4444"
        elif streak >= 3: feu, cb = "🔥🔥",   "#f97316"
        elif streak >= 1: feu, cb = "🔥",     "#f59e0b"
        else:             feu, cb = "👋",     "#6366f1"

        streak_txt = (f"{feu} {streak} jour{'s' if streak > 1 else ''} d'affilée !"
                      if streak > 0 else "Lance ton premier streak aujourd'hui !")
        st.markdown(
            f'<div style="background:linear-gradient(135deg,#13162a,#1e2235);'
            f'border:1px solid {cb}44;border-radius:10px;padding:12px 14px;margin:6px 0">'
            f'<div style="font-size:.82rem;color:#94a3b8">{salut} 👋</div>'
            f'<div style="font-family:Outfit,sans-serif;font-weight:800;color:#e2e8f0;font-size:1rem;margin:2px 0">{code_eleve}</div>'
            f'<div style="font-size:.82rem;color:{cb};margin-top:4px">{streak_txt}</div>'
            + (f'<div style="font-size:.72rem;color:#475569;margin-top:2px">Record : {record} jour{"s" if record > 1 else ""}</div>' if record > 1 else "")
            + '</div>', unsafe_allow_html=True)
    else:
        st.markdown('<div class="warn-box">⚠️ Entre ton code pour sauvegarder ta progression.</div>', unsafe_allow_html=True)

    st.divider()
    st.markdown('<div style="font-family:Outfit,sans-serif;font-weight:700;color:#94a3b8;font-size:.8rem;letter-spacing:1px;margin-bottom:6px">🔑 CLÉ GEMINI</div>', unsafe_allow_html=True)
    cle_api = st.text_input("Clé API", type="password", key="gemini_key", label_visibility="collapsed")
    if cle_api:
        st.markdown('<div class="ok-box">✅ Prêt à générer !</div>', unsafe_allow_html=True)
    else:
        st.markdown('<div class="warn-box">⚠️ Clé API manquante.</div>', unsafe_allow_html=True)
    with st.expander("📋 Obtenir une clé gratuite (2 min)"):
        st.markdown("""**1.** → [aistudio.google.com](https://aistudio.google.com/app/apikey)
**2.** Connecte-toi avec Google  
**3.** Clique **"Create API Key"**  
**4.** Copie et colle la clé ici  
💡 **C'est gratuit** — pas de carte bancaire requise.""")
    st.divider()
    st.caption("Entraînement Bac Pro · Gemini 2.5 Flash")

# ============================================================
# 6. TITRE
# ============================================================
st.markdown("""
<div style="padding:24px 0 8px">
    <div style="font-family:'Outfit',sans-serif;font-weight:800;font-size:2rem;
                background:linear-gradient(90deg,#a5b4fc,#c084fc,#67e8f9);
                -webkit-background-clip:text;-webkit-text-fill-color:transparent;
                background-clip:text;line-height:1.2">📚 Entraînement Bac Pro</div>
    <div style="color:#475569;font-size:.9rem;margin-top:4px">Exercices · CCF · Progression — propulsé par l'IA</div>
</div>""", unsafe_allow_html=True)

# ============================================================
# 7. ONGLETS
# ============================================================
tab_gen, tab_ccf, tab_graphique, tab_progression = st.tabs([
    "📝 Exercices d'entraînement",
    "🎯 Sujets CCF",
    "📈 Graphique GeoGebra",
    "📊 Ma progression",
])

# ─────────────────────────────────────────────────────────────
# ONGLET 1 — GÉNÉRATEUR
# ─────────────────────────────────────────────────────────────
with tab_gen:
    # Objectif du jour
    if code_eleve and st.session_state.progression_cache:
        chap_actif  = (st.session_state.meta_gen or {}).get("chapitre", "")
        diff_active = (st.session_state.meta_gen or {}).get("difficulte", "")
        obj = calculer_objectif(st.session_state.progression_cache, chap_actif, diff_active)
        st.markdown(
            f'<div style="background:#0d0f1a;border:1px solid {obj["couleur"]}55;'
            f'border-left:4px solid {obj["couleur"]};border-radius:0 10px 10px 0;'
            f'padding:10px 16px;margin-bottom:12px">'
            f'<span style="font-size:1.1rem">{obj["emoji"]}</span> '
            f'<span style="color:#e2e8f0;font-size:.88rem">{obj["message"]}</span>'
            f'</div>', unsafe_allow_html=True)

    st.subheader("📝 Générateur de sujets et exercices")
    col1, col2 = st.columns(2)
    with col1:
        cat     = st.selectbox("Type d'établissement", list(NIVEAUX_CATEGORIES.keys()), key="gen_cat")
        niv     = st.selectbox("Classe", NIVEAUX_CATEGORIES[cat], key="gen_niv")
        filiere = ""
        if cat in ["Bac Pro", "CAP"]:
            fil = st.selectbox("Filière", LISTE_FILIERES, key="gen_fil")
            if fil == "Autre (Préciser ci-dessous)":
                filiere = st.text_input("Filière libre", key="gen_spec")
            else:
                filiere = fil
                if fil in CONTEXTES_FILIERES:
                    st.info(f"📌 {CONTEXTES_FILIERES[fil]['nom_complet']}")
    with col2:
        mat  = st.selectbox("Matière", MATIERES, key="gen_mat")
        chap = st.selectbox("Chapitre (BO)", get_chapitres(mat, niv, cat), key="gen_chap")

    consignes  = st.text_area("Consignes particulières (optionnel)", height=80, key="gen_consignes",
                               placeholder="Ex : 3 exercices, niveau accessible…")
    st.markdown("**🎯 Niveau de difficulté**")
    difficulte = st.select_slider("Niveau", options=list(NIVEAUX_DIFFICULTE.keys()),
                                  value="🟡 Moyen", key="gen_diff", label_visibility="collapsed")
    st.markdown(f'<div class="info-box">{difficulte} — {NIVEAUX_DIFFICULTE[difficulte]}</div>',
                unsafe_allow_html=True)

    sujets_banque      = charger_banque(niv, filiere, chap, difficulte)
    sujets_interactifs = charger_banque_interactif(niv, filiere, chap, difficulte)
    nb_banque          = len(sujets_banque)
    nb_interactifs     = len(sujets_interactifs)

    col_btn1, col_btn2, col_btn3 = st.columns(3)
    with col_btn1:
        btn_interactif = st.button(f"⚡ Mode interactif ({nb_interactifs} dispo)",
                                   type="primary", use_container_width=True,
                                   disabled=(nb_interactifs == 0), key="btn_interactif")
        if nb_interactifs == 0: st.caption("Aucun sujet interactif en banque.")
    with col_btn2:
        btn_banque = st.button(f"📚 Sujet classique ({nb_banque} dispo)",
                               type="primary", use_container_width=True,
                               disabled=(nb_banque == 0), key="btn_banque")
        if nb_banque == 0: st.caption("Aucun sujet classique en banque.")
    with col_btn3:
        btn_gemini = st.button("✨ Générer (Gemini)", type="primary",
                               use_container_width=True, key="btn_gemini")

    # ── Bouton interactif ─────────────────────────────────────
    if btn_interactif and nb_interactifs > 0:
        combo_key = f"{niv}|{filiere}|{chap}|{difficulte}"
        if not st.session_state.interactif_done:
            st.session_state.interactif_done = {}
        done_set        = set(st.session_state.interactif_done.get(combo_key, []))
        indices_dispo   = [i for i in range(nb_interactifs) if i not in done_set]
        if indices_dispo:
            idx = random.choice(indices_dispo)
            done_set.add(idx)
            st.session_state.interactif_done[combo_key] = list(done_set)
            st.session_state.interactif_sujet    = sujets_interactifs[idx]
            st.session_state.interactif_idx      = 0
            st.session_state.interactif_reponses = {}
            st.session_state.interactif_termine  = False
            st.session_state.interactif_eval_done = None
            st.session_state.generated_md         = None
            st.session_state.eval_gen_done        = None
            st.session_state.meta_gen = {"niveau": niv, "matiere": mat, "chapitre": chap,
                                         "filiere": filiere, "difficulte": difficulte,
                                         "source": "Banque interactif"}
            st.rerun()
        else:
            st.warning(f"🎉 Tu as déjà fait les {nb_interactifs} exercice(s) disponibles !")
            c1, c2 = st.columns(2)
            with c1:
                if st.button("🔄 Recommencer depuis le début", use_container_width=True,
                             key="btn_reset_done"):
                    st.session_state.interactif_done[combo_key] = []
                    st.rerun()
            with c2:
                if st.button("✨ Générer via Gemini", use_container_width=True,
                             key="btn_gen_from_done"):
                    if not cle_api:
                        st.error("🔑 Clé API manquante !")
                    else:
                        with st.spinner("⏳ Génération…"):
                            try:
                                res = call_gemini(cle_api, build_prompt_exercices(
                                    niv, cat, mat, chap, "", filiere, difficulte))
                                st.session_state.generated_md     = res
                                st.session_state.interactif_sujet = None
                                st.session_state.eval_gen_done    = None
                                st.session_state.meta_gen = {"niveau": niv, "matiere": mat,
                                    "chapitre": chap, "filiere": filiere,
                                    "difficulte": difficulte, "source": "Gemini"}
                                st.rerun()
                            except Exception as e:
                                st.error(f"Erreur : {e}")

    # ── Bouton banque classique ────────────────────────────────
    if btn_banque and nb_banque > 0:
        sujet = random.choice(sujets_banque)
        st.session_state.generated_md         = sujet["contenu"]
        st.session_state.interactif_sujet     = None
        st.session_state.eval_gen_done        = None
        st.session_state.meta_gen = {"niveau": niv, "matiere": mat, "chapitre": chap,
                                     "filiere": filiere, "difficulte": difficulte,
                                     "source": "Banque"}
        st.rerun()

    # ── Bouton Gemini ─────────────────────────────────────────
    if btn_gemini:
        if not cle_api:
            st.error("🔑 Renseigne ta clé API dans le panneau gauche !")
        else:
            with st.spinner("⏳ Génération en cours…"):
                try:
                    res = call_gemini(cle_api, build_prompt_exercices(
                        niv, cat, mat, chap, consignes, filiere, difficulte))
                    st.session_state.generated_md     = res
                    st.session_state.interactif_sujet = None
                    st.session_state.eval_gen_done    = None
                    st.session_state.meta_gen = {"niveau": niv, "matiere": mat,
                        "chapitre": chap, "filiere": filiere,
                        "difficulte": difficulte, "source": "Gemini"}
                    envoyer_proposition_grist(code_eleve or "anonyme",
                                              st.session_state.meta_gen, res)
                    st.success("✅ Sujet généré !")
                except Exception as e:
                    if "429" in str(e):
                        st.error("⏱️ Quota dépassé. Attends 1 minute ou utilise la banque.")
                    else:
                        st.error(f"Erreur API : {e}")

    # ── MODE INTERACTIF ───────────────────────────────────────
    if st.session_state.get("interactif_sujet"):
        sujet_i   = st.session_state.interactif_sujet
        questions = sujet_i.get("questions", [])
        idx       = st.session_state.get("interactif_idx", 0)
        reponses  = st.session_state.get("interactif_reponses", {})
        termine   = st.session_state.get("interactif_termine", False)

        if sujet_i.get("contexte"):
            st.markdown(f'<div class="info-box">🏢 <strong>Mise en situation</strong><br>{sujet_i["contexte"]}</div>',
                        unsafe_allow_html=True)
        if sujet_i.get("rappel_cours"):
            st.markdown(f'<div style="background:#1a2040;border:1px solid #3d4480;border-radius:10px;'
                        f'padding:12px 16px;margin:8px 0;font-size:.88rem;color:#a5b4fc">'
                        f'📐 <strong>Rappel</strong> : {sujet_i["rappel_cours"]}</div>',
                        unsafe_allow_html=True)

        if not termine:
            st.markdown(f'<div style="margin:12px 0 4px;font-size:.82rem;color:#64748b">'
                        f'Question {idx + 1} / {len(questions)}</div>', unsafe_allow_html=True)
            st.progress(idx / len(questions))
            q = questions[idx]
            st.markdown(f'<div style="background:#13162a;border:1px solid #3d4480;border-radius:12px;'
                        f'padding:18px 20px;margin:12px 0">'
                        f'<div style="font-size:.8rem;color:#6366f1;font-weight:700;margin-bottom:6px">'
                        f'QUESTION {q["numero"]}</div>'
                        f'<div style="color:#e2e8f0;font-size:.95rem">{q["enonce"]}</div>'
                        f'</div>', unsafe_allow_html=True)

            col_rep, col_ind = st.columns([3, 1])
            with col_rep:
                unite      = q.get("unite", "")
                rep_saisie = st.number_input(f"Ta réponse {f'({unite})' if unite else ''}",
                                             value=0.0, step=0.01, format="%.2f",
                                             key=f"rep_q_{idx}")
            with col_ind:
                if q.get("indice") and st.button("💡 Indice", key=f"ind_{idx}",
                                                  use_container_width=True):
                    st.info(q["indice"])

            col_v, col_p = st.columns(2)
            with col_v:
                if st.button("✅ Valider", type="primary", use_container_width=True,
                             key=f"val_{idx}"):
                    tolerance = float(q.get("tolerance", 0.5))
                    bonne     = abs(rep_saisie - float(q["reponse"])) <= tolerance
                    reponses[idx] = {"saisie": rep_saisie, "bonne": bonne,
                                     "reponse": q["reponse"], "unite": unite,
                                     "corrige": q.get("corrige", "")}
                    st.session_state.interactif_reponses = reponses
                    if idx + 1 < len(questions):
                        st.session_state.interactif_idx = idx + 1
                    else:
                        st.session_state.interactif_termine = True
                    st.rerun()
            with col_p:
                if idx > 0 and st.button("⬅️ Précédente", use_container_width=True,
                                          key=f"prev_{idx}"):
                    st.session_state.interactif_idx = idx - 1
                    st.rerun()
        else:
            # ── RÉSULTATS FINAUX ──────────────────────────────
            nb_bonnes_i = sum(1 for r in reponses.values() if r["bonne"])
            nb_total_i  = len(questions)
            score_pct   = int(nb_bonnes_i / nb_total_i * 100) if nb_total_i else 0

            if score_pct == 100: st.balloons(); verdict, couleur = "🏆 Parfait !",               "#065f46"
            elif score_pct >= 70:               verdict, couleur = "👍 Bien joué !",              "#1e3a5f"
            elif score_pct >= 50:               verdict, couleur = "😐 Peut mieux faire",         "#3f2a00"
            else:                               verdict, couleur = "💪 Continue à t'entraîner !", "#3b0000"

            st.markdown(f'<div style="background:{couleur};border-radius:14px;padding:20px;'
                        f'text-align:center;margin:12px 0">'
                        f'<div style="font-size:2rem">{verdict}</div>'
                        f'<div style="font-size:1.5rem;font-weight:800;color:white;margin:8px 0">'
                        f'{nb_bonnes_i} / {nb_total_i} — {score_pct}%</div>'
                        f'</div>', unsafe_allow_html=True)

            # Étape 1 : Ressenti AVANT la correction
            if not st.session_state.interactif_eval_done:
                st.divider()
                st.markdown("**🎯 Et ton ressenti avant de voir la correction ?**")
                col_e1, col_e2, col_e3, col_e4 = st.columns(4)
                eval_choix_i = None
                with col_e1:
                    if st.button("😕 Difficile", use_container_width=True, key="eval_i_1"): eval_choix_i = "😕 Difficile"
                with col_e2:
                    if st.button("😐 Moyen",     use_container_width=True, key="eval_i_2"): eval_choix_i = "😐 Moyen"
                with col_e3:
                    if st.button("😊 Bien",      use_container_width=True, key="eval_i_3"): eval_choix_i = "😊 Bien"
                with col_e4:
                    if st.button("🌟 Très bien", use_container_width=True, key="eval_i_4"): eval_choix_i = "🌟 Très bien"
                if eval_choix_i:
                    m_i = st.session_state.meta_gen or {}
                    envoyer_grist(code_eleve or "anonyme",
                                  f"Exercice-interactif ({score_pct}%)",
                                  {**m_i, "score_auto": f"{nb_bonnes_i}/{nb_total_i}"},
                                  eval_choix_i,
                                  genre=st.session_state.genre_pref,
                                  avatar=st.session_state.avatar_pref)
                    st.session_state.interactif_eval_done = eval_choix_i
                    st.session_state.streak_cache         = None
                    st.rerun()
            else:
                # Étape 2 : Correction + bouton suivant
                st.markdown(f'<div class="ok-box">✅ Ressenti enregistré : '
                            f'<strong>{st.session_state.interactif_eval_done}</strong></div>',
                            unsafe_allow_html=True)
                st.markdown("### 📋 Correction détaillée")
                for i, q in enumerate(questions):
                    r     = reponses.get(i, {})
                    icone = "✅" if r.get("bonne") else "❌"
                    unite = r.get("unite", "")
                    st.markdown(
                        f'<div style="background:#13162a;border-left:3px solid '
                        f'{"#22c55e" if r.get("bonne") else "#ef4444"};'
                        f'border-radius:0 10px 10px 0;padding:12px 16px;margin:6px 0">'
                        f'<div style="font-size:.8rem;color:#64748b">Q{q["numero"]}</div>'
                        f'<div style="color:#e2e8f0;margin:4px 0">{q["enonce"]}</div>'
                        f'<div style="font-size:.88rem;margin-top:6px">'
                        f'{icone} Ta réponse : <strong>{r.get("saisie","?"):.2f} {unite}</strong> — '
                        f'Bonne réponse : <strong>{float(q["reponse"]):.2f} {unite}</strong></div>'
                        f'<div style="font-size:.82rem;color:#94a3b8;margin-top:4px">'
                        f'💡 {r.get("corrige","")}</div></div>', unsafe_allow_html=True)

                st.divider()
                m_i       = st.session_state.meta_gen or {}
                combo_key = f"{m_i.get('niveau')}|{m_i.get('filiere')}|{m_i.get('chapitre')}|{m_i.get('difficulte')}"
                if not st.session_state.interactif_done:
                    st.session_state.interactif_done = {}
                done_set         = set(st.session_state.interactif_done.get(combo_key, []))
                indices_restants = [i for i in range(nb_interactifs) if i not in done_set]
                nb_restants      = len(indices_restants)

                col_next, col_reset = st.columns(2)
                with col_next:
                    label_next = (f"➡️ Exercice suivant ({nb_restants} restant{'s' if nb_restants > 1 else ''})"
                                  if nb_restants > 0 else "🔄 Recommencer depuis le début")
                    if st.button(label_next, type="primary", use_container_width=True,
                                 key="btn_suivant_i"):
                        if nb_restants > 0:
                            idx = random.choice(indices_restants)
                            done_set.add(idx)
                            st.session_state.interactif_done[combo_key] = list(done_set)
                            sujet_suivant = sujets_interactifs[idx]
                        else:
                            st.session_state.interactif_done[combo_key] = []
                            sujet_suivant = random.choice(sujets_interactifs)
                        st.session_state.interactif_sujet     = sujet_suivant
                        st.session_state.interactif_idx       = 0
                        st.session_state.interactif_reponses  = {}
                        st.session_state.interactif_termine   = False
                        st.session_state.interactif_eval_done = None
                        st.rerun()
                with col_reset:
                    if st.button("🏠 Retour au menu", use_container_width=True, key="btn_reset_i"):
                        st.session_state.interactif_sujet     = None
                        st.session_state.interactif_reponses  = {}
                        st.session_state.interactif_termine   = False
                        st.session_state.interactif_eval_done = None
                        st.rerun()

    # ── SUJET CLASSIQUE ───────────────────────────────────────
    if st.session_state.generated_md:
        m = st.session_state.meta_gen or {}
        st.divider()
        badges_html = (
            f'<span class="badge">📚 {m.get("matiere","")}</span>'
            f'<span class="badge">🎓 {m.get("niveau","")}</span>'
            f'<span class="badge">📖 {m.get("chapitre","")}</span>'
        )
        if m.get("filiere"):   badges_html += f'<span class="badge">🏢 {m["filiere"]}</span>'
        if m.get("difficulte"):badges_html += f'<span class="badge">{m["difficulte"]}</span>'
        st.markdown(badges_html, unsafe_allow_html=True)
        st.markdown(st.session_state.generated_md)

        st.divider()
        st.markdown("**🎯 Comment tu t'en es sorti ?**")
        col_e1, col_e2, col_e3, col_e4 = st.columns(4)
        eval_choix = None
        with col_e1:
            if st.button("😕 Difficile", use_container_width=True, key="eval_gen_1"): eval_choix = "😕 Difficile"
        with col_e2:
            if st.button("😐 Moyen",     use_container_width=True, key="eval_gen_2"): eval_choix = "😐 Moyen"
        with col_e3:
            if st.button("😊 Bien",      use_container_width=True, key="eval_gen_3"): eval_choix = "😊 Bien"
        with col_e4:
            if st.button("🌟 Très bien", use_container_width=True, key="eval_gen_4"): eval_choix = "🌟 Très bien"

        if eval_choix:
            envoyer_grist(code_eleve or "anonyme", "Exercice", m, eval_choix,
                          genre=st.session_state.genre_pref,
                          avatar=st.session_state.avatar_pref)
            if m.get("source") == "Gemini":
                if envoyer_proposition_grist(code_eleve or "anonyme", m,
                                             st.session_state.generated_md, eval_choix):
                    st.success("📤 Ton sujet Gemini proposé au prof ! 👏")
            st.session_state.eval_gen_done = eval_choix
            st.session_state.streak_cache  = None
            if eval_choix in ("😊 Bien", "🌟 Très bien") and code_eleve:
                records_check = lire_progression_grist(code_eleve)
                records_check.append({"chapitre": m.get("chapitre",""),
                                      "niveau_difficulte": m.get("difficulte",""),
                                      "auto_evaluation": eval_choix,
                                      "type_activite": "Exercice"})
                prog        = calculer_progression(records_check, m.get("chapitre",""))
                diff_actuel = m.get("difficulte","")
                if (prog.get(diff_actuel, {}).get("valide") and
                        not prog.get(diff_actuel, {}).get("boss_vaincu") and
                        diff_actuel in ORDRE_NIVEAUX_DIFF):
                    st.session_state.boss_actif    = True
                    st.session_state.boss_niveau   = diff_actuel
                    st.session_state.boss_chapitre = m.get("chapitre","")
            st.rerun()

        if st.session_state.eval_gen_done:
            st.markdown(f'<div class="ok-box">✅ Auto-évaluation enregistrée : '
                        f'<strong>{st.session_state.eval_gen_done}</strong> — continue comme ça !</div>',
                        unsafe_allow_html=True)
            if st.session_state.streak_cache is None:
                s_data = calculer_streak(lire_progression_grist(code_eleve or ""))
                st.session_state.streak_cache = s_data
                if s_data.get("nouveau_record") and s_data.get("streak", 0) >= 3:
                    st.snow()

        # ── BOSS DÉBLOQUÉ ─────────────────────────────────────
        if st.session_state.boss_actif and st.session_state.boss_chapitre == m.get("chapitre",""):
            diff_boss = st.session_state.boss_niveau
            mascotte  = MASCOTTES.get(diff_boss, {})
            st.markdown("---")
            st.markdown(
                f'<div class="boss-banner">'
                f'<div style="font-size:4rem;margin-bottom:8px">{mascotte.get("animal","⚔️")}</div>'
                f'<div style="font-family:Outfit,sans-serif;font-size:1.5rem;font-weight:800;color:#e9d5ff">BOSS DÉBLOQUÉ !</div>'
                f'<div style="font-size:1.1rem;font-weight:700;color:#c084fc;margin:6px 0">{mascotte.get("nom","Boss").upper()} t\'attend…</div>'
                f'<div style="font-size:.9rem;color:#94a3b8;max-width:400px;margin:0 auto">'
                f'Tu as validé le niveau <strong style="color:#a5b4fc">{diff_boss.split(" ",1)[-1]}</strong> sur '
                f'<strong style="color:#a5b4fc">{m.get("chapitre","")[:45]}</strong>.<br>'
                f'Bats ce boss pour débloquer le niveau suivant !</div></div>',
                unsafe_allow_html=True)
            col_boss1, col_boss2 = st.columns(2)
            with col_boss1:
                if st.button(f"⚔️ Affronter le {mascotte.get('nom','Boss')} !",
                             type="primary", use_container_width=True, key="btn_boss"):
                    if not cle_api:
                        st.error("🔑 Clé API manquante !")
                    else:
                        with st.spinner(f"⚔️ Le {mascotte.get('nom','Boss')} se prépare…"):
                            try:
                                res_boss = call_gemini(cle_api, build_prompt_boss(
                                    niv, filiere,
                                    st.session_state.boss_chapitre,
                                    st.session_state.boss_niveau))
                                st.session_state.boss_md        = res_boss
                                st.session_state.eval_boss_done = None
                                st.session_state.boss_actif     = False
                                st.rerun()
                            except Exception as e:
                                st.error(f"Erreur : {e}")
            with col_boss2:
                if st.button("⏭️ Plus tard", use_container_width=True, key="btn_boss_skip"):
                    st.session_state.boss_actif = False
                    st.rerun()

        # ── SUJET BOSS AFFICHÉ ────────────────────────────────
        if st.session_state.boss_md:
            diff_boss = st.session_state.boss_niveau or "🟢 Débutant"
            mascotte  = MASCOTTES.get(diff_boss, {})
            st.markdown(f'<div style="background:#1e1b4b;color:white;padding:12px 16px;'
                        f'border-radius:8px;margin:8px 0;font-size:1.1rem;font-weight:bold">'
                        f'{mascotte.get("animal","⚔️")} DÉFI BOSS — {mascotte.get("nom","Boss").upper()}'
                        f'</div>', unsafe_allow_html=True)
            st.markdown(st.session_state.boss_md)
            st.divider()
            st.markdown("**⚔️ As-tu vaincu le Boss ?**")
            col_b1, col_b2 = st.columns(2)
            with col_b1:
                if st.button("😤 Pas encore…", use_container_width=True, key="boss_fail"):
                    envoyer_grist(code_eleve or "anonyme", f"Boss-{diff_boss}",
                                  {"chapitre": st.session_state.boss_chapitre or "",
                                   "niveau": niv, "filiere": filiere, "matiere": mat},
                                  "😕 Difficile",
                                  genre=st.session_state.genre_pref,
                                  avatar=st.session_state.avatar_pref)
                    st.session_state.eval_boss_done = "fail"
                    st.session_state.boss_md        = None
                    st.rerun()
            with col_b2:
                if st.button("🏆 Boss vaincu !", use_container_width=True, key="boss_win"):
                    envoyer_grist(code_eleve or "anonyme", f"Boss-{diff_boss}",
                                  {"chapitre": st.session_state.boss_chapitre or "",
                                   "niveau": niv, "filiere": filiere, "matiere": mat},
                                  "🌟 Très bien",
                                  genre=st.session_state.genre_pref,
                                  avatar=st.session_state.avatar_pref)
                    st.session_state.eval_boss_done    = "win"
                    st.session_state.boss_md           = None
                    st.session_state.progression_cache = None
                    st.rerun()

        if st.session_state.eval_boss_done == "win":
            st.balloons()
            msg = MESSAGES_VICTOIRE.get(st.session_state.boss_niveau or "", "🏆 Félicitations !")
            st.markdown(f'<div style="background:linear-gradient(135deg,#052e16,#14532d);'
                        f'border:2px solid #22c55e;color:white;padding:20px;border-radius:14px;'
                        f'text-align:center;box-shadow:0 0 30px rgba(34,197,94,.3);margin:8px 0">'
                        f'<div style="font-size:2.5rem;margin-bottom:8px">🏆</div>'
                        f'<div style="font-size:1.2rem;font-weight:800;color:#86efac">{msg}</div>'
                        f'</div>', unsafe_allow_html=True)
        elif st.session_state.eval_boss_done == "fail":
            st.markdown('<div class="info-box">💪 Pas grave ! Continue à t\'entraîner et reviens affronter le Boss.</div>',
                        unsafe_allow_html=True)

        # ── Téléchargements ───────────────────────────────────
        titre_doc = f"Sujet — {m.get('matiere','')} {m.get('niveau','')} {m.get('difficulte','')}"
        c1, c2, c3 = st.columns(3)
        with c1:
            st.download_button("📥 .md",  st.session_state.generated_md,
                               file_name="sujet.md",  mime="text/markdown", key="dl1_md")
        with c2:
            st.download_button("📄 .txt", st.session_state.generated_md,
                               file_name="sujet.txt", mime="text/plain",    key="dl1_txt")
        with c3:
            st.download_button("📝 .docx", markdown_to_docx(st.session_state.generated_md, titre_doc),
                               file_name="sujet.docx",
                               mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                               key="dl1_docx")