import streamlit as st
from utils import (
    NIVEAUX_CATEGORIES, MATIERES, LISTE_FILIERES, CONTEXTES_FILIERES, 
    NIVEAUX_DIFFICULTE, get_chapitres,
    CONTEXTES_PAR_FAMILLE,
    build_prompt_exercices, build_prompt_ccf_entrainement, build_prompt_ccf_officiel,
    call_gemini, markdown_to_docx, generate_ccf_officiel_docx,
)




# ============================================================
# CO-INTERVENTION — Référentiels pro par filière
# ============================================================

THEMES_COINTERVENTION = {
    "MCVB": {
        "nom_complet": "Métiers du Commerce et de la Vente — option B",
        "themes": {
            "Prix d'achat, prix de vente et marge commerciale": {
                "formules": "Marge = PV HT − PA HT  ·  Taux de marge = Marge / PA HT  ·  Taux de marque = Marge / PV HT",
                "contexte": "Un commercial calcule sa marge sur un produit, compare deux fournisseurs.",
                "documents": "bon de commande, catalogue fournisseur, tableau comparatif"
            },
            "Taux de remise et prix remisé": {
                "formules": "Prix remisé = Prix initial × (1 − taux)  ·  Montant remise = Prix initial × taux",
                "contexte": "Négociation client, grille de remises selon quantité commandée.",
                "documents": "grille tarifaire, bon de commande, facture proforma"
            },
            "TVA — calcul HT / TTC": {
                "formules": "TTC = HT × (1 + taux TVA)  ·  HT = TTC / (1 + taux TVA)  ·  TVA = TTC − HT",
                "contexte": "Établissement de devis, vérification de factures clients et fournisseurs.",
                "documents": "facture, devis, ticket de caisse"
            },
            "Taux d'évolution du chiffre d'affaires": {
                "formules": "Taux d'évolution = (VF − VI) / VI  ·  VF = VI × (1 + t)",
                "contexte": "Évolution du CA, suivi des ventes mensuelles, comparaison N/N-1.",
                "documents": "tableau de bord commercial, graphique de ventes"
            },
            "Coefficient multiplicateur": {
                "formules": "CM = PV TTC / PA HT  ·  PV TTC = PA HT × CM",
                "contexte": "Calcul rapide du prix de vente à partir du prix d'achat.",
                "documents": "catalogue, liste de prix, étiquettes"
            },
            "Commissionnement et salaire variable": {
                "formules": "Commission = CA réalisé × taux  ·  Salaire total = Fixe + Commission",
                "contexte": "Calcul de la rémunération d'un commercial selon ses performances.",
                "documents": "bulletin de salaire simplifié, tableau de suivi CA"
            },
            "Budget prévisionnel et écarts": {
                "formules": "Écart = Réalisé − Prévu  ·  Taux d'écart = Écart / Prévu",
                "contexte": "Analyse des résultats commerciaux vs objectifs fixés.",
                "documents": "tableau de bord, fiche de suivi objectifs"
            },
        }
    },
    "MCVA": {
        "nom_complet": "Métiers du Commerce et de la Vente — option A",
        "themes": {
            "TVA — calcul HT / TTC": {
                "formules": "TTC = HT × (1 + taux TVA)  ·  HT = TTC / (1 + taux TVA)  ·  TVA = TTC − HT",
                "contexte": "Gestion des prix en rayon, vérification des étiquettes, encaissement.",
                "documents": "étiquettes de rayon, facture fournisseur, ticket de caisse"
            },
            "Coefficient multiplicateur": {
                "formules": "CM = PV TTC / PA HT  ·  PV TTC = PA HT × CM",
                "contexte": "Calcul du prix de vente d'un produit à partir du tarif fournisseur.",
                "documents": "catalogue fournisseur, liste de prix magasin"
            },
            "Taux de marque et marge": {
                "formules": "Taux de marque = Marge / PV HT  ·  Marge = PV HT − PA HT",
                "contexte": "Analyse de la rentabilité d'un rayon ou d'une famille de produits.",
                "documents": "tableau de gestion rayon, inventaire"
            },
            "Gestion des stocks — taux de rotation": {
                "formules": "Taux de rotation = CA / Stock moyen  ·  Stock moyen = (SI + SF) / 2",
                "contexte": "Suivi des entrées/sorties de marchandises, optimisation du réapprovisionnement.",
                "documents": "bon de livraison, fiche de stock, inventaire"
            },
            "Taux de transformation et panier moyen": {
                "formules": "Taux de transformation = Nb acheteurs / Nb visiteurs  ·  Panier moyen = CA / Nb acheteurs",
                "contexte": "Analyse de la performance d'un point de vente.",
                "documents": "comptage client, tableau de bord magasin"
            },
            "Surfaces et linéaires": {
                "formules": "Surface de vente = longueur × largeur  ·  Linéaire développé = nb niveaux × longueur",
                "contexte": "Implantation d'un rayon, calcul de la surface allouée à une famille de produits.",
                "documents": "plan de masse, schéma d'implantation"
            },
            "Taux d'évolution du chiffre d'affaires": {
                "formules": "Taux d'évolution = (CA N − CA N-1) / CA N-1",
                "contexte": "Comparaison des performances du magasin d'une période à l'autre.",
                "documents": "tableau de bord, rapport d'activité"
            },
        }
    },
    "AGORA": {
        "nom_complet": "Assistance à la Gestion des Organisations",
        "themes": {
            "Calcul de salaire brut et net": {
                "formules": "Salaire brut = Taux horaire × Nb heures  ·  Salaire net = Brut − Cotisations salariales",
                "contexte": "Aide à la paie, vérification de bulletins de salaire simplifiés.",
                "documents": "bulletin de salaire, contrat de travail, fiche de pointage"
            },
            "TVA — calcul HT / TTC": {
                "formules": "TTC = HT × (1 + taux TVA)  ·  HT = TTC / (1 + taux TVA)  ·  TVA collectée − TVA déductible",
                "contexte": "Établissement et vérification de factures, déclaration de TVA simplifiée.",
                "documents": "facture, bon de commande, déclaration TVA"
            },
            "Intérêts simples et emprunts": {
                "formules": "Intérêts = Capital × taux × durée  ·  Montant remboursé = Capital + Intérêts",
                "contexte": "Calcul du coût d'un emprunt bancaire pour financer du matériel.",
                "documents": "tableau d'amortissement simplifié, offre de prêt"
            },
            "Budget et suivi des dépenses": {
                "formules": "Solde = Recettes − Dépenses  ·  Taux de consommation = Dépenses / Budget prévu",
                "contexte": "Gestion du budget d'une organisation, suivi des dépenses réelles vs prévisions.",
                "documents": "tableau de bord budgétaire, relevé bancaire"
            },
            "Cotisations sociales": {
                "formules": "Cotisation = Base × taux  ·  Part salariale et part patronale",
                "contexte": "Compréhension de la fiche de paie, calcul des charges sociales.",
                "documents": "bulletin de salaire, tableau des taux de cotisation"
            },
            "Taux d'évolution et indices": {
                "formules": "Taux d'évolution = (VF − VI) / VI  ·  Indice = (Valeur / Valeur de base) × 100",
                "contexte": "Suivi de l'évolution des indicateurs RH (absentéisme, masse salariale).",
                "documents": "tableau de bord RH, rapport annuel"
            },
            "Facturation et remises": {
                "formules": "Montant net = Montant brut × (1 − remise)  ·  Escompte = Montant × taux",
                "contexte": "Établissement de devis et factures, calcul de remises et d'escomptes.",
                "documents": "devis, facture, bon de commande"
            },
        }
    },
    "ASSP": {
        "nom_complet": "Accompagnement, Soins et Services à la Personne",
        "themes": {
            "Calcul de doses médicamenteuses": {
                "formules": "Dose à administrer = Dose prescrite / Concentration  ·  Règle de trois",
                "contexte": "Préparation de médicaments, dosages pour les résidents.",
                "documents": "fiche de prescription, protocole de soins"
            },
            "Calcul de dilutions": {
                "formules": "C1 × V1 = C2 × V2  ·  Volume à prélever = (C2 × V2) / C1",
                "contexte": "Préparation de solutions désinfectantes, produits d'hygiène dilués.",
                "documents": "protocole d'hygiène, fiche technique produit"
            },
            "Gestion du temps et plannings": {
                "formules": "Durée = Heure de fin − Heure de début  ·  Répartition horaire en pourcentage",
                "contexte": "Organisation des tournées d'aide à domicile, plannings.",
                "documents": "planning hebdomadaire, feuille de route"
            },
            "Calcul de coûts de prise en charge": {
                "formules": "Coût total = Tarif horaire × Nb heures  ·  Reste à charge = Coût − Aides",
                "contexte": "Estimation du coût d'une prise en charge, calcul des aides financières.",
                "documents": "devis de prise en charge, dossier APA"
            },
            "IMC et indicateurs de santé": {
                "formules": "IMC = Poids (kg) / Taille² (m)  ·  Interprétation des valeurs normales",
                "contexte": "Suivi nutritionnel des résidents, surveillance du poids.",
                "documents": "fiche de suivi patient, courbe de poids"
            },
            "Statistiques sur données de santé": {
                "formules": "Moyenne, médiane, étendue  ·  Lecture et construction de graphiques",
                "contexte": "Analyse des données de suivi (tension, glycémie, fréquence cardiaque).",
                "documents": "fiche de surveillance, graphique de suivi"
            },
        }
    },
}

SYSTEM_COINTERVENTION = """\
Tu es un professeur qui anime une séance de CO-INTERVENTION entre mathématiques et matière professionnelle en Bac Pro.
La co-intervention relie explicitement les notions mathématiques aux situations professionnelles réelles du référentiel.

Tes exercices doivent TOUJOURS :
- Partir d'un document professionnel réaliste (facture, bon de commande, bulletin de salaire, planning, étiquette...)
- Utiliser le vocabulaire professionnel exact de la filière
- Montrer explicitement POURQUOI les maths sont utiles dans le métier
- Être ancrés dans des situations que l'élève vivra réellement en stage ou en emploi
- Rappeler la formule mathématique dans son contexte pro (jamais de façon abstraite)

Structure de sortie (Markdown) :

### 📄 Document professionnel
[Reproduire un document réaliste avec données chiffrées : facture, tableau, bulletin, bon de commande...]

### 🎯 Mise en situation
[2-3 lignes de contexte professionnel concret — qui fait quoi, dans quelle entreprise]

### 📐 Rappel de la formule
[Formule + signification de chaque terme dans le vocabulaire professionnel]

### ✏️ Exercices
[3 à 5 questions progressives, du plus guidé au plus autonome]

### ✅ Corrigé détaillé
[Correction complète avec justifications dans le vocabulaire professionnel]\
"""


def build_prompt_cointervention(niveau, filiere, theme, consignes, difficulte="🟡 Moyen"):
    """Construit le prompt pour un exercice de co-intervention."""
    data_filiere = THEMES_COINTERVENTION.get(filiere, {})
    nom_filiere  = data_filiere.get("nom_complet", filiere)
    data_theme   = data_filiere.get("themes", {}).get(theme, {})
    formules     = data_theme.get("formules", "")
    contexte     = data_theme.get("contexte", "")
    documents    = data_theme.get("documents", "")
    diff_label   = difficulte.split(" ", 1)[-1].upper()
    conseils_diff = {
        "DÉBUTANT":  "formules rappelées, calculs en une étape, données déjà extraites du document.",
        "MOYEN":     "quelques étapes guidées, document partiellement exploité.",
        "CONFIRMÉ":  "questions autonomes, document complet à analyser.",
        "EXPERT":    "situation nouvelle, données brutes, raisonnement et conclusion attendus.",
    }.get(diff_label, "")

    user = (
        f"Génère un exercice de CO-INTERVENTION de niveau **{diff_label}** pour :\n"
        f"- Filière : {nom_filiere}\n"
        f"- Niveau scolaire : {niveau} (Bac Pro)\n"
        f"- Thème : {theme}\n"
        f"- Formules clés : {formules}\n"
        f"- Contexte professionnel : {contexte}\n"
        f"- Documents supports : {documents}\n"
        f"- Instructions complémentaires : {consignes or 'Aucune'}\n\n"
        f"Niveau {diff_label} : {conseils_diff}"
    )
    return SYSTEM_COINTERVENTION, user




def build_prompt_correction(bareme, ton, niveau, matiere, note_sur):
    return f"""Tu es un professeur correcteur expert.

Contexte :
- Niveau : {niveau or 'Non précisé'}
- Matière : {matiere or 'Non précisée'}
- Barème : {bareme or f'Non fourni — évalue sur {note_sur}'}
- Ton : {ton}
- Note sur : {note_sur}

Mission :
1. Transcris le texte manuscrit visible.
2. Identifie et explique chaque erreur avec pédagogie.
3. Attribue une note partielle par question.
4. Calcule la note globale /{note_sur}.
5. Rédige une appréciation finale ({ton}) de 3 à 5 lignes.

Réponds en Markdown avec sections claires."""




# ============================================================
# 7. INTERFACE STREAMLIT
# ============================================================

st.set_page_config(
    page_title="Assistant Professeur IA",
    page_icon="🎓",
    layout="wide",
    initial_sidebar_state="expanded"
)

st.markdown("""
<style>
  .stButton>button { border-radius: 8px; font-weight: 600; }
  .info-box  { background:#f0f4ff; border-left:4px solid #4a6cf7; padding:12px 16px; border-radius:4px; margin:8px 0; font-size:.9rem; }
  .warn-box  { background:#fff8e1; border-left:4px solid #f59e0b; padding:12px 16px; border-radius:4px; margin:8px 0; font-size:.9rem; }
  .ok-box    { background:#f0fdf4; border-left:4px solid #22c55e; padding:12px 16px; border-radius:4px; margin:8px 0; font-size:.9rem; }
  .badge     { display:inline-block; background:#e0e7ff; color:#3730a3; border-radius:6px;
               padding:3px 10px; font-size:.8rem; margin-right:6px; margin-bottom:4px; }
</style>
""", unsafe_allow_html=True)

# ── SESSION STATE ─────────────────────────────────────────────
for key in ["generated_md", "generated_ccf_md", "correction_md", "meta_gen", "meta_ccf", "generated_coint_md"]:
    if key not in st.session_state:
        st.session_state[key] = None

# ── SIDEBAR ──────────────────────────────────────────────────
with st.sidebar:
    st.header("⚙️ Configuration")

    st.subheader("🔑 Clé API Gemini")
    cle_api = st.text_input("Clé Google Gemini", type="password", key="gemini_key")
    if cle_api:
        st.markdown('<div class="ok-box">✅ Clé API renseignée</div>', unsafe_allow_html=True)
    else:
        st.markdown('<div class="warn-box">⚠️ Clé API manquante</div>', unsafe_allow_html=True)

    with st.expander("📋 Obtenir une clé gratuite"):
        st.markdown("""
**3 étapes :**

**1.** → [aistudio.google.com](https://aistudio.google.com/app/apikey)

**2.** Connectez-vous → **"Create API Key"**

**3.** Copiez et collez ici

---
💡 **Quota gratuit :** 1 500 req/jour, 15/min.

🔧 **Erreur 429 ?** Attendez 1 min ou activez la facturation sur votre projet Google Cloud (gratuit).
        """)

    st.divider()
    st.subheader("🏫 Établissement")
    nom_etablissement = st.text_input(
        "Nom de l'établissement",
        value="Lino Ventura (Ozoir-la-Ferrière)",
        help="Apparaît dans l'en-tête des CCF officiels"
    )
    annee_scolaire = st.text_input("Année scolaire", value="2025/2026")

    st.divider()
    st.caption("Version 6.0 · Gemini 2.5 Flash")

# ── TITRE ────────────────────────────────────────────────────
st.title("🎓 Assistant Professeur IA")
st.caption("Génération de sujets · CCF Bac Pro · Correction de copies · Export Word")

# ── ONGLETS ──────────────────────────────────────────────────
tab_gen, tab_ccf, tab_coint, tab_correction, tab_export, tab_prop = st.tabs([
    "📝 Générateur de Sujets",
    "🎯 Sujets CCF",
    "🔗 Co-intervention",
    "📸 Correction de Copies",
    "📊 Export Pronote",
    "📥 Propositions élèves",
])


# ─────────────────────────────────────────────────────────────
# ONGLET 1 — GÉNÉRATEUR
# ─────────────────────────────────────────────────────────────
with tab_gen:
    st.subheader("📝 Générateur de sujets et exercices")
    col1, col2 = st.columns(2)

    with col1:
        cat = st.selectbox("Type d'établissement", list(NIVEAUX_CATEGORIES.keys()), key="gen_cat")
        niv = st.selectbox("Classe", NIVEAUX_CATEGORIES[cat], key="gen_niv")
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
        mat = st.selectbox("Matière", MATIERES, key="gen_mat")
        chap = st.selectbox("Chapitre (BO)", get_chapitres(mat, niv, cat, filiere), key="gen_chap")

    consignes = st.text_area("Consignes particulières (optionnel)", height=80, key="gen_consignes",
                              placeholder="Ex : 3 exercices, niveau accessible, inclure un graphique…")

    # ── Niveau de difficulté ──────────────────────────────────
    st.markdown("**🎯 Niveau de difficulté (pédagogie différenciée)**")
    difficulte = st.select_slider(
        "Niveau",
        options=list(NIVEAUX_DIFFICULTE.keys()),
        value="🟡 Moyen",
        key="gen_diff",
        label_visibility="collapsed"
    )
    st.markdown(
        f'<div class="info-box">{difficulte} — {NIVEAUX_DIFFICULTE[difficulte]}</div>',
        unsafe_allow_html=True
    )

    if st.button("✨ Générer le sujet", type="primary", use_container_width=True):
        if not cle_api:
            st.error("🔑 Renseigne ta clé API dans le panneau gauche !")
        else:
            with st.spinner("⏳ Génération en cours…"):
                try:
                    res = call_gemini(cle_api, build_prompt_exercices(niv, cat, mat, chap, consignes, filiere, difficulte))
                    st.session_state.generated_md = res
                    st.session_state.meta_gen = {"niveau": niv, "matiere": mat, "chapitre": chap, "filiere": filiere, "difficulte": difficulte}
                    st.success("✅ Sujet généré !")
                except Exception as e:
                    if "429" in str(e):
                        st.error("⏱️ Quota dépassé (429). Attendez 1 minute et réessayez.")
                    else:
                        st.error(f"Erreur API : {e}")

    if st.session_state.generated_md:
        m = st.session_state.meta_gen or {}
        st.divider()
        badges = (
            f'<span class="badge">📚 {m.get("matiere","")}</span>'
            f'<span class="badge">🎓 {m.get("niveau","")}</span>'
            f'<span class="badge">📖 {m.get("chapitre","")}</span>'
        )
        if m.get("filiere"):
            badges += f'<span class="badge">🏢 {m["filiere"]}</span>'
        if m.get("difficulte"):
            badges += f'<span class="badge">{m["difficulte"]}</span>'
        st.markdown(badges, unsafe_allow_html=True)
        st.markdown(st.session_state.generated_md)

        titre_doc = f"Sujet — {m.get('matiere','')} {m.get('niveau','')} {m.get('difficulte','')}"
        c1, c2, c3 = st.columns(3)
        with c1:
            st.download_button("📥 .md", st.session_state.generated_md,
                               file_name="sujet.md", mime="text/markdown", key="dl1_md")
        with c2:
            st.download_button("📄 .txt", st.session_state.generated_md,
                               file_name="sujet.txt", mime="text/plain", key="dl1_txt")
        with c3:
            st.download_button("📝 .docx",
                markdown_to_docx(st.session_state.generated_md, titre_doc),
                file_name="sujet.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                key="dl1_docx")


# ─────────────────────────────────────────────────────────────
# ONGLET 2 — CCF
# ─────────────────────────────────────────────────────────────
with tab_ccf:
    st.subheader("🎯 Générateur de Sujets CCF — Bac Pro")

    mode = st.radio(
        "Mode de génération",
        ["📋 Entraînement (sans en-tête officiel)",
         "📄 Examen officiel (avec en-tête et barème imprimable)"],
        key="ccf_mode"
    )
    is_officiel = "officiel" in mode

    if is_officiel:
        st.markdown('<div class="info-box">📄 <strong>Mode Examen officiel</strong> — En-tête académique complet, cadre candidat, grille de compétences et fiche d\'évaluation. Prêt à imprimer.</div>', unsafe_allow_html=True)
    else:
        st.markdown('<div class="info-box">📋 <strong>Mode Entraînement</strong> — Structure CCF sans en-tête officiel. Idéal pour entraîner les élèves avant l\'examen.</div>', unsafe_allow_html=True)

    col1, col2 = st.columns(2)
    with col1:
        ccf_cat = st.selectbox("Catégorie", ["Bac Pro", "CAP"], key="ccf_cat")
        # Les CCF n'ont lieu qu'en 1ère et Terminale — 2nde Pro exclue
        niveaux_ccf = [n for n in NIVEAUX_CATEGORIES[ccf_cat] if n != "2nde Pro"]
        ccf_niv = st.selectbox("Classe", niveaux_ccf, key="ccf_niv")
        ccf_fil = st.selectbox("Filière", LISTE_FILIERES, key="ccf_fil")
        ccf_filiere = ""
        if ccf_fil == "Autre (Préciser ci-dessous)":
            ccf_filiere = st.text_input("Filière libre", key="ccf_spec")
        else:
            ccf_filiere = ccf_fil
            if ccf_fil in CONTEXTES_FILIERES:
                st.info(f"📌 {CONTEXTES_FILIERES[ccf_fil]['nom_complet']}")

    with col2:
        ccf_mat = st.selectbox("Matière", MATIERES, key="ccf_mat")
        ccf_duree = st.selectbox("Durée de l'épreuve", ["30 min", "45 min", "1h00", "1h30", "2h00"], index=1, key="ccf_duree")
        num_sit = st.number_input("N° de situation CCF", min_value=1, max_value=3, value=1, step=1, key="ccf_num_sit")

    # ── Chapitres ────────────────────────────────────────────
    st.markdown("**📖 Chapitres évalués**")
    chapitres_dispo = get_chapitres(ccf_mat, ccf_niv, ccf_cat, ccf_filiere)
    col_a, col_b = st.columns(2)
    with col_a:
        ccf_chap = st.selectbox("Chapitre — Partie A", chapitres_dispo, key="ccf_chap")
    with col_b:
        double_chap = st.checkbox("Ajouter un 2ème chapitre (Partie B)", key="ccf_double_chap")
        if double_chap:
            chapitres_b = [c for c in chapitres_dispo if c != ccf_chap]
            ccf_chap_b = st.selectbox("Chapitre — Partie B", chapitres_b, key="ccf_chap_b")
        else:
            ccf_chap_b = ""

    if double_chap and ccf_chap_b:
        st.markdown(
            f'<div class="info-box">📐 <strong>Partie A :</strong> {ccf_chap} &nbsp;|&nbsp; '
            f'<strong>Partie B :</strong> {ccf_chap_b}<br>'
            f'Les deux parties s\'appuieront sur la même situation professionnelle.</div>',
            unsafe_allow_html=True
        )

    avec_corrige = st.checkbox(
        "📝 Inclure le corrigé détaillé *(document professeur)*",
        value=True,
        key="ccf_avec_corrige",
        help="Décochez pour générer uniquement le sujet élève — réponse plus courte, moins de tokens consommés."
    )
    if not avec_corrige:
        st.markdown('<div class="info-box">💡 Sujet seul — génération plus rapide et moins gourmande en quota.</div>', unsafe_allow_html=True)

    ccf_consignes = st.text_area("Consignes spécifiques (optionnel)", height=80, key="ccf_consignes",
                                  placeholder="Ex : Situation en EHPAD, tableau de données, niveau accessible…")

    btn_lbl = "📄 Générer le Sujet CCF Officiel" if is_officiel else "📋 Générer le Sujet d'Entraînement CCF"
    if st.button(btn_lbl, type="primary", use_container_width=True):
        if not cle_api:
            st.error("🔑 Renseigne ta clé API dans le panneau gauche !")
        else:
            with st.spinner("⏳ Génération du sujet CCF…"):
                try:
                    if is_officiel:
                        prompt = build_prompt_ccf_officiel(
                            ccf_niv, ccf_cat, ccf_mat, ccf_chap,
                            ccf_consignes, ccf_filiere, ccf_duree, str(num_sit),
                            avec_corrige=avec_corrige,
                            chapitre_b=ccf_chap_b
                        )
                    else:
                        prompt = build_prompt_ccf_entrainement(
                            ccf_niv, ccf_cat, ccf_mat, ccf_chap,
                            ccf_consignes, ccf_filiere,
                            avec_corrige=avec_corrige,
                            chapitre_b=ccf_chap_b
                        )
                    res = call_gemini(cle_api, prompt)
                    if res:
                        st.session_state.generated_ccf_md = res
                        chap_label = ccf_chap
                        if ccf_chap_b:
                            chap_label += f" + {ccf_chap_b}"
                        st.session_state.meta_ccf = {
                            "niveau": ccf_niv,
                            "matiere": ccf_mat,
                            "chapitre": chap_label,
                            "filiere": ccf_filiere,
                            "mode": "Officiel" if is_officiel else "Entraînement",
                            "num_situation": str(num_sit),
                            "duree": ccf_duree,
                            "annee_scolaire": annee_scolaire,
                        }
                        st.success("✅ Sujet CCF généré !")
                    else:
                        st.error("❌ L'IA n'a pas renvoyé de texte. Réessayez.")
                except Exception as e:
                    if "429" in str(e):
                        st.error("⏱️ Quota dépassé (429). Attendez 1 minute et réessayez.")
                    else:
                        st.error(f"Erreur API : {e}")

    if st.session_state.generated_ccf_md:
        m = st.session_state.meta_ccf or {}
        st.divider()
        badges = (
            f'<span class="badge">🎯 CCF {m.get("mode","")}</span>'
            f'<span class="badge">📚 {m.get("matiere","")}</span>'
            f'<span class="badge">🎓 {m.get("niveau","")}</span>'
            f'<span class="badge">📖 {m.get("chapitre","")}</span>'
        )
        if m.get("filiere"):
            badges += f'<span class="badge">🏢 {m["filiere"]}</span>'
        if m.get("duree"):
            badges += f'<span class="badge">⏱ {m["duree"]}</span>'
        st.markdown(badges, unsafe_allow_html=True)
        st.markdown(st.session_state.generated_ccf_md)

        titre_doc = f"CCF_{m.get('mode','')}_{m.get('matiere','')}_{m.get('niveau','')}"
        st.subheader("📥 Télécharger")

        if is_officiel:
            c1, c2 = st.columns(2)
            with c1:
                docx_off = generate_ccf_officiel_docx(
                    st.session_state.generated_ccf_md, m, nom_etablissement
                )
                if docx_off:
                    st.download_button(
                        "🏛️ Word Officiel (avec en-tête et fiche)",
                        docx_off,
                        file_name=f"{titre_doc}_officiel.docx",
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        type="primary",
                        key="dl_ccf_officiel"
                    )
            with c2:
                st.download_button("📥 .md (brut)", st.session_state.generated_ccf_md,
                                   file_name=f"{titre_doc}.md", mime="text/markdown", key="dl2_md_off")
        else:
            c1, c2, c3 = st.columns(3)
            with c1:
                st.download_button("📝 .docx",
                    markdown_to_docx(st.session_state.generated_ccf_md, titre_doc),
                    file_name=f"{titre_doc}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    key="dl2_docx")
            with c2:
                st.download_button("📥 .md", st.session_state.generated_ccf_md,
                                   file_name=f"{titre_doc}.md", mime="text/markdown", key="dl2_md")
            with c3:
                st.download_button("📄 .txt", st.session_state.generated_ccf_md,
                                   file_name=f"{titre_doc}.txt", mime="text/plain", key="dl2_txt")


# ─────────────────────────────────────────────────────────────
# ONGLET 3 — CO-INTERVENTION
# ─────────────────────────────────────────────────────────────
with tab_coint:
    st.subheader("🔗 Co-intervention — Maths & Matières Professionnelles")
    st.markdown(
        '<div class="info-box">🤝 <strong>Co-intervention</strong> — Exercices qui relient '
        'les mathématiques aux situations professionnelles réelles du référentiel. '
        'Chaque exercice part d\'un document pro concret (facture, bulletin, bon de commande…).</div>',
        unsafe_allow_html=True
    )

    col1, col2 = st.columns(2)
    with col1:
        coint_fil = st.selectbox(
            "Filière",
            list(THEMES_COINTERVENTION.keys()),
            key="coint_fil",
            format_func=lambda k: f"{k} — {THEMES_COINTERVENTION[k]['nom_complet']}"
        )
        coint_niv = st.selectbox(
            "Classe",
            ["2nde Pro", "1ère Pro", "Term Pro"],
            key="coint_niv"
        )
    with col2:
        themes_dispo = list(THEMES_COINTERVENTION[coint_fil]["themes"].keys())
        coint_theme = st.selectbox("Thème de co-intervention", themes_dispo, key="coint_theme")

        # Afficher les infos du thème sélectionné
        data_th = THEMES_COINTERVENTION[coint_fil]["themes"][coint_theme]
        st.markdown(
            f'<div class="info-box" style="font-size:.82rem">'
            f'📐 <strong>Formules :</strong> {data_th["formules"]}<br>'
            f'📄 <strong>Documents :</strong> {data_th["documents"]}</div>',
            unsafe_allow_html=True
        )

    st.markdown("**🎯 Niveau de difficulté**")
    coint_diff = st.select_slider(
        "Niveau co-intervention",
        options=list(NIVEAUX_DIFFICULTE.keys()),
        value="🟡 Moyen",
        key="coint_diff",
        label_visibility="collapsed"
    )
    st.markdown(
        f'<div class="info-box">{coint_diff} — {NIVEAUX_DIFFICULTE[coint_diff]}</div>',
        unsafe_allow_html=True
    )

    coint_consignes = st.text_area(
        "Consignes particulières (optionnel)", height=80, key="coint_consignes",
        placeholder="Ex : Utiliser une facture de fournitures de bureau, insister sur le calcul de TVA à 20%…"
    )

    if st.button("🔗 Générer l'exercice de co-intervention", type="primary", use_container_width=True):
        if not cle_api:
            st.error("🔑 Renseigne ta clé API dans le panneau gauche !")
        else:
            with st.spinner("⏳ Génération de l'exercice co-intervention…"):
                try:
                    sys_instr, user_prompt = build_prompt_cointervention(
                        coint_niv, coint_fil, coint_theme, coint_consignes, coint_diff
                    )
                    res = call_gemini(cle_api, (sys_instr, user_prompt))
                    st.session_state.generated_coint_md = res
                    st.success("✅ Exercice co-intervention généré !")
                except Exception as e:
                    if "429" in str(e):
                        st.error("⏱️ Quota dépassé (429). Attendez 1 minute et réessayez.")
                    else:
                        st.error(f"Erreur API : {e}")

    if st.session_state.generated_coint_md:
        nom_filiere = THEMES_COINTERVENTION[coint_fil]["nom_complet"]
        st.divider()
        badges = (
            f'<span class="badge">🔗 Co-intervention</span>'
            f'<span class="badge">🏢 {coint_fil}</span>'
            f'<span class="badge">🎓 {coint_niv}</span>'
            f'<span class="badge">📐 {coint_theme[:40]}</span>'
            f'<span class="badge">{coint_diff}</span>'
        )
        st.markdown(badges, unsafe_allow_html=True)
        st.markdown(st.session_state.generated_coint_md)

        titre_doc = f"CoIntervention_{coint_fil}_{coint_niv}_{coint_theme[:30]}"
        c1, c2, c3 = st.columns(3)
        with c1:
            st.download_button("📥 .md", st.session_state.generated_coint_md,
                               file_name=f"{titre_doc}.md", mime="text/markdown", key="dl_coint_md")
        with c2:
            st.download_button("📄 .txt", st.session_state.generated_coint_md,
                               file_name=f"{titre_doc}.txt", mime="text/plain", key="dl_coint_txt")
        with c3:
            st.download_button("📝 .docx",
                markdown_to_docx(st.session_state.generated_coint_md,
                                 f"Co-intervention — {coint_fil} · {coint_theme}"),
                file_name=f"{titre_doc}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                key="dl_coint_docx")


# ─────────────────────────────────────────────────────────────
# ONGLET 4 — CORRECTION
# ─────────────────────────────────────────────────────────────
with tab_correction:
    st.subheader("📸 Correction IA de Copies par Photo")

    col1, col2 = st.columns(2)
    with col1:
        corr_cat = st.selectbox("Catégorie", list(NIVEAUX_CATEGORIES.keys()), key="corr_cat")
        corr_niv = st.selectbox("Classe", NIVEAUX_CATEGORIES[corr_cat], key="corr_niv")
    with col2:
        corr_mat = st.selectbox("Matière", MATIERES, key="corr_mat")
        note_sur = st.number_input("Note sur :", min_value=10, max_value=100, value=20, step=5)

    img_file = st.file_uploader("📤 Téléverser la photo de la copie", type=["jpg", "jpeg", "png"])

    col1, col2 = st.columns(2)
    with col1:
        bareme = st.text_area("Barème (optionnel)", height=100,
                               placeholder="Ex : Q1 : 3 pts, Q2 : 5 pts, Q3 : 7 pts, Q4 : 5 pts")
    with col2:
        ton = st.select_slider("Ton", options=["Très bienveillant", "Bienveillant", "Encourageant", "Factuel", "Exigeant"])

    if img_file:
        col_img, col_btn = st.columns([1, 2])
        with col_img:
            st.image(img_file, caption="Copie à corriger", width=250)
        with col_btn:
            if st.button("🔍 Lancer la correction IA", type="primary", use_container_width=True):
                if not cle_api:
                    st.error("🔑 Clé API manquante !")
                else:
                    with st.spinner("🔬 Analyse de la copie…"):
                        try:
                            prompt = build_prompt_correction(bareme, ton, corr_niv, corr_mat, note_sur)
                            res = call_gemini(cle_api, prompt, image=img_file)
                            st.session_state.correction_md = res
                            st.success("✅ Correction terminée !")
                        except Exception as e:
                            if "429" in str(e):
                                st.error("⏱️ Quota dépassé (429). Attendez 1 minute.")
                            else:
                                st.error(f"Erreur : {e}")

    if st.session_state.correction_md:
        st.divider()
        st.subheader("📝 Rapport de Correction")
        st.markdown(st.session_state.correction_md)
        c1, c2 = st.columns(2)
        with c1:
            st.download_button("📥 .md", st.session_state.correction_md,
                               file_name="correction.md", mime="text/markdown", key="dl3_md")
        with c2:
            st.download_button("📝 .docx",
                markdown_to_docx(st.session_state.correction_md, "Rapport de correction"),
                file_name="correction.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                key="dl3_docx")


# ─────────────────────────────────────────────────────────────
# ONGLET 4 — EXPORT PRONOTE
# ─────────────────────────────────────────────────────────────
with tab_export:
    st.subheader("📊 Export Pronote — Module en développement")
    st.markdown('<div class="info-box">🚧 <strong>Ce module est en cours de construction.</strong> La prochaine version permettra d\'exporter les notes vers Pronote via CSV compatible.</div>', unsafe_allow_html=True)
    st.markdown("""
### 📋 Fonctionnalités prévues
- Saisie manuelle des notes après correction IA
- Export CSV au format compatible Pronote (import direct)
- Calcul automatique des moyennes de classe
- Appréciation groupée générée par IA pour chaque élève

### 💡 En attendant
Copiez les notes depuis l'onglet **Correction de Copies** et saisissez-les manuellement dans Pronote.
    """)


# ─────────────────────────────────────────────────────────────
# ONGLET 6 — PROPOSITIONS ÉLÈVES
# ─────────────────────────────────────────────────────────────
with tab_prop:
    st.subheader("📥 Propositions élèves — Sujets à valider")
    st.markdown(
        '<div class="info-box">🎓 Sujets générés par tes élèves via Gemini. '
        'Valide les meilleurs pour les ajouter à la banque.</div>',
        unsafe_allow_html=True
    )

    import requests as _req
    import json as _json

    # ── Chargement des propositions ──────────────────────────
    def charger_propositions(statut_filtre="en attente"):
        try:
            api_key  = st.secrets.get("GRIST_API_KEY", "")
            doc_id   = st.secrets.get("GRIST_DOC_ID", "")
            base_url = st.secrets.get("GRIST_URL", "https://grist.numerique.gouv.fr")
            if not api_key or not doc_id:
                return []
            import urllib.parse
            filtre = urllib.parse.quote(_json.dumps({"statut": [statut_filtre]}))
            url = f"{base_url}/api/docs/{doc_id}/tables/Banque_propositions/records?filter={filtre}"
            headers = {"Authorization": f"Bearer {api_key}"}
            resp = _req.get(url, headers=headers, timeout=8)
            if resp.status_code == 200:
                return [(r["id"], r["fields"]) for r in resp.json().get("records", [])]
        except Exception:
            pass
        return []

    def maj_statut_proposition(record_id, nouveau_statut):
        try:
            api_key  = st.secrets.get("GRIST_API_KEY", "")
            doc_id   = st.secrets.get("GRIST_DOC_ID", "")
            base_url = st.secrets.get("GRIST_URL", "https://grist.numerique.gouv.fr")
            url = f"{base_url}/api/docs/{doc_id}/tables/Banque_propositions/records"
            headers = {"Authorization": f"Bearer {api_key}", "Content-Type": "application/json"}
            payload = {"records": [{"id": record_id, "fields": {"statut": nouveau_statut}}]}
            _req.patch(url, headers=headers, json=payload, timeout=5)
        except Exception:
            pass

    # ── Filtres ──────────────────────────────────────────────
    col_f1, col_f2, col_f3 = st.columns(3)
    with col_f1:
        filtre_statut = st.selectbox(
            "Statut", ["en attente", "validé", "rejeté"], key="prop_statut"
        )
    with col_f2:
        if st.button("🔄 Actualiser", use_container_width=True, key="prop_refresh"):
            st.rerun()
    with col_f3:
        st.markdown("")  # spacer

    propositions = charger_propositions(filtre_statut)

    if not propositions:
        st.markdown(
            f'<div class="info-box">📭 Aucune proposition en statut '
            f'<strong>{filtre_statut}</strong>.</div>',
            unsafe_allow_html=True
        )
    else:
        st.markdown(f"**{len(propositions)} proposition(s) trouvée(s)**")

        for record_id, fields in propositions:
            diff_prop = fields.get("niveau_difficulte") or fields.get("difficulte", "")
            with st.expander(
                f"📄 {fields.get('chapitre','?')[:50]} — "
                f"{fields.get('filiere','?')} {fields.get('niveau','?')} "
                f"{diff_prop :} — "
                f"par {fields.get('code_eleve','?')} le {fields.get('date','?')}"
            ):
                # Badges
                badges = (
                    f'<span class="badge">🎓 {fields.get("niveau","")}</span>'
                    f'<span class="badge">🏢 {fields.get("filiere","")}</span>'
                    f'<span class="badge">📖 {fields.get("chapitre","")[:35]}</span>'
                    f'<span class="badge">{diff_prop :}</span>'
                )
                if fields.get("auto_evaluation"):
                    badges += f'<span class="badge">👤 {fields["auto_evaluation"]}</span>'
                st.markdown(badges, unsafe_allow_html=True)

                # Aperçu du contenu
                contenu = fields.get("contenu", "")
                if contenu:
                    st.markdown(contenu[:2000] + ("…" if len(contenu) > 2000 else ""))

                # Actions
                if filtre_statut == "en attente":
                    col_v, col_r, col_dl = st.columns(3)
                    with col_v:
                        if st.button(
                            "✅ Valider", type="primary",
                            use_container_width=True, key=f"val_{record_id}"
                        ):
                            maj_statut_proposition(record_id, "validé")
                            st.success("✅ Proposition validée !")
                            st.rerun()
                    with col_r:
                        if st.button(
                            "❌ Rejeter",
                            use_container_width=True, key=f"rej_{record_id}"
                        ):
                            maj_statut_proposition(record_id, "rejeté")
                            st.info("Proposition rejetée.")
                            st.rerun()
                    with col_dl:
                        # Télécharger en JSON prêt pour la banque
                        import re as _re
                        def _slug_dl(text):
                            text = text.replace(" ", "_").replace("—", "").replace("/", "_")
                            return _re.sub(r"[^a-zA-Z0-9_\-àâéèêëîïôùûüç]", "", text)[:50]

                        nom_fichier_dl = (
                            f"{_slug_dl(fields.get('niveau',''))}_{_slug_dl(fields.get('filiere',''))}"
                            f"_{_slug_dl(fields.get('chapitre',''))}_{_slug_dl(fields.get('difficulte',''))}.json"
                        )
                        contenu_json = _json.dumps([{
                            "niveau":     fields.get("niveau", ""),
                            "categorie":  "Bac Pro",
                            "filiere":    fields.get("filiere", ""),
                            "matiere":    fields.get("matiere", "Mathématiques"),
                            "chapitre":   fields.get("chapitre", ""),
                            "difficulte": fields.get('difficulte',''),
                            "contenu":    contenu,
                            "source":     f"élève {fields.get('code_eleve','')}",
                            "date_generation": fields.get("date", ""),
                        }], ensure_ascii=False, indent=2)

                        st.download_button(
                            "📥 JSON banque",
                            contenu_json,
                            file_name=nom_fichier_dl,
                            mime="application/json",
                            use_container_width=True,
                            key=f"dl_{record_id}"
                        )
                        st.caption("Télécharge → place dans banque/ → git push")
