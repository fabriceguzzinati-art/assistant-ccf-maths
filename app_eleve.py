#!/usr/bin/env python3
"""
generer_banque.py — Script PROFESSEUR pour pré-générer la banque d'exercices.

Lance avec : streamlit run generer_banque.py
Les exercices sont sauvegardés dans le dossier banque/ du projet.
Commite ensuite le dossier banque/ sur GitHub — l'app élève les chargera.
"""

import google.generativeai as genai
import streamlit as st
import json
import os
import re
import time
from datetime import datetime

# ── Configuration ──────────────────────────────────────────────
BANQUE_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "banque")
os.makedirs(BANQUE_DIR, exist_ok=True)

# ── Données BO ─────────────────────────────────────────────────
CHAPITRES_MATHS_BAC_PRO = {
    "1ère Pro": [
        "Statistique à deux variables — ajustement affine et coefficient de détermination",
        "Probabilités — événements, tableaux croisés, probabilités conditionnelles",
        "Suites numériques — suites arithmétiques",
        "Résolution graphique d'équations et d'inéquations f(x)=g(x)",
        "Fonctions polynômes de degré 2 — racines, signe, forme factorisée",
        "Fonction dérivée — variations, extremums, fonction inverse",
        "Calculs commerciaux et financiers — intérêts simples, coûts",
        "Géométrie dans l'espace — solides usuels et sections par un plan",
        "Vecteurs du plan — coordonnées, opérations, norme",
        "Trigonométrie — cercle trigonométrique, fonctions sinus et cosinus",
        "Algorithmique et programmation Python — listes, fonctions, boucles",
        "Automatismes — calcul, grandeurs, lecture graphique",
    ],
    "Term Pro": [
        "Statistiques à deux variables — ajustements non affines, changements de variable",
        "Probabilités — arbres pondérés, formule des probabilités totales, indépendance",
        "Suites géométriques — terme général, sens de variation, somme",
        "Fonctions polynômes de degré 3 — dérivée, variations, extremums",
        "Fonctions exponentielles de base q et logarithme décimal",
        "Calculs commerciaux et financiers — intérêts composés, amortissements",
        "Vecteurs dans l'espace — coordonnées, norme, colinéarité",
        "Trigonométrie — équations, vecteurs de Fresnel",
        "Algorithmique et programmation Python — approfondissement",
        "Automatismes — probabilités, suites, dérivation, vecteurs",
    ],
}

FILIERES = ["ASSP", "MCVB", "MCVA", "AGORA"]
NIVEAUX_DIFF = ["🟢 Débutant", "🟡 Moyen", "🟠 Confirmé", "🔴 Expert"]

CONTEXTES_FILIERES = {
    "ASSP": {"nom_complet": "Accompagnement, Soins et Services à la Personne",
             "mots_cles": "soins, patients, résidents, personnes âgées, EHPAD, pharmacie"},
    "MCVB": {"nom_complet": "Métiers du Commerce et de la Vente — option B",
              "mots_cles": "vente, client, prospection, chiffre d'affaires, commission"},
    "MCVA": {"nom_complet": "Métiers du Commerce et de la Vente — option A",
              "mots_cles": "magasin, rayon, linéaire, stock, inventaire, merchandising"},
    "AGORA": {"nom_complet": "Assistance à la Gestion des Organisations",
               "mots_cles": "entreprise, comptabilité, salaire, facture, ressources humaines"},
}

SYSTEM_EXERCICES = """\
Tu es un professeur expert en pédagogie différenciée pour lycée professionnel (Bac Pro).
Tes élèves ont un niveau en mathématiques fragile : tes énoncés sont toujours clairs,
bienveillants et ancrés dans des contextes professionnels concrets.

En fonction du niveau de difficulté demandé, adapte PRÉCISÉMENT la structure :

━━ DÉBUTANT ━━
- Questions très guidées, micro-étapes (une opération par question).
- Résultats intermédiaires fournis, vocabulaire ultra-simplifié.
- Rappel de cours détaillé avec exemple résolu pas à pas.
- Exercice d'application : calcul direct, données déjà extraites.
- Mise en situation : contexte simple, une seule inconnue.
- Pas de problème ouvert — remplacer par une question bilan guidée.

━━ MOYEN ━━
- Questions semi-guidées (formule rappelée, première étape donnée).
- Rappel de cours synthétique avec un exemple.
- Exercice d'application : 3 questions progressives.
- Mise en situation professionnelle simple avec tableau de données.
- Problème ouvert court (1 question de synthèse guidée).

━━ CONFIRMÉ ━━
- Questions autonomes, aucune aide dans l'énoncé.
- Rappel de cours en points clés uniquement.
- Exercice d'application : questions progressives avec barème.
- Mise en situation professionnelle réaliste et complète.
- Problème ouvert avec raisonnement attendu.

━━ EXPERT ━━
- Questions ouvertes sans guidage, transfert de compétences vers situation nouvelle.
- Rappel de cours absent ou très succinct (2 lignes max).
- Exercice d'application : données brutes à extraire soi-même.
- Mise en situation complexe avec plusieurs informations à croiser.
- Problème ouvert ambitieux — toujours réaliste pour un élève de Bac Pro.

Dans tous les cas : JAMAIS de calcul hors programme Bac Pro, JAMAIS de piège inutile.
Correction détaillée adaptée au même niveau.

Structure (Markdown) :
1. **Rappel de cours** (adapté au niveau)
2. **Exercice d'application** (adapté au niveau)
3. **Exercice de mise en situation** (contexte professionnel de la filière)
4. **Problème ouvert** (adapté au niveau)
5. **Corrections détaillées**\
"""

SYSTEM_INTERACTIF = """\
Tu es un professeur expert en pédagogie pour lycée professionnel (Bac Pro).
Tu génères des exercices interactifs structurés en JSON.

RÈGLES STRICTES :
- Les questions doivent porter sur des CALCULS NUMÉRIQUES uniquement (résultat = un nombre).
- Chaque question est indépendante et auto-suffisante.
- Les données nécessaires sont incluses dans l'énoncé de chaque question.
- L'indice doit aider sans donner la réponse (rappeler la formule, la méthode).
- Le corrigé doit montrer le calcul complet étape par étape.
- Jamais de calcul hors programme Bac Pro.

FORMAT DE SORTIE — JSON pur, sans balises markdown, sans texte avant ou après :
{
  "contexte": "Description de la mise en situation professionnelle (3-4 lignes max)",
  "rappel_cours": "Formule(s) clé(s) et leur signification (2-3 lignes max)",
  "questions": [
    {
      "numero": 1,
      "enonce": "Énoncé complet et autonome de la question avec toutes les données nécessaires",
      "type": "numerique",
      "reponse": "valeur numérique exacte sous forme de nombre (ex: 27.5)",
      "unite": "€",
      "tolerance": 0.5,
      "indice": "Indice utile sans donner la réponse",
      "corrige": "Calcul complet étape par étape avec résultat final"
    }
  ]
}\
"""


# ── Utilitaires ────────────────────────────────────────────────

def slug(text: str) -> str:
    """Transforme un texte en nom de fichier sûr."""
    text = text.replace(" ", "_").replace("—", "").replace("/", "_")
    text = re.sub(r"[^a-zA-Z0-9_\-àâéèêëîïôùûüç]", "", text)
    return text[:60]


def nom_fichier(niveau, filiere, chapitre, difficulte) -> str:
    diff_label = difficulte.split(" ", 1)[-1]
    return os.path.join(
        BANQUE_DIR,
        f"{slug(niveau)}_{slug(filiere)}_{slug(chapitre)}_{slug(diff_label)}.json"
    )


def nom_fichier_interactif(niveau, filiere, chapitre, difficulte) -> str:
    diff_label = difficulte.split(" ", 1)[-1]
    return os.path.join(
        BANQUE_DIR,
        f"{slug(niveau)}_{slug(filiere)}_{slug(chapitre)}_{slug(diff_label)}_interactif.json"
    )


def charger_fichier(path: str) -> list:
    if os.path.exists(path):
        try:
            with open(path, encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            return []
    return []


def sauvegarder_sujet(path: str, sujet: dict) -> None:
    sujets = charger_fichier(path)
    sujets.append(sujet)
    with open(path, "w", encoding="utf-8") as f:
        json.dump(sujets, f, ensure_ascii=False, indent=2)


def generer_un_sujet(cle_api, niveau, filiere, chapitre, difficulte) -> str:
    """Appelle Gemini et retourne le markdown du sujet."""
    ctx_fil = CONTEXTES_FILIERES.get(filiere, {})
    nom_fil = ctx_fil.get("nom_complet", filiere)
    mots    = ctx_fil.get("mots_cles", "")
    diff_label = difficulte.split(" ", 1)[-1].upper()

    user_prompt = (
        f"Génère un contenu pédagogique de niveau **{diff_label}** pour :\n"
        f"- Niveau scolaire : {niveau} (Bac Pro)\n"
        f"- Matière : Mathématiques\n"
        f"- Chapitre : {chapitre}\n"
        f"- Filière : {nom_fil} (univers : {mots})\n\n"
        f"Applique scrupuleusement les consignes du niveau {diff_label}."
    )

    genai.configure(api_key=cle_api.strip())
    model = genai.GenerativeModel(
        model_name="gemini-2.5-flash",
        system_instruction=SYSTEM_EXERCICES
    )
    response = model.generate_content(user_prompt)
    if response and response.text:
        return response.text
    raise ValueError("Gemini n'a pas renvoyé de texte.")


def generer_un_sujet_interactif(cle_api, niveau, filiere, chapitre, difficulte, nb_questions=4) -> dict:
    """
    Appelle Gemini et retourne un dict JSON structuré pour le mode interactif.
    Le JSON contient : contexte, rappel_cours, questions (avec réponses et corrigés).
    """
    ctx_fil    = CONTEXTES_FILIERES.get(filiere, {})
    nom_fil    = ctx_fil.get("nom_complet", filiere)
    mots       = ctx_fil.get("mots_cles", "")
    diff_label = difficulte.split(" ", 1)[-1].upper()

    nb_map = {"DÉBUTANT": 3, "MOYEN": 4, "CONFIRMÉ": 5, "EXPERT": 5}
    nb_q   = nb_map.get(diff_label, nb_questions)

    user_prompt = (
        f"Génère un exercice interactif de niveau {diff_label} avec exactement {nb_q} questions pour :\n"
        f"- Niveau scolaire : {niveau} (Bac Pro)\n"
        f"- Matière : Mathématiques\n"
        f"- Chapitre : {chapitre}\n"
        f"- Filière : {nom_fil} (univers : {mots})\n\n"
        f"IMPORTANT :\n"
        f"- Toutes les questions doivent avoir une réponse NUMÉRIQUE (un seul nombre).\n"
        f"- Les données de chaque question doivent être dans son énoncé.\n"
        f"- Retourne UNIQUEMENT le JSON, sans texte avant ou après, sans balises markdown.\n"
        f"- Le champ 'reponse' doit être un nombre décimal (ex: 27.5 et non '27,5 €').\n"
        f"- Le champ 'unite' contient l'unité séparément (€, %, kg, etc.).\n"
        f"- La 'tolerance' est la marge d'erreur acceptée (0.5 pour €, 0.01 pour %)."
    )

    genai.configure(api_key=cle_api.strip())
    model = genai.GenerativeModel(
        model_name="gemini-2.5-flash",
        system_instruction=SYSTEM_INTERACTIF
    )
    response = model.generate_content(user_prompt)
    if not response or not response.text:
        raise ValueError("Gemini n'a pas renvoyé de texte.")

    # Nettoyer et parser le JSON
    texte = response.text.strip()
    texte = re.sub(r'^```[a-zA-Z]*\n?', '', texte)
    texte = re.sub(r'```$', '', texte).strip()

    try:
        data = json.loads(texte)
    except json.JSONDecodeError as e:
        raise ValueError(f"JSON invalide : {e}\n---\n{texte[:300]}")

    # Vérifications minimales
    if "questions" not in data or not isinstance(data["questions"], list):
        raise ValueError("Le JSON ne contient pas de liste 'questions'.")
    if len(data["questions"]) == 0:
        raise ValueError("La liste 'questions' est vide.")

    # Normaliser les réponses (s'assurer que ce sont des floats)
    for q in data["questions"]:
        try:
            q["reponse"] = float(str(q.get("reponse", "0")).replace(",", ".").replace(" ", ""))
        except (ValueError, TypeError):
            q["reponse"] = 0.0
        if "tolerance" not in q:
            q["tolerance"] = 0.5
        if "indice" not in q:
            q["indice"] = ""
        if "unite" not in q:
            q["unite"] = ""

    return data


# ── Interface Streamlit ────────────────────────────────────────

st.set_page_config(
    page_title="Génération Banque d'Exercices",
    page_icon="🏦",
    layout="wide"
)

st.title("🏦 Génération de la Banque d'Exercices")
st.caption("Outil professeur — les exercices générés sont sauvegardés localement puis commités sur GitHub.")

st.markdown("""
<style>
  .ok-box   { background:#f0fdf4; border-left:4px solid #22c55e; padding:10px 14px; border-radius:4px; margin:6px 0; font-size:.9rem; }
  .warn-box { background:#fff8e1; border-left:4px solid #f59e0b; padding:10px 14px; border-radius:4px; margin:6px 0; font-size:.9rem; }
  .info-box { background:#f0f4ff; border-left:4px solid #4a6cf7; padding:10px 14px; border-radius:4px; margin:6px 0; font-size:.9rem; }
</style>
""", unsafe_allow_html=True)

# ── Sidebar ────────────────────────────────────────────────────
with st.sidebar:
    st.header("⚙️ Configuration")
    cle_api = st.text_input("Clé API Gemini", type="password", key="cle_api")
    if cle_api:
        st.markdown('<div class="ok-box">✅ Clé renseignée</div>', unsafe_allow_html=True)
    else:
        st.markdown('<div class="warn-box">⚠️ Clé API manquante</div>', unsafe_allow_html=True)

    st.divider()
    st.markdown(f"**📁 Dossier banque :**\n`{BANQUE_DIR}`")

    # Comptage des sujets existants
    nb_fichiers = len([f for f in os.listdir(BANQUE_DIR) if f.endswith(".json")])
    nb_sujets   = sum(
        len(charger_fichier(os.path.join(BANQUE_DIR, f)))
        for f in os.listdir(BANQUE_DIR) if f.endswith(".json")
    )
    st.markdown(f"**📊 Banque actuelle :** {nb_sujets} sujet(s) dans {nb_fichiers} fichier(s)")

    st.divider()
    st.markdown("""
**Après génération :**
```bash
git add banque/
git commit -m "Ajout sujets banque"
git push
```
""")

# ── Onglets ────────────────────────────────────────────────────
tab_gen, tab_inventaire, tab_interactif = st.tabs(["➕ Générer des sujets", "📋 Inventaire de la banque", "⚡ Générer mode Interactif"])

# ─────────────────────────────────────────────────────────
# ONGLET 1 — Génération
# ─────────────────────────────────────────────────────────
with tab_gen:
    st.subheader("Paramètres de génération")

    col1, col2 = st.columns(2)
    with col1:
        niveau   = st.selectbox("Classe", ["1ère Pro", "Term Pro"], key="b_niv")
        filiere  = st.selectbox("Filière", FILIERES, key="b_fil")
    with col2:
        chapitres_dispo = CHAPITRES_MATHS_BAC_PRO[niveau]
        chapitres_sel   = st.multiselect(
            "Chapitres à générer",
            chapitres_dispo,
            default=[chapitres_dispo[0]],
            key="b_chap"
        )
        niveaux_sel = st.multiselect(
            "Niveaux de difficulté",
            NIVEAUX_DIFF,
            default=["🟢 Débutant", "🟡 Moyen"],
            key="b_diff"
        )

    nb_par_combo = st.number_input(
        "Nombre de sujets à générer par combinaison",
        min_value=1, max_value=5, value=2, step=1,
        help="2 sujets par combinaison = bonne base pour l'aléatoire."
    )

    # Calcul du total et estimation du coût
    total = len(chapitres_sel) * len(niveaux_sel) * nb_par_combo
    if chapitres_sel and niveaux_sel:
        st.markdown(
            f'<div class="info-box">📊 <strong>{total} sujet(s)</strong> à générer '
            f'({len(chapitres_sel)} chapitre(s) × {len(niveaux_sel)} niveau(x) × {nb_par_combo} sujet(s))<br>'
            f'⏱ Durée estimée : ~{total * 15} secondes (pause anti-429 entre chaque appel)</div>',
            unsafe_allow_html=True
        )

    if st.button("🚀 Lancer la génération", type="primary",
                 use_container_width=True, disabled=not cle_api):
        if not chapitres_sel:
            st.error("Sélectionne au moins un chapitre.")
        elif not niveaux_sel:
            st.error("Sélectionne au moins un niveau de difficulté.")
        else:
            progress_bar = st.progress(0)
            status       = st.empty()
            log_area     = st.empty()
            logs         = []
            total_done   = 0

            for chapitre in chapitres_sel:
                for diff in niveaux_sel:
                    for i in range(nb_par_combo):
                        label = f"{chapitre[:40]}… | {diff} | #{i+1}"
                        status.markdown(f"⏳ Génération : **{label}**")
                        try:
                            contenu = generer_un_sujet(cle_api, niveau, filiere, chapitre, diff)
                            path = nom_fichier(niveau, filiere, chapitre, diff)
                            sauvegarder_sujet(path, {
                                "niveau":      niveau,
                                "categorie":   "Bac Pro",
                                "filiere":     filiere,
                                "matiere":     "Mathématiques",
                                "chapitre":    chapitre,
                                "difficulte":  diff,
                                "contenu":     contenu,
                                "date_generation": datetime.now().strftime("%Y-%m-%d"),
                            })
                            logs.append(f"✅ {label}")
                        except Exception as e:
                            err = str(e)
                            if "429" in err:
                                logs.append(f"⏳ 429 sur {label} — pause 60s…")
                                log_area.text("\n".join(logs))
                                time.sleep(60)
                                # Réessai
                                try:
                                    contenu = generer_un_sujet(cle_api, niveau, filiere, chapitre, diff)
                                    path = nom_fichier(niveau, filiere, chapitre, diff)
                                    sauvegarder_sujet(path, {
                                        "niveau": niveau, "categorie": "Bac Pro",
                                        "filiere": filiere, "matiere": "Mathématiques",
                                        "chapitre": chapitre, "difficulte": diff,
                                        "contenu": contenu,
                                        "date_generation": datetime.now().strftime("%Y-%m-%d"),
                                    })
                                    logs.append(f"✅ {label} (2ème tentative)")
                                except Exception as e2:
                                    logs.append(f"❌ Échec définitif : {label} — {e2}")
                            else:
                                logs.append(f"❌ {label} — {e}")

                        total_done += 1
                        progress_bar.progress(total_done / total)
                        log_area.text("\n".join(logs))

                        # Pause anti-429 entre chaque appel (sauf dernier)
                        if total_done < total:
                            time.sleep(8)

            status.markdown("✅ **Génération terminée !**")
            st.success(f"{total_done} sujet(s) traités. N'oublie pas de commiter le dossier `banque/` sur GitHub !")

# ─────────────────────────────────────────────────────────
# ONGLET 2 — Inventaire
# ─────────────────────────────────────────────────────────
with tab_inventaire:
    st.subheader("📋 Inventaire de la banque")

    fichiers = sorted([f for f in os.listdir(BANQUE_DIR) if f.endswith(".json")])
    if not fichiers:
        st.info("La banque est vide. Génère des sujets dans l'onglet précédent.")
    else:
        data = []
        for f in fichiers:
            sujets = charger_fichier(os.path.join(BANQUE_DIR, f))
            if sujets and isinstance(sujets, list) and len(sujets) > 0:
                s = sujets[0]
                data.append({
                    "Niveau":      s.get("niveau", ""),
                    "Filière":     s.get("filiere", ""),
                    "Chapitre":    s.get("chapitre", "")[:50] + "…" if len(s.get("chapitre", "")) > 50 else s.get("chapitre", ""),
                    "Difficulté":  s.get("difficulte", ""),
                    "Nb sujets":   len(sujets),
                    "Dernière génération": sujets[-1].get("date_generation", "—"),
                })

        import pandas as pd
        df = pd.DataFrame(data)
        st.dataframe(df, use_container_width=True, hide_index=True)
        st.markdown(f"**Total : {df['Nb sujets'].sum()} sujet(s) dans {len(df)} combinaison(s)**")

        # Option suppression
        with st.expander("🗑️ Supprimer des fichiers"):
            fichier_del = st.selectbox("Fichier à supprimer", fichiers, key="del_file")
            if st.button("Supprimer", type="secondary"):
                os.remove(os.path.join(BANQUE_DIR, fichier_del))
                st.success(f"Fichier {fichier_del} supprimé.")
                st.rerun()

# ─────────────────────────────────────────────────────────
# ONGLET 3 — Mode Interactif
# ─────────────────────────────────────────────────────────
with tab_interactif:
    st.subheader("⚡ Générer des exercices en mode Interactif")
    st.markdown(
        '<div class="info-box">🎮 <strong>Mode Interactif</strong> — '
        'L\'élève répond question par question, voit son score en fin d\'exercice. '
        'Questions à réponses numériques uniquement. '
        'Sauvegardé avec le suffixe <code>_interactif.json</code> — '
        'coexiste avec les sujets classiques.</div>',
        unsafe_allow_html=True
    )

    col1, col2 = st.columns(2)
    with col1:
        niveau_i  = st.selectbox("Classe", ["1ère Pro", "Term Pro"], key="i_niv")
        filiere_i = st.selectbox("Filière", FILIERES, key="i_fil")
    with col2:
        chapitres_i = CHAPITRES_MATHS_BAC_PRO[niveau_i]
        chapitres_sel_i = st.multiselect(
            "Chapitres à générer",
            chapitres_i,
            default=[chapitres_i[0]],
            key="i_chap"
        )
        niveaux_sel_i = st.multiselect(
            "Niveaux de difficulté",
            NIVEAUX_DIFF,
            default=["🟢 Débutant", "🟡 Moyen"],
            key="i_diff"
        )

    nb_par_combo_i = st.number_input(
        "Nombre de sujets interactifs par combinaison",
        min_value=1, max_value=3, value=1, step=1,
        key="i_nb",
        help="1 ou 2 suffisent — la génération JSON est plus longue."
    )

    total_i = len(chapitres_sel_i) * len(niveaux_sel_i) * nb_par_combo_i
    if chapitres_sel_i and niveaux_sel_i:
        st.markdown(
            f'<div class="info-box">📊 <strong>{total_i} sujet(s) interactif(s)</strong> à générer<br>'
            f'⏱ Durée estimée : ~{total_i * 20} secondes</div>',
            unsafe_allow_html=True
        )

    if st.button("⚡ Lancer la génération interactive", type="primary",
                 use_container_width=True, disabled=not cle_api):
        if not chapitres_sel_i:
            st.error("Sélectionne au moins un chapitre.")
        elif not niveaux_sel_i:
            st.error("Sélectionne au moins un niveau.")
        else:
            progress_bar = st.progress(0)
            status       = st.empty()
            log_area     = st.empty()
            logs         = []
            total_done   = 0

            for chapitre in chapitres_sel_i:
                for diff in niveaux_sel_i:
                    for i in range(nb_par_combo_i):
                        label = f"{chapitre[:35]}… | {diff} | #{i+1}"
                        status.markdown(f"⚡ Génération interactive : **{label}**")
                        try:
                            data = generer_un_sujet_interactif(
                                cle_api, niveau_i, filiere_i, chapitre, diff
                            )
                            path = nom_fichier_interactif(niveau_i, filiere_i, chapitre, diff)
                            sujets_exist = charger_fichier(path)
                            sujets_exist.append({
                                "niveau":           niveau_i,
                                "filiere":          filiere_i,
                                "chapitre":         chapitre,
                                "difficulte":       diff,
                                "format":           "interactif",
                                "date_generation":  datetime.now().strftime("%Y-%m-%d"),
                                **data,
                            })
                            with open(path, "w", encoding="utf-8") as f:
                                json.dump(sujets_exist, f, ensure_ascii=False, indent=2)

                            nb_q = len(data.get("questions", []))
                            logs.append(f"✅ {label} — {nb_q} questions")

                        except Exception as e:
                            err = str(e)
                            if "429" in err:
                                logs.append(f"⏳ 429 sur {label} — pause 60s…")
                                log_area.text("\n".join(logs))
                                time.sleep(60)
                                try:
                                    data = generer_un_sujet_interactif(
                                        cle_api, niveau_i, filiere_i, chapitre, diff
                                    )
                                    path = nom_fichier_interactif(niveau_i, filiere_i, chapitre, diff)
                                    sujets_exist = charger_fichier(path)
                                    sujets_exist.append({
                                        "niveau": niveau_i, "filiere": filiere_i,
                                        "chapitre": chapitre, "difficulte": diff,
                                        "format": "interactif",
                                        "date_generation": datetime.now().strftime("%Y-%m-%d"),
                                        **data,
                                    })
                                    with open(path, "w", encoding="utf-8") as f:
                                        json.dump(sujets_exist, f, ensure_ascii=False, indent=2)
                                    logs.append(f"✅ {label} (2ème tentative)")
                                except Exception as e2:
                                    logs.append(f"❌ Échec : {label} — {e2}")
                            else:
                                logs.append(f"❌ {label} — {err[:80]}")

                        total_done += 1
                        progress_bar.progress(total_done / total_i)
                        log_area.text("\n".join(logs))
                        if total_done < total_i:
                            time.sleep(10)

            status.markdown("✅ **Génération interactive terminée !**")
            st.success(f"{total_done} sujet(s) traités. Commite le dossier `banque/` sur GitHub !")

    # Aperçu des fichiers interactifs existants
    st.divider()
    st.markdown("**📂 Fichiers interactifs existants**")
    fichiers_i = sorted([
        f for f in os.listdir(BANQUE_DIR)
        if f.endswith("_interactif.json")
    ])
    if not fichiers_i:
        st.info("Aucun fichier interactif encore généré.")
    else:
        for f in fichiers_i:
            sujets = charger_fichier(os.path.join(BANQUE_DIR, f))
            nb_q_total = sum(len(s.get("questions", [])) for s in sujets)
            st.markdown(
                f'<span style="font-size:.85rem">📄 <code>{f}</code> — '
                f'{len(sujets)} sujet(s), {nb_q_total} questions au total</span>',
                unsafe_allow_html=True
            )
