import google.generativeai as genai
import PIL.Image
import streamlit as st
from io import BytesIO
from docx import Document as DocxDocument
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import re
import os

# Chemin absolu du dossier contenant app.py — utilisé pour trouver les images
APP_DIR = os.path.dirname(os.path.abspath(__file__))

# ============================================================
# 1. DONNÉES OFFICIELLES — BO MATHS BAC PRO
# ============================================================

NIVEAUX_CATEGORIES = {
    "Collège": ["6ème", "5ème", "4ème", "3ème"],
    "Lycée Général": ["2nde", "1ère", "Terminale"],
    "Bac Pro": ["2nde Pro", "1ère Pro", "Term Pro"],
    "CAP": ["1ère année CAP", "2ème année CAP"],
    "BTS": ["BTS 1", "BTS 2"]
}

MATIERES = [
    "Mathématiques", "Français", "Histoire-Géographie", "SVT",
    "Physique-Chimie", "Anglais", "Espagnol", "SES",
    "Philosophie", "Sciences de l'ingénieur", "EMC"
]

CHAPITRES_MATHS_BAC_PRO = {
    # ── 2nde Pro ── (programme déjà conforme au BO)
    "2nde Pro": [
        "Calcul numérique et algébrique",
        "Puissances et notations scientifiques",
        "Proportionnalité et pourcentages",
        "Équations du 1er degré à une inconnue",
        "Inéquations du 1er degré",
        "Géométrie plane — figures usuelles",
        "Trigonométrie dans le triangle rectangle",
        "Vecteurs — notions de base",
        "Notion de fonction — représentation graphique",
        "Fonctions linéaires et affines",
        "Statistiques descriptives — série à une variable",
        "Notion de probabilité — expériences aléatoires",
    ],

    # ── 1ère Pro ── conforme au BO Annexe 1 (Mathématiques — Classe de première professionnelle)
    "1ère Pro": [
        # Statistique et probabilités
        "Statistique à deux variables — ajustement affine et coefficient de détermination",
        "Probabilités — événements, tableaux croisés, probabilités conditionnelles",
        # Algèbre – Analyse
        "Suites numériques — suites arithmétiques",
        "Résolution graphique d'équations et d'inéquations f(x)=g(x)",
        "Fonctions polynômes de degré 2 — racines, signe, forme factorisée",
        "Fonction dérivée — variations, extremums, fonction inverse",
        "Calculs commerciaux et financiers — intérêts simples, coûts (filières sans physique-chimie)",
        # Géométrie
        "Géométrie dans l'espace — solides usuels et sections par un plan",
        "Vecteurs du plan — coordonnées, opérations, norme (groupements A et B)",
        "Trigonométrie — cercle trigonométrique, fonctions sinus et cosinus (groupements A et B)",
        # Modules transversaux
        "Algorithmique et programmation Python — listes, fonctions, boucles",
        "Automatismes — calcul, grandeurs, lecture graphique",
    ],

    # ── Term Pro ── conforme au BO Annexe 2 (Mathématiques — Classe terminale professionnelle)
    "Term Pro": [
        # Statistique et probabilités
        "Statistiques à deux variables — ajustements non affines, changements de variable",
        "Probabilités — arbres pondérés, formule des probabilités totales, indépendance",
        # Algèbre – Analyse
        "Suites géométriques — terme général, sens de variation, somme",
        "Fonctions polynômes de degré 3 — dérivée, variations, extremums",
        "Fonctions exponentielles de base q et logarithme décimal",
        "Calculs commerciaux et financiers — intérêts composés, amortissements (filières sans physique-chimie)",
        # Géométrie
        "Vecteurs dans l'espace — coordonnées, norme, colinéarité (groupement B)",
        "Trigonométrie — équations, vecteurs de Fresnel (groupement A)",
        # Modules transversaux
        "Algorithmique et programmation Python — approfondissement listes et fonctions",
        "Automatismes — probabilités, suites, dérivation, vecteurs",
        # Programme complémentaire (poursuite d'études)
        "Calcul intégral — primitives, intégrale, aire (programme complémentaire)",
        "Fonctions logarithme népérien et exponentielle de base e (programme complémentaire)",
        "Nombres complexes — forme algébrique et trigonométrique (programme complémentaire)",
        "Produit scalaire de deux vecteurs du plan (programme complémentaire)",
    ],
}

CHAPITRES_PAR_MATIERE_GENERAL = {
    "Mathématiques": ["Nombres et opérations", "Fractions", "Équations", "Fonctions", "Géométrie", "Statistiques", "Probabilités", "Algorithmique"],
    "Français": ["Grammaire", "Orthographe", "Conjugaison", "Analyse littéraire", "Argumentation", "Expression écrite", "Oral"],
    "Histoire-Géographie": ["Antiquité", "Moyen Âge", "Époque Moderne", "Époque Contemporaine", "Géographie de la France", "Géographie mondiale", "EMC"],
    "SVT": ["Cellule et génétique", "Évolution", "Corps humain et santé", "Écosystèmes", "Géologie"],
    "Physique-Chimie": ["Mécanique", "Électricité", "Optique", "Thermodynamique", "Chimie organique", "Chimie des solutions"],
    "Anglais": ["Compréhension écrite", "Expression écrite", "Compréhension orale", "Expression orale", "Grammaire", "Vocabulaire professionnel"],
    "Espagnol": ["Compréhension écrite", "Expression écrite", "Grammaire", "Civilisation", "Vocabulaire"],
    "SES": ["Économie", "Sociologie", "Science politique", "Mondialisation"],
    "Philosophie": ["La conscience", "Le langage", "La liberté", "La morale", "La politique", "La vérité", "L'art"],
    "Sciences de l'ingénieur": ["Mécanique", "Électronique", "Informatique industrielle", "Matériaux"],
    "EMC": ["Démocratie et citoyenneté", "Droits et libertés", "Laïcité", "Engagement"],
}

# ============================================================
# 2. FILIÈRES PRO
# ============================================================

LISTE_FILIERES = ["ASSP", "MCVB", "MCVA", "AGORA", "Autre (Préciser ci-dessous)"]

CONTEXTES_FILIERES = {
    "ASSP": {
        "nom_complet": "Accompagnement, Soins et Services à la Personne",
        "contextes_maths": [
            "Calcul de doses et de dilutions médicamenteuses",
            "Gestion de plannings d'intervenants à domicile",
            "Statistiques sur données de santé (IMC, fréquence cardiaque)",
            "Calcul de coûts de prise en charge",
            "Lecture et interprétation de graphiques médicaux",
        ],
        "mots_cles": "soins, patients, résidents, personnes âgées, handicap, domicile, EHPAD, pharmacie, hygiène"
    },
    "MCVB": {
        "nom_complet": "Métiers du Commerce et de la Vente — option B (Prospection Clientèle)",
        "contextes_maths": [
            "Calcul de marges, taux de remise et prix de vente",
            "Statistiques de vente et de prospection",
            "Gestion d'un budget commercial",
            "Pourcentages d'évolution du chiffre d'affaires",
            "Représentations graphiques des ventes",
        ],
        "mots_cles": "vente, client, prospection, chiffre d'affaires, commission, remise, catalogue"
    },
    "MCVA": {
        "nom_complet": "Métiers du Commerce et de la Vente — option A (Animation et Gestion de l'Espace Commercial)",
        "contextes_maths": [
            "Calcul de surfaces de vente et d'implantation linéaire",
            "Statistiques de fréquentation et taux de transformation",
            "Gestion des stocks — taux de rotation, coefficient multiplicateur",
            "Calcul de TVA, de prix TTC et HT",
            "Optimisation de l'espace commercial",
        ],
        "mots_cles": "magasin, rayon, linéaire, stock, inventaire, promotion, merchandising, caisse"
    },
    "AGORA": {
        "nom_complet": "Assistance à la Gestion des Organisations et de leurs Activités",
        "contextes_maths": [
            "Calcul de salaires, cotisations et charges sociales",
            "Gestion de budgets d'entreprise",
            "Statistiques sur données RH",
            "Facturation et suivi comptable de base",
            "Calcul d'intérêts pour emprunts professionnels",
        ],
        "mots_cles": "entreprise, comptabilité, salaire, facture, devis, ressources humaines, secrétariat"
    },
}

# Compétences officielles du BO avec leurs indicateurs détaillés
COMPETENCES_CCF = [
    {
        "nom": "S'approprier",
        "indicateurs": [
            "Rechercher, extraire et organiser l'information.",
            "Traduire des informations, des codages.",
        ]
    },
    {
        "nom": "Analyser / Raisonner",
        "indicateurs": [
            "Émettre des conjectures, formuler des hypothèses.",
            "Proposer, choisir une méthode de résolution ou un protocole expérimental.",
            "Élaborer un algorithme.",
        ]
    },
    {
        "nom": "Réaliser",
        "indicateurs": [
            "Mettre en œuvre une méthode de résolution, des algorithmes ou un protocole expérimental en respectant les règles de sécurité.",
            "Utiliser un modèle, représenter, calculer.",
            "Expérimenter, faire une simulation.",
        ]
    },
    {
        "nom": "Valider",
        "indicateurs": [
            "Exploiter et interpréter des résultats ou des observations de façon critique et argumentée.",
            "Contrôler la vraisemblance d'une conjecture, de la valeur d'une mesure.",
            "Valider un modèle ou une hypothèse.",
            "Mener un raisonnement logique et établir une conclusion.",
        ]
    },
    {
        "nom": "Communiquer",
        "indicateurs": [
            "Rendre compte d'un résultat, à l'oral ou à l'écrit en utilisant des outils et un langage approprié.",
            "Expliquer une démarche.",
        ]
    },
]

# ============================================================
# 3. PROMPTS
# ============================================================

def get_chapitres(matiere, niveau, categorie):
    if matiere == "Mathématiques" and categorie == "Bac Pro" and niveau in CHAPITRES_MATHS_BAC_PRO:
        return CHAPITRES_MATHS_BAC_PRO[niveau]
    return CHAPITRES_PAR_MATIERE_GENERAL.get(matiere, ["Chapitre général"])


def build_contexte_filiere(filiere):
    if filiere in CONTEXTES_FILIERES:
        ctx = CONTEXTES_FILIERES[filiere]
        lignes = "\n".join(f"- {c}" for c in ctx["contextes_maths"])
        return f"\n**Filière : {ctx['nom_complet']}**\nUnivers : {ctx['mots_cles']}\nContextes maths :\n{lignes}\n"
    return f"\nFilière : {filiere}\n" if filiere else ""


def build_prompt_exercices(niveau, categorie, matiere, chapitre, consignes, filiere=""):
    ctx = build_contexte_filiere(filiere)
    return f"""Tu es un professeur expert en pédagogie différenciée.

Génère un contenu pédagogique complet pour :
- Niveau : {niveau} ({categorie})
- Matière : {matiere}
- Chapitre : {chapitre}
{ctx}
- Instructions : {consignes or 'Aucune'}

STRUCTURE :
1. **Rappel de cours** — points clés (5 à 8 lignes)
2. **Exercice d'application directe** — 3 questions progressives avec barème
3. **Exercice de mise en situation** — contexte professionnel lié à la filière
4. **Problème ouvert** — question de synthèse

Réponds en Markdown avec titres clairs. Inclus les corrections détaillées à la fin."""


def build_prompt_ccf_entrainement(niveau, categorie, matiere, chapitre, consignes, filiere=""):
    ctx = build_contexte_filiere(filiere)
    return f"""Tu es un professeur de mathématiques expert en Bac Pro et en évaluation CCF conforme au BO.

Génère un SUJET D'ENTRAÎNEMENT AU CCF pour :
- Niveau : {niveau} ({categorie})
- Matière : {matiere}
- Chapitre du BO : {chapitre}
{ctx}
- Instructions : {consignes or 'Aucune'}

## RÈGLES STRICTES DU FORMAT CCF OFFICIEL :

1. La MISE EN SITUATION présente le contexte professionnel avec données chiffrées réalistes.
2. La PROBLÉMATIQUE est UNE SEULE question (avec point d'interrogation ?) encadrée en gras, placée juste après la mise en situation. C'est la seule question du document avec un "?".
3. Les questions sont numérotées et chacune porte le nom de la compétence évaluée en majuscules : S'APPROPRIER / RÉALISER / ANALYSER-RAISONNER / VALIDER / COMMUNIQUER.
4. Chaque question se termine par des lignes de réponse (______).
5. Les questions servent à répondre progressivement à la problématique.
6. La dernière question COMMUNIQUER demande de répondre à la problématique initiale.
7. L'évaluation est sur 10 points (pas 20), répartis en niveaux 0/1/2 par compétence.

## STRUCTURE À RESPECTER :

### MISE EN SITUATION PROFESSIONNELLE
[Description du contexte, de l'entreprise/structure, avec données chiffrées et document support : tableau ou graphique]

### PROBLÉMATIQUE
**[UNE SEULE question centrale se terminant par ?]**

### PARTIE A — [Titre lié au premier thème mathématique]

1. **S'APPROPRIER**
- [question]
______

2. **RÉALISER**
- [question]
______

3. **ANALYSER / RAISONNER**
- [question]
______

4. **RÉALISER** 🛎️ APPELER L'EXAMINATEUR
- [question nécessitant outil/calculatrice]
______

5. **VALIDER**
- [question de vérification/interprétation]
______

6. **COMMUNIQUER**
- Répondre à la [première partie de la] problématique.
______

### PARTIE B — [Titre lié au second thème mathématique]
[Mêmes règles, questions numérotées avec compétences]

### CORRIGÉ DÉTAILLÉ *(document professeur)*
Correction complète de chaque question.

Réponds entièrement en Markdown."""


def build_prompt_ccf_officiel(niveau, categorie, matiere, chapitre, consignes, filiere="", duree="45 min", num_sit="1"):
    ctx = build_contexte_filiere(filiere)
    nom_filiere = CONTEXTES_FILIERES[filiere]["nom_complet"] if filiere in CONTEXTES_FILIERES else filiere
    return f"""Tu es un professeur de mathématiques expert en Bac Pro et en évaluation CCF conforme au BO de l'Éducation Nationale.

Génère un SUJET DE CCF OFFICIEL complet et prêt à imprimer pour :
- Niveau : {niveau} ({categorie})
- Filière : {nom_filiere}
- Matière : {matiere}
- Chapitre du BO : {chapitre}
- Situation d'évaluation n° : {num_sit}
- Durée : {duree}
{ctx}
- Instructions complémentaires : {consignes or 'Aucune'}

## RÈGLES STRICTES DU FORMAT CCF OFFICIEL ÉDUCATION NATIONALE :

1. La MISE EN SITUATION présente le contexte professionnel avec une image mentale de la structure/entreprise, des données chiffrées réalistes et un document support (tableau de données ou graphique nommé).
2. La PROBLÉMATIQUE est UNE SEULE question (avec ?) encadrée en gras juste après la mise en situation. C'est la SEULE question du document avec un point d'interrogation.
3. Les questions sont numérotées (1. 2. 3. ...) et chacune porte le nom de la compétence BO en majuscules gras : **S'APPROPRIER** / **RÉALISER** / **ANALYSER / RAISONNER** / **VALIDER** / **COMMUNIQUER**.
4. Toutes les 5 compétences doivent être évaluées au moins une fois.
5. Les questions comportent des lignes de réponse (______).
6. La dernière question **COMMUNIQUER** demande explicitement de répondre à la problématique initiale.
7. Quand une question nécessite la calculatrice ou un outil, ajouter : 🛎️ APPELER L'EXAMINATEUR
8. L'évaluation est notée sur /10 avec niveaux d'acquisition 0/1/2 par compétence (PAS de barème en points).
9. La fiche d'évaluation (page séparée) liste : capacités et connaissances évaluées du BO, puis le tableau d'évaluation avec compétences / indicateurs / questions / appréciation 0-1-2 / note /10.

## STRUCTURE OBLIGATOIRE :

### MISE EN SITUATION PROFESSIONNELLE
[Présentation de la structure professionnelle ({nom_filiere}), contexte détaillé, données chiffrées]
[Document support : tableau ou données nommées]

L'usage de la calculatrice avec mode examen actif est autorisé.

### PROBLÉMATIQUE
**[Question centrale unique se terminant par ? — clairement encadrée]**

### PARTIE A — [Titre mathématique]

1. **S'APPROPRIER**
- [question d'appropriation de la situation]
______

2. **RÉALISER**
- [question de calcul/réalisation]
______

3. **ANALYSER / RAISONNER**
- [question d'analyse]
______

4. **RÉALISER** 🛎️ APPELER L'EXAMINATEUR
- [question avec outil numérique]
______

5. **VALIDER**
- [question de validation/interprétation]
______

6. **COMMUNIQUER**
- Répondre à la première partie de la problématique.
______

### PARTIE B — [Second thème mathématique]
[Mêmes règles]

---
### FICHE D'ÉVALUATION *(document professeur — NE PAS DISTRIBUER)*

**1. Capacités et connaissances évaluées**
[Liste des capacités et connaissances du BO correspondant au chapitre]

**2. Tableau d'évaluation**
| Compétences | Indicateurs | Questions | Appréciation (0/1/2) |
|---|---|---|---|
| S'approprier | Rechercher, extraire et organiser l'information. | | 0  1  2 |
| Analyser / Raisonner | Émettre des conjectures, proposer une méthode. | | 0  1  2 |
| Réaliser | Mettre en œuvre une méthode, utiliser un modèle, calculer. | | 0  1  2 |
| Valider | Interpréter des résultats, contrôler la vraisemblance. | | 0  1  2 |
| Communiquer | Rendre compte d'un résultat, expliquer une démarche. | | 0  1  2 |
**Note : /10**

### CORRIGÉ DÉTAILLÉ *(document professeur — NE PAS DISTRIBUER)*
[Correction complète question par question avec justifications]

Réponds entièrement en Markdown avec mise en page soignée et professionnelle."""


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
# 4. APPEL API GEMINI
# ============================================================

def call_gemini(api_key, prompt, image=None):
    genai.configure(api_key=api_key.strip())
    model = genai.GenerativeModel(model_name="gemini-2.5-flash")
    try:
        if image:
            img = PIL.Image.open(image)
            response = model.generate_content([prompt, img])
        else:
            response = model.generate_content(prompt)
        if response and response.text:
            return response.text
        return "L'IA a répondu mais le texte est vide. Réessayez."
    except Exception as e:
        if "404" in str(e):
            try:
                model_alt = genai.GenerativeModel("models/gemini-2.5-flash")
                response = model_alt.generate_content([prompt, PIL.Image.open(image)] if image else prompt)
                return response.text
            except Exception as e2:
                raise e2
        raise e


# ============================================================
# 5. EXPORT WORD — VERSION STANDARD
# ============================================================

def markdown_to_docx(md_text, titre="Document"):
    doc = DocxDocument()
    doc.styles["Normal"].font.name = "Arial"
    doc.styles["Normal"].font.size = Pt(11)

    h = doc.add_heading(titre, level=0)
    h.alignment = WD_ALIGN_PARAGRAPH.CENTER

    for line in md_text.split("\n"):
        line = line.rstrip()
        if line.startswith("### "):
            doc.add_heading(line[4:], level=3)
        elif line.startswith("## "):
            doc.add_heading(line[3:], level=2)
        elif line.startswith("# "):
            doc.add_heading(line[2:], level=1)
        elif line.startswith("---"):
            p = doc.add_paragraph()
            p.paragraph_format.space_after = Pt(4)
        elif line.startswith("| "):
            p = doc.add_paragraph(line)
            for run in p.runs:
                run.font.name = "Courier New"
                run.font.size = Pt(9)
        elif re.match(r"^\d+\.", line):
            doc.add_paragraph(line, style="List Number")
        elif line.startswith("- ") or line.startswith("* "):
            doc.add_paragraph(line[2:], style="List Bullet")
        elif line == "":
            doc.add_paragraph("")
        else:
            p = doc.add_paragraph()
            parts = re.split(r"(\*\*.*?\*\*|\*.*?\*)", line)
            for part in parts:
                if part.startswith("**") and part.endswith("**") and len(part) > 4:
                    p.add_run(part[2:-2]).bold = True
                elif part.startswith("*") and part.endswith("*") and len(part) > 2:
                    p.add_run(part[1:-1]).italic = True
                else:
                    p.add_run(part)

    buf = BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf.getvalue()


# ============================================================
# 6. EXPORT WORD — VERSION OFFICIELLE CCF
# ============================================================

def set_cell_border(cell):
    """Ajoute une bordure fine sur toutes les faces d'une cellule."""
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    tcBorders = OxmlElement("w:tcBorders")
    for side in ("top", "left", "bottom", "right"):
        border = OxmlElement(f"w:{side}")
        border.set(qn("w:val"), "single")
        border.set(qn("w:sz"), "4")
        border.set(qn("w:space"), "0")
        border.set(qn("w:color"), "AAAAAA")
        tcBorders.append(border)
    tcPr.append(tcBorders)


def set_run(paragraph, text, bold=False, size=10):
    """Ajoute un run stylé — évite le crash sur paragraphe vide."""
    run = paragraph.add_run(text)
    run.bold = bold
    run.font.size = Pt(size)
    return run


def fill_cell(cell, text, bold=False, size=10):
    """Remplit une cellule proprement sans crasher sur runs vides."""
    p = cell.paragraphs[0]
    p.clear()
    set_run(p, text, bold=bold, size=size)
    set_cell_border(cell)


def generate_ccf_officiel_docx(content_md, metadata, nom_etablissement="Mon Établissement"):
    try:
        doc = DocxDocument()

        # Marges réduites pour plus d'espace
        for section in doc.sections:
            section.top_margin = Inches(0.8)
            section.bottom_margin = Inches(0.8)
            section.left_margin = Inches(1.0)
            section.right_margin = Inches(1.0)

        doc.styles["Normal"].font.name = "Arial"
        doc.styles["Normal"].font.size = Pt(11)

        # ── EN-TÊTE : tableau 2 colonnes ─────────────────────
        table_h = doc.add_table(rows=1, cols=2)
        table_h.autofit = True

        # Gauche : logo Académie de Créteil
        cell_l = table_h.cell(0, 0)
        logo_rep = os.path.join(APP_DIR, "logo_republique.png")
        if os.path.exists(logo_rep):
            cell_l.paragraphs[0].add_run().add_picture(logo_rep, width=Inches(1.4))
        else:
            p_l = cell_l.paragraphs[0]
            run = p_l.add_run("ACADÉMIE DE CRÉTEIL\n")
            run.bold = True
            run.font.size = Pt(11)
            p_l.add_run("Liberté • Égalité • Fraternité").font.size = Pt(9)

        # Droite : logo matière + infos session
        cell_r = table_h.cell(0, 1)
        p_r = cell_r.paragraphs[0]
        p_r.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        logo_mat = os.path.join(APP_DIR, "logo_matiere.png")
        if os.path.exists(logo_mat):
            p_r.add_run().add_picture(logo_mat, width=Inches(0.8))
            p_r.add_run("\n")
        r1 = p_r.add_run(f"Session : {metadata.get('annee_scolaire', '2025/2026')}\n")
        r1.bold = True
        r1.font.size = Pt(10)
        p_r.add_run(f"Situation d'évaluation n° {metadata.get('num_situation', '1')}\n").font.size = Pt(10)
        p_r.add_run(f"Durée : {metadata.get('duree', '45 min')}").font.size = Pt(10)

        doc.add_paragraph()

        # ── TITRE ─────────────────────────────────────────────
        titre = doc.add_heading("CONTRÔLE EN COURS DE FORMATION", level=1)
        titre.alignment = WD_ALIGN_PARAGRAPH.CENTER

        sous_titre = doc.add_paragraph(
            f"Baccalauréat Professionnel — {metadata.get('filiere', '')}"
        )
        sous_titre.alignment = WD_ALIGN_PARAGRAPH.CENTER
        for run in sous_titre.runs:
            run.bold = True
            run.font.size = Pt(11)

        mat_p = doc.add_paragraph(f"Matière : {metadata.get('matiere', '')} | {metadata.get('niveau', '')}")
        mat_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        for run in mat_p.runs:
            run.font.size = Pt(10)

        doc.add_paragraph()

        # ── CADRE CANDIDAT ────────────────────────────────────
        table_c = doc.add_table(rows=2, cols=2)
        candidat_data = [
            ["Nom, Prénom du candidat : ___________________________", "Date : _______________"],
            [f"Établissement : {nom_etablissement}", f"Classe : {metadata.get('niveau', '')} — {metadata.get('filiere', '')}"],
        ]
        for i, row_data in enumerate(candidat_data):
            for j, text in enumerate(row_data):
                fill_cell(table_c.cell(i, j), text, size=10)

        doc.add_paragraph()

        # ── CONTENU DU SUJET ──────────────────────────────────
        skip_keywords = ["Établissement :", "Baccalauréat Professionnel —", "Épreuve E3", "Calculatrice autorisée"]
        for line in content_md.split("\n"):
            line = line.rstrip()
            if any(kw in line for kw in skip_keywords):
                continue
            if line.startswith("#### "):
                doc.add_heading(line[5:].strip(), level=3)
            elif line.startswith("### "):
                doc.add_heading(line[4:].strip(), level=2)
            elif line.startswith("## "):
                doc.add_heading(line[3:].strip(), level=2)
            elif "APPELER L'EXAMINATEUR" in line.upper() or "🛎️" in line:
                p = doc.add_paragraph()
                logo_appel = os.path.join(APP_DIR, "logo_appel.png")
                if os.path.exists(logo_appel):
                    p.add_run().add_picture(logo_appel, width=Inches(0.25))
                    p.add_run("  ")
                run = p.add_run("APPELER L'EXAMINATEUR")
                run.bold = True
                run.font.size = Pt(11)
                pPr = p._p.get_or_add_pPr()
                pBdr = OxmlElement("w:pBdr")
                for side in ("top", "left", "bottom", "right"):
                    border = OxmlElement(f"w:{side}")
                    border.set(qn("w:val"), "single")
                    border.set(qn("w:sz"), "6")
                    border.set(qn("w:space"), "4")
                    border.set(qn("w:color"), "4A6CF7")
                    pBdr.append(border)
                pPr.append(pBdr)
            elif line.startswith("| "):
                p = doc.add_paragraph(line)
                for run in p.runs:
                    run.font.name = "Courier New"
                    run.font.size = Pt(9)
            elif line.startswith("---"):
                continue
            elif line == "":
                doc.add_paragraph("")
            else:
                p = doc.add_paragraph()
                parts = re.split(r"(\*\*.*?\*\*|\*.*?\*)", line)
                for part in parts:
                    if part.startswith("**") and part.endswith("**") and len(part) > 4:
                        p.add_run(part[2:-2]).bold = True
                    elif part.startswith("*") and part.endswith("*") and len(part) > 2:
                        p.add_run(part[1:-1]).italic = True
                    else:
                        p.add_run(part)

        # ── FICHE D'ÉVALUATION OFFICIELLE (page 2) ───────────
        doc.add_page_break()

        titre_fiche = doc.add_heading("FICHE INDIVIDUELLE D'ÉVALUATION", level=1)
        titre_fiche.alignment = WD_ALIGN_PARAGRAPH.CENTER

        # En-tête fiche
        table_ef = doc.add_table(rows=2, cols=2)
        fill_cell(table_ef.cell(0, 0),
                  f"Session : {metadata.get('annee_scolaire','2025/2026')} — Établissement : {nom_etablissement} — Académie : Créteil",
                  bold=True, size=9)
        fill_cell(table_ef.cell(0, 1),
                  f"Spécialité : {metadata.get('matiere','')} — Évaluateur : _______________ — Date : _______________",
                  size=9)
        fill_cell(table_ef.cell(1, 0),
                  f"Situation n° {metadata.get('num_situation','1')} — {metadata.get('filiere','')}",
                  bold=True, size=9)
        fill_cell(table_ef.cell(1, 1), "", size=9)

        doc.add_paragraph()
        cand_p = doc.add_paragraph()
        set_run(cand_p, "Nom et prénom du candidat : _______________________________________________",
                bold=True, size=10)

        doc.add_paragraph()

        # 1. Capacités et connaissances évaluées
        cap_titre = doc.add_paragraph()
        set_run(cap_titre, "1.  Liste des capacités et connaissances évaluées", bold=True, size=11)

        table_cap = doc.add_table(rows=2, cols=2)
        fill_cell(table_cap.cell(0, 0), "Capacités", bold=True, size=10)
        fill_cell(table_cap.cell(0, 1), "(voir section Fiche d'évaluation du sujet généré)", size=10)
        fill_cell(table_cap.cell(1, 0), "Connaissances", bold=True, size=10)
        fill_cell(table_cap.cell(1, 1), "(voir section Fiche d'évaluation du sujet généré)", size=10)

        doc.add_paragraph()

        # 2. Tableau d'évaluation officiel
        eval_titre = doc.add_paragraph()
        set_run(eval_titre, "2.  Évaluation", bold=True, size=11)

        # Calcul nb lignes : 1 en-tête + 1 ligne par indicateur + 1 ligne note
        total_rows = 1
        for comp in COMPETENCES_CCF:
            total_rows += len(comp["indicateurs"])
        total_rows += 1  # ligne note finale

        grid = doc.add_table(rows=total_rows, cols=5)

        # En-têtes colonnes
        h_texts = ["Compétences", "", "Indicateurs / Capacités", "Questions", "Appréciation\n(0 / 1 / 2)"]
        for j, h in enumerate(h_texts):
            fill_cell(grid.cell(0, j), h, bold=True, size=9)

        # Remplissage compétences et indicateurs
        row_idx = 1
        for comp in COMPETENCES_CCF:
            for k, indicateur in enumerate(comp["indicateurs"]):
                fill_cell(grid.cell(row_idx, 0),
                          comp["nom"] if k == 0 else "",
                          bold=(k == 0), size=9)
                fill_cell(grid.cell(row_idx, 1), "", size=9)
                fill_cell(grid.cell(row_idx, 2), indicateur, size=9)
                fill_cell(grid.cell(row_idx, 3), "", size=9)
                fill_cell(grid.cell(row_idx, 4), "0    1    2", size=9)
                row_idx += 1

        # Ligne note finale
        fill_cell(grid.cell(row_idx, 0), "", size=9)
        fill_cell(grid.cell(row_idx, 1), "", size=9)
        fill_cell(grid.cell(row_idx, 2), "", size=9)
        fill_cell(grid.cell(row_idx, 3), "Note :", bold=True, size=10)
        fill_cell(grid.cell(row_idx, 4), "          / 10", bold=True, size=11)

        doc.add_paragraph()
        obs_p = doc.add_paragraph()
        set_run(obs_p, "Observations : ___________________________________________________________________", size=10)

        buf = BytesIO()
        doc.save(buf)
        buf.seek(0)
        return buf.getvalue()

    except Exception as e:
        st.error(f"Erreur lors de la génération du Word officiel : {e}")
        return None


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
for key in ["generated_md", "generated_ccf_md", "correction_md", "meta_gen", "meta_ccf"]:
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
tab_gen, tab_ccf, tab_correction, tab_export = st.tabs([
    "📝 Générateur de Sujets",
    "🎯 Sujets CCF",
    "📸 Correction de Copies",
    "📊 Export Pronote"
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
        chap = st.selectbox("Chapitre (BO)", get_chapitres(mat, niv, cat), key="gen_chap")

    consignes = st.text_area("Consignes particulières (optionnel)", height=80, key="gen_consignes",
                              placeholder="Ex : 3 exercices, niveau accessible, inclure un graphique…")

    if st.button("✨ Générer le sujet", type="primary", use_container_width=True):
        if not cle_api:
            st.error("🔑 Renseigne ta clé API dans le panneau gauche !")
        else:
            with st.spinner("⏳ Génération en cours…"):
                try:
                    res = call_gemini(cle_api, build_prompt_exercices(niv, cat, mat, chap, consignes, filiere))
                    st.session_state.generated_md = res
                    st.session_state.meta_gen = {"niveau": niv, "matiere": mat, "chapitre": chap, "filiere": filiere}
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
        st.markdown(badges, unsafe_allow_html=True)
        st.markdown(st.session_state.generated_md)

        c1, c2, c3 = st.columns(3)
        with c1:
            st.download_button("📥 .md", st.session_state.generated_md,
                               file_name="sujet.md", mime="text/markdown", key="dl1_md")
        with c2:
            st.download_button("📄 .txt", st.session_state.generated_md,
                               file_name="sujet.txt", mime="text/plain", key="dl1_txt")
        with c3:
            st.download_button("📝 .docx",
                markdown_to_docx(st.session_state.generated_md, f"Sujet — {m.get('matiere','')} {m.get('niveau','')}"),
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
        ccf_niv = st.selectbox("Classe", NIVEAUX_CATEGORIES[ccf_cat], key="ccf_niv")
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
        ccf_chap = st.selectbox("Chapitre (BO)", get_chapitres(ccf_mat, ccf_niv, ccf_cat), key="ccf_chap")
        ccf_duree = st.selectbox("Durée de l'épreuve", ["30 min", "45 min", "1h00", "1h30", "2h00"], index=1, key="ccf_duree")
        num_sit = st.number_input("N° de situation CCF", min_value=1, max_value=3, value=1, step=1, key="ccf_num_sit")

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
                            ccf_consignes, ccf_filiere, ccf_duree, str(num_sit)
                        )
                    else:
                        prompt = build_prompt_ccf_entrainement(
                            ccf_niv, ccf_cat, ccf_mat, ccf_chap, ccf_consignes, ccf_filiere
                        )
                    res = call_gemini(cle_api, prompt)
                    if res:
                        st.session_state.generated_ccf_md = res
                        st.session_state.meta_ccf = {
                            "niveau": ccf_niv,
                            "matiere": ccf_mat,
                            "chapitre": ccf_chap,
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
# ONGLET 3 — CORRECTION
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
