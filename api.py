"""
api.py — Service Railway complet
Backend PDF + IA Claude pour DocAdmin

Variables d'environnement requises sur Railway :
  ANTHROPIC_API_KEY=sk-ant-...
  PORT=8080  (déjà en place)
"""

import io
import os
import json
import traceback

import anthropic
from flask import Flask, request, jsonify, send_file
from flask_cors import CORS

from generate_annexes import (
    generate_annexe1_pdf,
    generate_annexe1bis_pdf,
)
from generate_vacataire import generate_vacataire_pdf

# ── App Flask ──────────────────────────────────────────────────────────────────
app = Flask(__name__)
CORS(app, origins="*")

# ── Client Claude ──────────────────────────────────────────────────────────────
claude = anthropic.Anthropic(api_key=os.environ.get("ANTHROPIC_API_KEY", ""))
CLAUDE_MODEL = "claude-sonnet-4-20250514"


# ══════════════════════════════════════════════════════════════════════════════
# UTILITAIRES
# ══════════════════════════════════════════════════════════════════════════════

def claude_text(system: str, user: str, max_tokens: int = 800) -> str:
    """Appel Claude et retourne le texte brut de la réponse."""
    msg = claude.messages.create(
        model=CLAUDE_MODEL,
        max_tokens=max_tokens,
        system=system,
        messages=[{"role": "user", "content": user}],
    )
    return msg.content[0].text.strip()


# ══════════════════════════════════════════════════════════════════════════════
# ENDPOINTS EXISTANTS — AAP
# ══════════════════════════════════════════════════════════════════════════════

@app.route("/generate-annexe1", methods=["POST"])
def generate_annexe1():
    try:
        data = request.get_json(force=True)
        pdf_bytes = generate_annexe1_pdf(data)
        return send_file(
            io.BytesIO(pdf_bytes),
            mimetype="application/pdf",
            as_attachment=True,
            download_name="annexe1.pdf",
        )
    except Exception as e:
        traceback.print_exc()
        return jsonify({"error": str(e)}), 500


@app.route("/generate-annexe1bis", methods=["POST"])
def generate_annexe1bis():
    try:
        data = request.get_json(force=True)
        pdf_bytes = generate_annexe1bis_pdf(data)
        return send_file(
            io.BytesIO(pdf_bytes),
            mimetype="application/pdf",
            as_attachment=True,
            download_name="annexe1bis.pdf",
        )
    except Exception as e:
        traceback.print_exc()
        return jsonify({"error": str(e)}), 500


@app.route("/generate-all", methods=["POST"])
def generate_all():
    import zipfile
    try:
        data = request.get_json(force=True)
        pdf1    = generate_annexe1_pdf(data)
        pdf1bis = generate_annexe1bis_pdf(data)

        buf = io.BytesIO()
        with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
            z.writestr("annexe1.pdf",    pdf1)
            z.writestr("annexe1bis.pdf", pdf1bis)
        buf.seek(0)

        return send_file(
            buf,
            mimetype="application/zip",
            as_attachment=True,
            download_name="dossier-complet.zip",
        )
    except Exception as e:
        traceback.print_exc()
        return jsonify({"error": str(e)}), 500


# ══════════════════════════════════════════════════════════════════════════════
# ENDPOINT NOUVEAU — PDF VACATAIRE
# ══════════════════════════════════════════════════════════════════════════════

@app.route("/generate-vacataire", methods=["POST"])
def generate_vacataire():
    """
    Génère le PDF du dossier vacataire.
    Corps JSON : { intervenant, intervention, justificatifs }
    """
    try:
        data = request.get_json(force=True)
        if not data:
            return jsonify({"error": "Corps JSON vide"}), 400

        i = data.get("intervenant", {})
        if not i.get("nom") and not i.get("prenom"):
            return jsonify({
                "error": "Champs obligatoires manquants : nom/prénom"
            }), 400

        pdf_bytes = generate_vacataire_pdf(data)

        nom    = i.get("nom", "intervenant").replace(" ", "-").lower()
        prenom = i.get("prenom", "").replace(" ", "-").lower()
        filename = f"dossier-vacataire-{prenom}-{nom}.pdf"

        return send_file(
            io.BytesIO(pdf_bytes),
            mimetype="application/pdf",
            as_attachment=True,
            download_name=filename,
        )
    except Exception as e:
        traceback.print_exc()
        return jsonify({"error": str(e)}), 500


# ══════════════════════════════════════════════════════════════════════════════
# ENDPOINTS IA — CLAUDE
# ══════════════════════════════════════════════════════════════════════════════

@app.route("/ai/justification-vacataire", methods=["POST"])
def ai_justification_vacataire():
    """
    Génère le texte de justification d'un dossier vacataire via Claude.

    Corps attendu :
    {
      "prenom": "Marie",
      "nom": "Dupont",
      "type_profil": "Auto-entrepreneur",
      "specialites": "IA générative, formation, SIC",
      "situation_pro": "Auto-entrepreneur",
      "employeur_nom": "",
      "intitule": "Introduction à l'IA générative",
      "nature_intervention": "Cours magistral",
      "niveau": "Master 2",
      "etablissement_nom": "Université de Lille"
    }
    """
    try:
        d = request.get_json(force=True) or {}

        system = (
            "Tu es un assistant administratif expert dans la rédaction "
            "de dossiers universitaires français. "
            "Ton style est factuel, professionnel et précis. "
            "Tu n'utilises pas de formules creuses ni d'adverbes superflus. "
            "Tu rédiges uniquement en français."
        )

        user = f"""Rédige en 6 à 8 lignes un texte justifiant la compétence 
d'un intervenant vacataire pour une mission d'enseignement universitaire.

Commence directement par les compétences sans phrase d'introduction 
du type "Je soussigné..." ou "Monsieur/Madame X...".
Utilise la troisième personne.
Sois concret : mentionne les domaines d'expertise, l'expérience pratique 
et le lien avec la mission confiée.

Informations sur l'intervenant :
- Nom et prénom : {d.get('prenom', '')} {d.get('nom', '')}
- Type de profil : {d.get('type_profil', '')}
- Spécialités / domaines : {d.get('specialites', '')}
- Situation professionnelle : {d.get('situation_pro', '')}
- Employeur principal : {d.get('employeur_nom', '') or 'Indépendant'}

Mission d'enseignement :
- Intitulé : {d.get('intitule', '')}
- Nature : {d.get('nature_intervention', '')}
- Niveau : {d.get('niveau', '')}
- Établissement : {d.get('etablissement_nom', '')}

Texte de justification :"""

        texte = claude_text(system, user, max_tokens=600)
        return jsonify({"texte": texte})

    except anthropic.AuthenticationError:
        return jsonify({
            "error": "Clé API Claude invalide. "
                     "Vérifiez ANTHROPIC_API_KEY sur Railway."
        }), 401
    except Exception as e:
        traceback.print_exc()
        return jsonify({"error": str(e)}), 500


@app.route("/ai/generer-section-aap", methods=["POST"])
def ai_generer_section_aap():
    """
    Génère une section d'un dossier AAP via Claude.

    Corps attendu :
    {
      "section": "description",   // description | resume | calendrier | budget
      "projet": {
        "titre": "...",
        "description_libre": "...",
        "axe": "...",
        "mots_clefs": "...",
        "coordinateur": "...",
        "membres": [...],
        "montant": 8000
      }
    }
    """
    try:
        d       = request.get_json(force=True) or {}
        section = d.get("section", "description")
        projet  = d.get("projet", {})

        system = (
            "Tu es un expert en rédaction de dossiers de candidature "
            "pour des appels à projets de recherche en éducation français "
            "(ANR, INSPÉ, Région, etc.). "
            "Ton style est académique, structuré et convaincant. "
            "Tu respectes les contraintes de longueur demandées. "
            "Tu rédiges uniquement en français."
        )

        prompts = {
            "resume": f"""Rédige un résumé de 8 à 10 lignes maximum pour le projet 
de recherche suivant, destiné à être publié sur le site de l'institution.

Titre : {projet.get('titre', '')}
Description : {projet.get('description_libre', '')}
Axe thématique : {projet.get('axe', '')}
Mots-clefs : {projet.get('mots_clefs', '')}
Équipe : {projet.get('coordinateur', '')} + {len(projet.get('membres', []))} membre(s)

Le résumé doit présenter : le contexte et la problématique, 
l'approche méthodologique, les résultats attendus et les livrables.
Résumé :""",

            "description": f"""Rédige la description scientifique complète (3 pages max) 
du projet de recherche suivant pour un dossier de candidature AAP.

Structure ta réponse avec ces sous-sections en gras :
**Inscription générale et problématique**
**Résultats antérieurs, originalité et enjeux**
**Objectifs précis et hypothèses de travail**
**Cadres théoriques, méthodologie et sources**
**Résultats attendus, restitution et diffusion**

Titre : {projet.get('titre', '')}
Description libre : {projet.get('description_libre', '')}
Axe : {projet.get('axe', '')}
Mots-clefs : {projet.get('mots_clefs', '')}
Coordinateur : {projet.get('coordinateur', '')}
Budget demandé : {projet.get('montant', '')} €

Description :""",

            "calendrier": f"""Génère un calendrier prévisionnel en tableau Markdown 
pour le projet de recherche suivant. Le projet dure 2 ans (2027-2028).

Format du tableau (obligatoire) :
| Grandes étapes | Début prévisionnel | Fin prévisionnelle | Durée estimée |
|---|---|---|---|
| ... | ... | ... | ... |

Titre : {projet.get('titre', '')}
Description : {projet.get('description_libre', '')}

Génère 8 à 10 étapes réalistes couvrant : revue de littérature, 
collecte de données, analyse, restitution, livrables, valorisation.
Tableau :""",

            "budget": f"""Génère une justification des dépenses pour un projet 
de recherche avec un budget de {projet.get('montant', 8000)} € sur 2 ans.

Titre du projet : {projet.get('titre', '')}
Description : {projet.get('description_libre', '')}

Propose une répartition réaliste entre :
- Frais de mission (déplacements, hébergement)
- Prestations de recherche
- Organisation de séminaires / journées d'étude
- Vacations éventuelles
- Matériel consommable

Format : liste structurée avec montants estimés en euros.
Justification :""",
        }

        prompt_user = prompts.get(section, prompts["description"])
        texte = claude_text(system, prompt_user,
                            max_tokens=2000 if section == "description" else 800)
        return jsonify({"texte": texte, "section": section})

    except anthropic.AuthenticationError:
        return jsonify({
            "error": "Clé API Claude invalide. "
                     "Vérifiez ANTHROPIC_API_KEY sur Railway."
        }), 401
    except Exception as e:
        traceback.print_exc()
        return jsonify({"error": str(e)}), 500


@app.route("/ai/reviser-section", methods=["POST"])
def ai_reviser_section():
    """
    Révise/améliore un texte existant via Claude.

    Corps attendu :
    {
      "texte_original": "...",
      "instruction": "Rends ce texte plus concis",
      "contexte": "Section description d'un AAP recherche en éducation"
    }
    """
    try:
        d = request.get_json(force=True) or {}

        system = (
            "Tu es un expert en rédaction académique française. "
            "Tu révises des textes en respectant scrupuleusement "
            "les instructions données. "
            "Tu retournes uniquement le texte révisé, sans commentaire."
        )

        user = f"""Contexte : {d.get('contexte', 'Dossier administratif')}

Instruction : {d.get('instruction', 'Améliore ce texte')}

Texte original :
{d.get('texte_original', '')}

Texte révisé :"""

        texte = claude_text(system, user, max_tokens=1500)
        return jsonify({"texte": texte})

    except anthropic.AuthenticationError:
        return jsonify({
            "error": "Clé API Claude invalide. "
                     "Vérifiez ANTHROPIC_API_KEY sur Railway."
        }), 401
    except Exception as e:
        traceback.print_exc()
        return jsonify({"error": str(e)}), 500


@app.route("/ai/extraire-infos", methods=["POST"])
def ai_extraire_infos():
    """
    Extrait les informations structurées d'un texte brut de projet.
    Utile pour le wizard "Nouvel AAP" — import texte → pré-remplissage.

    Corps attendu :
    {
      "texte": "Texte brut du projet ou de l'AAP...",
      "type": "projet"   // projet | aap
    }
    """
    try:
        d    = request.get_json(force=True) or {}
        texte = d.get("texte", "")
        typ   = d.get("type", "projet")

        system = (
            "Tu es un assistant expert en extraction d'informations. "
            "Tu retournes UNIQUEMENT un objet JSON valide, "
            "sans texte avant ni après, sans balises markdown."
        )

        if typ == "projet":
            schema = """{
  "titre": "...",
  "acronyme": "...",
  "mots_clefs": "...",
  "resume": "...",
  "description": "...",
  "coordinateur_nom": "...",
  "coordinateur_email": "...",
  "institution": "...",
  "laboratoire": "...",
  "membres": [{"nom": "...", "email": "...", "grade": "..."}],
  "montant": 0,
  "axe": "..."
}"""
        else:
            schema = """{
  "nom_aap": "...",
  "institution": "...",
  "deadline": "...",
  "budget_max": 0,
  "axes": ["..."],
  "criteres_eligibilite": "...",
  "description": "..."
}"""

        user = f"""Extrais les informations du texte suivant et retourne 
un JSON avec exactement cette structure :
{schema}

Si une information est absente, utilise null ou une chaîne vide.

Texte à analyser :
{texte[:6000]}

JSON :"""

        raw = claude_text(system, user, max_tokens=1200)

        # Nettoyer les éventuelles balises markdown
        raw = raw.strip()
        if raw.startswith("```"):
            raw = raw.split("```")[1]
            if raw.startswith("json"):
                raw = raw[4:]
        raw = raw.strip()

        infos = json.loads(raw)
        return jsonify({"infos": infos})

    except json.JSONDecodeError as e:
        return jsonify({
            "error": f"Réponse Claude non parseable : {str(e)}",
            "raw": raw if "raw" in dir() else ""
        }), 500
    except anthropic.AuthenticationError:
        return jsonify({
            "error": "Clé API Claude invalide. "
                     "Vérifiez ANTHROPIC_API_KEY sur Railway."
        }), 401
    except Exception as e:
        traceback.print_exc()
        return jsonify({"error": str(e)}), 500


# ══════════════════════════════════════════════════════════════════════════════
# SANTÉ ET DÉMARRAGE
# ══════════════════════════════════════════════════════════════════════════════

@app.route("/health", methods=["GET"])
def health():
    key_ok = bool(os.environ.get("ANTHROPIC_API_KEY", ""))
    return jsonify({
        "status": "ok",
        "claude_model": CLAUDE_MODEL,
        "anthropic_key_set": key_ok,
        "endpoints": {
            "pdf": [
                "POST /generate-annexe1",
                "POST /generate-annexe1bis",
                "POST /generate-all",
                "POST /generate-vacataire",
            ],
            "ai": [
                "POST /ai/justification-vacataire",
                "POST /ai/generer-section-aap",
                "POST /ai/reviser-section",
                "POST /ai/extraire-infos",
            ],
        }
    }), 200


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 8080))
    app.run(host="0.0.0.0", port=port, debug=False)
