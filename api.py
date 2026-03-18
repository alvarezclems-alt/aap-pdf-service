"""
api.py — Service Railway complet DocAdmin
Tous les endpoints PDF + IA Claude + remplissage de documents
"""
import io, os, json, traceback, base64
from flask import Flask, request, jsonify, send_file
from flask_cors import CORS

app = Flask(__name__)
CORS(app, origins="*")

# ── Imports optionnels ─────────────────────────────────────────────────────────
try:
    import anthropic
    claude = anthropic.Anthropic(api_key=os.environ.get("ANTHROPIC_API_KEY", ""))
    CLAUDE_MODEL = "claude-sonnet-4-20250514"
    CLAUDE_OK = True
except ImportError:
    CLAUDE_OK = False
    print("WARNING: anthropic non installé")

try:
    from generate_annexes import generate_annexe1_pdf, generate_annexe1bis_pdf
    ANNEXES_OK = True
except ImportError as e:
    ANNEXES_OK = False
    print(f"WARNING: generate_annexes non trouvé: {e}")

try:
    from generate_vacataire import generate_vacataire_pdf
    VACATAIRE_OK = True
except ImportError as e:
    VACATAIRE_OK = False
    print(f"WARNING: generate_vacataire non trouvé: {e}")

try:
    from docx import Document as DocxDoc
    DOCX_OK = True
except ImportError as e:
    DOCX_OK = False
    print(f"WARNING: python-docx non trouvé: {e}")


# ── Utilitaire Claude ──────────────────────────────────────────────────────────
def claude_text(system, user, max_tokens=800):
    if not CLAUDE_OK:
        return "Service IA non disponible"
    msg = claude.messages.create(
        model=CLAUDE_MODEL, max_tokens=max_tokens,
        system=system,
        messages=[{"role": "user", "content": user}]
    )
    return msg.content[0].text.strip()


# ── Health ─────────────────────────────────────────────────────────────────────
@app.route("/health", methods=["GET"])
def health():
    return jsonify({
        "status": "ok",
        "modules": {
            "claude":    CLAUDE_OK,
            "annexes":   ANNEXES_OK,
            "vacataire": VACATAIRE_OK,
            "docx":      DOCX_OK,
        },
        "anthropic_key_set": bool(os.environ.get("ANTHROPIC_API_KEY")),
        "endpoints": [
            "POST /generate-annexe1",
            "POST /generate-annexe1bis",
            "POST /generate-all",
            "POST /generate-vacataire",
            "POST /ai/justification-vacataire",
            "POST /ai/generer-section-aap",
            "POST /ai/reviser-section",
            "POST /ai/extraire-infos",
            "POST /ai/remplir-document",
        ]
    }), 200


# ══════════════════════════════════════════════════════════════════════════════
# PDF AAP
# ══════════════════════════════════════════════════════════════════════════════

@app.route("/generate-annexe1", methods=["POST"])
def generate_annexe1():
    if not ANNEXES_OK:
        return jsonify({"error": "Module generate_annexes non disponible"}), 503
    try:
        data = request.get_json(force=True)
        pdf  = generate_annexe1_pdf(data)
        return send_file(io.BytesIO(pdf), mimetype="application/pdf",
                         as_attachment=True, download_name="annexe1.pdf")
    except Exception as e:
        traceback.print_exc()
        return jsonify({"error": str(e)}), 500


@app.route("/generate-annexe1bis", methods=["POST"])
def generate_annexe1bis():
    if not ANNEXES_OK:
        return jsonify({"error": "Module generate_annexes non disponible"}), 503
    try:
        data = request.get_json(force=True)
        pdf  = generate_annexe1bis_pdf(data)
        return send_file(io.BytesIO(pdf), mimetype="application/pdf",
                         as_attachment=True, download_name="annexe1bis.pdf")
    except Exception as e:
        traceback.print_exc()
        return jsonify({"error": str(e)}), 500


@app.route("/generate-all", methods=["POST"])
def generate_all():
    import zipfile
    if not ANNEXES_OK:
        return jsonify({"error": "Module generate_annexes non disponible"}), 503
    try:
        data = request.get_json(force=True)
        buf  = io.BytesIO()
        with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
            z.writestr("annexe1.pdf",    generate_annexe1_pdf(data))
            z.writestr("annexe1bis.pdf", generate_annexe1bis_pdf(data))
        buf.seek(0)
        return send_file(buf, mimetype="application/zip",
                         as_attachment=True, download_name="dossier-complet.zip")
    except Exception as e:
        traceback.print_exc()
        return jsonify({"error": str(e)}), 500


# ══════════════════════════════════════════════════════════════════════════════
# PDF VACATAIRE
# ══════════════════════════════════════════════════════════════════════════════

@app.route("/generate-vacataire", methods=["POST"])
def generate_vacataire():
    if not VACATAIRE_OK:
        return jsonify({"error": "Module generate_vacataire non disponible"}), 503
    try:
        data = request.get_json(force=True) or {}
        i    = data.get("intervenant", {})
        if not i.get("nom") and not i.get("prenom"):
            return jsonify({"error": "nom/prénom manquant"}), 400
        pdf    = generate_vacataire_pdf(data)
        nom    = i.get("nom", "intervenant").replace(" ", "-").lower()
        prenom = i.get("prenom", "").replace(" ", "-").lower()
        return send_file(io.BytesIO(pdf), mimetype="application/pdf",
                         as_attachment=True,
                         download_name=f"dossier-vacataire-{prenom}-{nom}.pdf")
    except Exception as e:
        traceback.print_exc()
        return jsonify({"error": str(e)}), 500


# ══════════════════════════════════════════════════════════════════════════════
# IA — JUSTIFICATION VACATAIRE
# ══════════════════════════════════════════════════════════════════════════════

@app.route("/ai/justification-vacataire", methods=["POST"])
def ai_justification_vacataire():
    if not CLAUDE_OK:
        return jsonify({"error": "Service Claude non disponible"}), 503
    try:
        d = request.get_json(force=True) or {}
        system = (
            "Tu es un assistant administratif expert dans la rédaction "
            "de dossiers universitaires français. Ton style est factuel, "
            "professionnel et précis. Tu rédiges uniquement en français."
        )
        user = f"""Rédige en 6 à 8 lignes un texte justifiant la compétence
d'un intervenant vacataire. Commence directement par les compétences,
sans phrase d'introduction. Utilise la troisième personne.

Intervenant : {d.get('prenom','')} {d.get('nom','')}
Profil : {d.get('type_profil','')}
Spécialités : {d.get('specialites','')}
Situation : {d.get('situation_pro','')}
Employeur : {d.get('employeur_nom','') or 'Indépendant'}

Mission :
- Intitulé : {d.get('intitule','')}
- Nature : {d.get('nature_intervention','')}
- Niveau : {d.get('niveau','')}
- Établissement : {d.get('etablissement_nom','')}

Texte :"""
        return jsonify({"texte": claude_text(system, user, 600)})
    except Exception as e:
        traceback.print_exc()
        return jsonify({"error": str(e)}), 500


# ══════════════════════════════════════════════════════════════════════════════
# IA — GÉNÉRER SECTION AAP
# ══════════════════════════════════════════════════════════════════════════════

@app.route("/ai/generer-section-aap", methods=["POST"])
def ai_generer_section_aap():
    if not CLAUDE_OK:
        return jsonify({"error": "Service Claude non disponible"}), 503
    try:
        d       = request.get_json(force=True) or {}
        section = d.get("section", "description")
        p       = d.get("projet", {})
        system  = (
            "Tu es un expert en rédaction de dossiers de candidature "
            "pour des appels à projets de recherche en éducation français. "
            "Ton style est académique, structuré et convaincant. "
            "Tu rédiges uniquement en français."
        )
        prompts = {
            "resume": f"""Rédige un résumé de 8 à 10 lignes pour ce projet.
Titre : {p.get('titre','')}
Description : {p.get('description_libre','')}
Axe : {p.get('axe','')} | Mots-clefs : {p.get('mots_clefs','')}
Présente : contexte, approche, résultats attendus, livrables.
Résumé :""",

            "description": f"""Rédige la description scientifique complète (3 pages max).
Structure avec ces sous-titres en gras :
**Inscription générale et problématique**
**Résultats antérieurs, originalité et enjeux**
**Objectifs précis et hypothèses de travail**
**Cadres théoriques, méthodologie et sources**
**Résultats attendus, restitution et diffusion**

Titre : {p.get('titre','')}
Description : {p.get('description_libre','')}
Axe : {p.get('axe','')} | Budget : {p.get('montant','')} €
Description :""",

            "calendrier": f"""Génère un calendrier en tableau Markdown pour ce projet
de recherche sur 2 ans (2027-2028). Format obligatoire :
| Grandes étapes | Début prévisionnel | Fin prévisionnelle | Durée estimée |
|---|---|---|---|
Titre : {p.get('titre','')}
8 à 10 étapes.
Tableau :""",

            "budget": f"""Justifie les dépenses pour un budget de {p.get('montant',8000)} €
sur 2 ans : {p.get('titre','')}
Répartition : missions, prestations, séminaires, vacations, matériel.
Format : liste avec montants en euros.
Justification :""",
        }
        max_t = 2000 if section == "description" else 800
        texte = claude_text(system, prompts.get(section, prompts["description"]), max_t)
        return jsonify({"texte": texte, "section": section})
    except Exception as e:
        traceback.print_exc()
        return jsonify({"error": str(e)}), 500


# ══════════════════════════════════════════════════════════════════════════════
# IA — RÉVISER UNE SECTION
# ══════════════════════════════════════════════════════════════════════════════

@app.route("/ai/reviser-section", methods=["POST"])
def ai_reviser_section():
    if not CLAUDE_OK:
        return jsonify({"error": "Service Claude non disponible"}), 503
    try:
        d = request.get_json(force=True) or {}
        system = (
            "Tu es un expert en rédaction académique française. "
            "Tu révises les textes en respectant les instructions. "
            "Tu retournes uniquement le texte révisé, sans commentaire."
        )
        user = f"""Contexte : {d.get('contexte', 'Dossier administratif')}
Instruction : {d.get('instruction', 'Améliore ce texte')}
Texte original :
{d.get('texte_original', '')}
Texte révisé :"""
        return jsonify({"texte": claude_text(system, user, 1500)})
    except Exception as e:
        traceback.print_exc()
        return jsonify({"error": str(e)}), 500


# ══════════════════════════════════════════════════════════════════════════════
# IA — EXTRAIRE INFOS D'UN TEXTE
# ══════════════════════════════════════════════════════════════════════════════

@app.route("/ai/extraire-infos", methods=["POST"])
def ai_extraire_infos():
    if not CLAUDE_OK:
        return jsonify({"error": "Service Claude non disponible"}), 503
    raw = ""
    try:
        d     = request.get_json(force=True) or {}
        texte = d.get("texte", "")
        typ   = d.get("type", "projet")
        system = (
            "Tu es un assistant expert en extraction d'informations. "
            "Tu retournes UNIQUEMENT un objet JSON valide, "
            "sans texte avant ni après, sans balises markdown."
        )
        schema = (
            '{"titre":"","acronyme":"","mots_clefs":"","resume":"",'
            '"description":"","coordinateur_nom":"","coordinateur_email":"",'
            '"institution":"","laboratoire":"","membres":[],'
            '"montant":0,"axe":""}') if typ == "projet" else (
            '{"nom_aap":"","institution":"","deadline":"",'
            '"budget_max":0,"axes":[],"criteres_eligibilite":"","description":""}')
        user = f"Extrais les informations et retourne ce JSON :\n{schema}\n\nTexte :\n{texte[:6000]}\n\nJSON :"
        raw  = claude_text(system, user, 1200)
        raw  = raw.strip()
        if raw.startswith("```"):
            raw = raw.split("```")[1]
            if raw.startswith("json"):
                raw = raw[4:]
        return jsonify({"infos": json.loads(raw.strip())})
    except json.JSONDecodeError as e:
        return jsonify({"error": f"JSON invalide: {e}", "raw": raw}), 500
    except Exception as e:
        traceback.print_exc()
        return jsonify({"error": str(e)}), 500


# ══════════════════════════════════════════════════════════════════════════════
# IA — REMPLIR UN DOCUMENT ADMINISTRATIF (.docx)
# ══════════════════════════════════════════════════════════════════════════════

def _appliquer_remplacements(doc, remplacements: dict):
    """Remplace les chaînes dans tous les paragraphes et tableaux."""
    for ancien, nouveau in remplacements.items():
        if not isinstance(ancien, str) or not isinstance(nouveau, str):
            continue
        for para in doc.paragraphs:
            if ancien in para.text:
                for run in para.runs:
                    if ancien in run.text:
                        run.text = run.text.replace(ancien, nouveau)
                        break
                else:
                    if para.runs:
                        para.runs[0].text = para.text.replace(ancien, nouveau)
                        for run in para.runs[1:]:
                            run.text = ""
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    for para in cell.paragraphs:
                        if ancien in para.text:
                            for run in para.runs:
                                if ancien in run.text:
                                    run.text = run.text.replace(ancien, nouveau)
                                    break
                            else:
                                if para.runs:
                                    para.runs[0].text = para.text.replace(ancien, nouveau)
                                    for run in para.runs[1:]:
                                        run.text = ""


def _fallback_remplacements(profil: dict) -> dict:
    p = profil
    adresse = f"{p.get('adresse','')} {p.get('code_postal','')} {p.get('ville','')}".strip()
    return {
        "Nom d'usage : ………………………………………………………." : f"Nom d'usage : {p.get('nom','')}",
        "Nom patronymique : …………………………………………"    : f"Nom patronymique : {p.get('nom','')}",
        "Prénom(s) : ……………………………………………………"      : f"Prénom(s) : {p.get('prenom','')}",
        "…….../…..…../…..….."                          : p.get('date_naissance',''),
        "Nationalité : …………………………"                   : f"Nationalité : {p.get('nationalite','Française')}",
        "N° Sécurité Sociale : "                       : f"N° Sécurité Sociale : {p.get('num_secu','')}",
        "Adresse : N°……."                              : f"Adresse : {adresse}",
    }


@app.route("/ai/remplir-document", methods=["POST"])
def ai_remplir_document():
    """
    Reçoit un DOCX en base64 + le profil utilisateur.
    Claude repère tous les champs vides et les remplit avec le profil.
    Retourne le DOCX complété.
    """
    if not CLAUDE_OK:
        return jsonify({"error": "Service Claude non disponible"}), 503
    if not DOCX_OK:
        return jsonify({"error": "python-docx non installé"}), 503
    try:
        data     = request.get_json(force=True) or {}
        profil   = data.get("profil", {})
        b64      = data.get("fichier_base64", "")
        nom_fich = data.get("nom_fichier", "document.docx")

        if not b64:
            return jsonify({"error": "fichier_base64 manquant"}), 400

        docx_bytes = base64.b64decode(b64)
        doc = DocxDoc(io.BytesIO(docx_bytes))

        # Extraire le texte brut
        texte_doc = []
        for para in doc.paragraphs:
            if para.text.strip():
                texte_doc.append(para.text)
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    if cell.text.strip():
                        texte_doc.append(cell.text)
        texte_brut = "\n".join(texte_doc)[:5000]

        system = (
            "Tu es un assistant expert en administration française. "
            "Tu analyses des documents administratifs et identifies "
            "les valeurs exactes à insérer pour chaque champ vide. "
            "Tu retournes UNIQUEMENT un objet JSON valide, "
            "sans texte avant ni après, sans balises markdown."
        )

        user = f"""Voici le texte d'un document administratif français à compléter :

{texte_brut}

Voici le profil de la personne :
{json.dumps(profil, ensure_ascii=False, indent=2)}

Retourne un objet JSON où chaque clé est le texte EXACT à remplacer
dans le document (copié mot pour mot depuis le document),
et la valeur est ce qu'il faut mettre à la place.

Règles :
- Copie exactement les chaînes à remplacer (avec les pointillés, etc.)
- Pour les cases □, remplace □ par ☑ pour la bonne case
- Ne remplis que les champs pour lesquels tu as l'information

Exemple :
{{
  "Nom d'usage : ………………………………………………………." : "Nom d'usage : DUPONT",
  "Prénom(s) : ……………………………………………………" : "Prénom(s) : Marie",
  "Né(e) le : …….../…..…../…..….." : "Né(e) le : 15/03/1985",
  "□ Travailleur non-salarié" : "☑ Travailleur non-salarié"
}}

JSON :"""

        raw = claude_text(system, user, max_tokens=2000)
        raw = raw.strip()
        if raw.startswith("```"):
            parts = raw.split("```")
            raw = parts[1] if len(parts) > 1 else raw
            if raw.startswith("json"):
                raw = raw[4:]
        raw = raw.strip()

        try:
            remplacements = json.loads(raw)
        except json.JSONDecodeError:
            print(f"WARNING: fallback remplacements (JSON invalide)")
            remplacements = _fallback_remplacements(profil)

        doc_out = DocxDoc(io.BytesIO(docx_bytes))
        _appliquer_remplacements(doc_out, remplacements)

        buf = io.BytesIO()
        doc_out.save(buf)
        buf.seek(0)

        p = profil
        if p.get("nom") and p.get("prenom"):
            nom_sortie = f"document-{p['prenom'].lower()}-{p['nom'].lower()}.docx"
        else:
            nom_sortie = f"{os.path.splitext(nom_fich)[0]}-complété.docx"

        return send_file(
            buf,
            mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            as_attachment=True,
            download_name=nom_sortie,
        )

    except Exception as e:
        traceback.print_exc()
        return jsonify({"error": str(e)}), 500


# ══════════════════════════════════════════════════════════════════════════════
# DÉMARRAGE
# ══════════════════════════════════════════════════════════════════════════════

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 8080))
    print(f"Démarrage sur port {port}")
    print(f"  Claude:{CLAUDE_OK} | Annexes:{ANNEXES_OK} | Vacataire:{VACATAIRE_OK} | Docx:{DOCX_OK}")
    app.run(host="0.0.0.0", port=port, debug=False)
