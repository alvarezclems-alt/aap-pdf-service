"""
api.py — Service Railway complet DocAdmin
PDF AAP + PDF Vacataire + IA Claude + Remplissage + Révision de documents
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
    from generate_annexes import build_annexe1, build_annexe1bis
    # Alias pour compatibilité
    def generate_annexe1_pdf(data):
        import os
        logo_path = os.path.join(os.path.dirname(__file__), 'logo_inspe.png')
        logo_bytes = None
        if os.path.exists(logo_path):
            with open(logo_path, 'rb') as f:
                logo_bytes = f.read()
        project = data.get('project', data)
        return build_annexe1(project, logo_bytes)

    def generate_annexe1bis_pdf(data):
        import os
        logo_path = os.path.join(os.path.dirname(__file__), 'logo_inspe.png')
        logo_bytes = None
        if os.path.exists(logo_path):
            with open(logo_path, 'rb') as f:
                logo_bytes = f.read()
        project = data.get('project', data)
        budget_lines = data.get('budget_lines', data.get('lignes', []))
        return build_annexe1bis(project, budget_lines, logo_bytes)

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
            "POST /ai/reviser-document",
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
8 à 10 étapes couvrant : littérature, collecte, analyse, restitution, livrables, valorisation.
Tableau :""",

            "budget": f"""Justifie les dépenses pour un budget de {p.get('montant',8000)} €
sur 2 ans pour ce projet : {p.get('titre','')}
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
# IA — RÉVISER UNE SECTION (champ texte)
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
# IA — RÉVISER UN DOCUMENT AAP COMPLET (chat)
# ══════════════════════════════════════════════════════════════════════════════

@app.route("/ai/reviser-document", methods=["POST"])
def ai_reviser_document():
    """
    Révise un document AAP complet via chat.
    Conserve l'historique de conversation pour les modifications successives.

    Corps attendu :
    {
      "texte_document": "texte complet du document",
      "instruction": "message de l'utilisateur",
      "historique": [
        {"role": "user", "content": "..."},
        {"role": "assistant", "content": "..."}
      ],
      "projet": { "titre": "...", "description": "..." }
    }

    Réponse :
    {
      "reponse": "Résumé des modifications pour l'utilisateur",
      "texte_modifie": "Texte complet du document modifié"
    }
    """
    if not CLAUDE_OK:
        return jsonify({"error": "Service Claude non disponible"}), 503
    try:
        d           = request.get_json(force=True) or {}
        texte       = d.get("texte_document", "")
        instruction = d.get("instruction", "")
        historique  = d.get("historique", [])
        projet      = d.get("projet", {})

        if not texte:
            return jsonify({"error": "texte_document manquant"}), 400
        if not instruction:
            return jsonify({"error": "instruction manquante"}), 400

        system = (
            "Tu es un assistant expert en rédaction de dossiers de candidature "
            "pour des appels à projets de recherche en éducation français. "
            "Quand on te demande de modifier un document, tu appliques les "
            "changements demandés avec précision et tu retournes le texte "
            "complet du document modifié. "
            "Tu rédiges uniquement en français. "
            "Tu retournes UNIQUEMENT un objet JSON valide, "
            "sans texte avant ni après, sans balises markdown."
        )

        # Construire les messages avec l'historique (6 derniers max)
        messages = []
        for msg in historique[-6:]:
            role    = msg.get("role", "user")
            content = msg.get("content", "")
            if role in ("user", "assistant") and content:
                messages.append({"role": role, "content": content})

        # Message courant
        messages.append({
            "role": "user",
            "content": f"""Voici le document AAP actuel :

---
{texte[:8000]}
---

Projet : {projet.get('titre', '')}
Description : {projet.get('description', '')[:500]}

Instruction de modification : {instruction}

Applique cette modification et retourne un JSON avec exactement ces deux clés :
{{
  "reponse": "Résumé clair en 2-3 lignes des modifications apportées",
  "texte_modifie": "Texte COMPLET du document après modification (conserve toute la structure)"
}}

JSON :"""
        })

        response = claude.messages.create(
            model=CLAUDE_MODEL,
            max_tokens=4000,
            system=system,
            messages=messages
        )

        raw = response.content[0].text.strip()

        # Nettoyer les éventuelles balises markdown
        if raw.startswith("```"):
            parts = raw.split("```")
            raw = parts[1] if len(parts) > 1 else raw
            if raw.startswith("json"):
                raw = raw[4:]
        raw = raw.strip()

        try:
            result = json.loads(raw)
            # S'assurer que les deux clés sont présentes
            if "reponse" not in result:
                result["reponse"] = "Document modifié avec succès."
            if "texte_modifie" not in result:
                result["texte_modifie"] = texte
            return jsonify(result)
        except json.JSONDecodeError:
            # Fallback : retourner la réponse brute comme message
            return jsonify({
                "reponse": raw[:500] if raw else "Modification appliquée.",
                "texte_modifie": texte
            })

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
# IA — EXTRAIRE JUSTIFICATIF
# ══════════════════════════════════════════════════════════════════════════════

@app.route("/ai/extraire-justificatif", methods=["POST"])
def ai_extraire_justificatif():
    if not CLAUDE_OK:
        return jsonify({"error": "Service Claude non disponible"}), 503
    try:
        d = request.get_json(force=True) or {}
        b64 = d.get("fichier_base64", "")
        mime = d.get("type_mime", "")
        nom = d.get("nom_fichier", "fichier")
        
        if not b64:
            return jsonify({"error": "Fichier manquant"}), 400
            
        system = (
            "Tu es un expert comptable et administratif. "
            "Tu analyses le fichier fourni (une facture, un billet, un RIB, etc.) "
            "et tu retournes UNIQUEMENT un JSON valide, sans code avant ni après."
        )
        user = f"""Analyse ce document ({nom}) et retourne un JSON avec cette structure (ne mets que les champs que tu trouves, ignore les autres) :
{{
  "type_document": "Choisis parmi: billet_train, billet_avion, facture_hotel, taxi, peage, rib, autre",
  "resume": "Courte description de 3-4 mots (ex: Billet SNCF Paris-Lyon, Facture Ibis, RIB)",
  "informations": {{
    "montant": Nombre flottant (ex: 45.50),
    "devise": "EUR",
    "date": "JJ/MM/AAAA",
    "origine": "Ville de départ",
    "destination": "Ville d'arrivée",
    "iban": "FR...",
    "bic": "...",
    "titulaire": "Nom du titulaire"
  }}
}}"""
        
        content = []
        if mime == "application/pdf":
            content.append({
                "type": "document",
                "source": {
                    "type": "base64",
                    "media_type": "application/pdf",
                    "data": b64
                }
            })
        elif mime in ["image/jpeg", "image/png", "image/webp", "image/gif"]:
            content.append({
                "type": "image",
                "source": {
                    "type": "base64",
                    "media_type": mime,
                    "data": b64
                }
            })
        elif mime == "application/vnd.openxmlformats-officedocument.wordprocessingml.document" or nom.endswith(".docx"):
            import io
            from docx import Document as DocxDoc
            docx_bytes = base64.b64decode(b64)
            doc = DocxDoc(io.BytesIO(docx_bytes))
            texte = "\\n".join([p.text for p in doc.paragraphs])
            content.append({
                "type": "text", 
                "text": f"Contenu du DOCX :\\n{texte[:10000]}\\n\\n"
            })
        else:
            return jsonify({"error": f"Format non supporté par l'IA: {mime}"}), 400
            
        content.append({"type": "text", "text": user})
        
        response = claude.messages.create(
            model=CLAUDE_MODEL,
            max_tokens=1000,
            system=system,
            messages=[{"role": "user", "content": content}]
        )
        
        raw = response.content[0].text.strip()
        if raw.startswith("```"):
            parts = raw.split("```")
            raw = parts[1] if len(parts) > 1 else raw
            if raw.startswith("json"):
                raw = raw[4:]
        raw = raw.strip()
        
        return jsonify(json.loads(raw))
    except Exception as e:
        traceback.print_exc()
        return jsonify({"error": str(e)}), 500


# ══════════════════════════════════════════════════════════════════════════════
# IA — REMPLIR UN DOCUMENT ADMINISTRATIF (.doc / .docx)
# ══════════════════════════════════════════════════════════════════════════════

def _fusionner_runs(para):
    """Fusionne tous les runs d'un paragraphe en un seul."""
    if len(para.runs) <= 1:
        return
    texte_complet = para.text
    if not texte_complet.strip():
        return
    para.runs[0].text = texte_complet
    for run in para.runs[1:]:
        run.text = ""


def _remplacer_dans_para(para, ancien, nouveau):
    """Remplace une chaîne dans un paragraphe, fusionne les runs si nécessaire."""
    if ancien not in para.text:
        return False
    for run in para.runs:
        if ancien in run.text:
            run.text = run.text.replace(ancien, nouveau)
            return True
    _fusionner_runs(para)
    for run in para.runs:
        if ancien in run.text:
            run.text = run.text.replace(ancien, nouveau)
            return True
    return False


def _appliquer_remplacements(doc, remplacements: dict):
    """Remplace les chaînes dans tous les paragraphes et tableaux."""
    for ancien, nouveau in remplacements.items():
        if not isinstance(ancien, str) or not isinstance(nouveau, str):
            continue
        for para in doc.paragraphs:
            _remplacer_dans_para(para, ancien, nouveau)
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    for para in cell.paragraphs:
                        _remplacer_dans_para(para, ancien, nouveau)


def _decomposer_adresse(adresse: str):
    """
    Décompose une adresse complète en numéro + rue.
    Ex: '50 rue saint jacques' → ('50', 'rue saint jacques')
    """
    import re
    adresse = adresse.strip()
    m = re.match(r'^(\d+\s*(?:bis|ter|quater)?)\s+(.+)$', adresse, re.IGNORECASE)
    if m:
        return m.group(1).strip(), m.group(2).strip()
    return "", adresse


def _convert_doc_to_docx(doc_bytes: bytes, nom_fich: str) -> bytes:
    """Convertit un .doc en .docx via LibreOffice headless."""
    import subprocess, tempfile, glob, os
    with tempfile.TemporaryDirectory() as tmpdir:
        input_path = os.path.join(tmpdir, nom_fich)
        with open(input_path, "wb") as f:
            f.write(doc_bytes)
            
        ext = os.path.splitext(nom_fich)[1].lower()
        cmd = ["libreoffice", "--headless"]
        if ext == ".pdf":
            cmd.extend(["--infilter=writer_pdf_import"])
        cmd.extend(["--convert-to", "docx", "--outdir", tmpdir, input_path])
        
        result = subprocess.run(
            cmd,
            capture_output=True, text=True, timeout=60
        )
        if result.returncode != 0:
            raise RuntimeError(f"LibreOffice erreur: {result.stderr}")
        docx_files = glob.glob(os.path.join(tmpdir, "*.docx"))
        if not docx_files:
            raise RuntimeError("Conversion LibreOffice : aucun .docx produit")
        with open(docx_files[0], "rb") as f:
            return f.read()

def _force_pdf_to_docx(pdf_bytes: bytes, nom_fich: str) -> bytes:
    import tempfile, os
    from pdf2docx import Converter
    with tempfile.TemporaryDirectory() as tmpdir:
        pdf_path = os.path.join(tmpdir, "input.pdf")
        docx_path = os.path.join(tmpdir, "output.docx")
        with open(pdf_path, "wb") as f:
            f.write(pdf_bytes)
        
        cv = Converter(pdf_path)
        cv.convert(docx_path)
        cv.close()
        
        with open(docx_path, "rb") as f:
            return f.read()


def _convert_docx_to_pdf(docx_bytes: bytes, nom_fich: str) -> bytes:
    """Convertit un .docx en .pdf via LibreOffice headless."""
    import subprocess, tempfile, glob, os
    with tempfile.TemporaryDirectory() as tmpdir:
        input_path = os.path.join(tmpdir, "doc.docx")
        with open(input_path, "wb") as f:
            f.write(docx_bytes)
        result = subprocess.run(
            ["libreoffice", "--headless", "--convert-to", "pdf",
             "--outdir", tmpdir, input_path],
            capture_output=True, text=True, timeout=60
        )
        if result.returncode != 0:
            raise RuntimeError(f"LibreOffice erreur: {result.stderr}")
        pdf_files = glob.glob(os.path.join(tmpdir, "*.pdf"))
        if not pdf_files:
            raise RuntimeError("Conversion LibreOffice : aucun .pdf produit")
        with open(pdf_files[0], "rb") as f:
            return f.read()


@app.route("/ai/remplir-document", methods=["POST"])
def ai_remplir_document():
    """
    Reçoit un DOC ou DOCX en base64 + le profil utilisateur.
    - Convertit automatiquement .doc → .docx via LibreOffice
    - Claude identifie les champs par numéro de ligne et les remplit
    - Retourne le DOCX complété en binaire

    Corps attendu :
    {
      "fichier_base64": "...",
      "nom_fichier": "fiche.docx",
      "profil": {
        "civilite": "M.", "nom": "Dupont", "prenom": "Marie",
        "date_naissance": "15/03/1985", "lieu_naissance": "Paris",
        "adresse": "12 rue de la Paix", "code_postal": "75001", "ville": "Paris",
        "telephone": "06 12 34 56 78", "email": "marie@example.com",
        "situation_pro": "Auto-entrepreneur", "siret": "123 456 789 00012",
        "nationalite": "Française", "pays": "France",
        "num_secu": "1 85 03 75 056 789 42", "iban": "FR76...",
        "situation_famille": "Célibataire"
      }
    }
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
        justifs  = data.get("justificatifs", [])

        if not b64:
            return jsonify({"error": "fichier_base64 manquant"}), 400

        raw_bytes = base64.b64decode(b64)

        ext = os.path.splitext(nom_fich)[1].lower()
        if ext == ".pdf":
            print(f"Surimpression PDF natif activée : {nom_fich}")
            import fitz
            import json
            try:
                doc_pdf = fitz.open(stream=raw_bytes, filetype="pdf")
                pdf_text = ""
                for page in doc_pdf:
                    pdf_text += page.get_text("text") + "\n"
                
                prompt = (
                    "Tu es un data-entry bot strict expert en traitement de formulaires.\n"
                    "Voici le texte brut d'un formulaire PDF vierge que tu dois remplir.\n"
                    f"PROFIL UTILISATEUR : {json.dumps(profil, ensure_ascii=False)}\n"
                    f"JUSTIFICATIFS A INJECTER : {json.dumps(justifs, ensure_ascii=False)}\n\n"
                    "INSTRUCTIONS CRUCIALES :\n"
                    "1. Remplis UNIQUEMENT les zones de saisie prévues (lignes pointillées, champs vides, tableau 'Catégories/Montant').\n"
                    "2. NE REMPLIS JAMAIS les paragraphes d'instruction en haut du PDF (ex: 'remboursement maximum', 'si vous voyagez en...').\n"
                    "3. Rends UNIQUEMENT un objet JSON valide. Pas de markdown, de bonjour ou d'explications.\n"
                    "4. Les clés json sont les textes EXACTS (incluant la ponctuation comme ':') situés juste avant la zone à remplir (ex: 'Nom :', 'Prénom :', 'Train'). Ne coupe pas les ':'.\n"
                    "5. Chaque valeur est un objet { \"val\": \"...\", \"type\": \"...\", \"col_header\": \"...\" }\n\n"
                    "TYPES D'ALIGNEMENT ('type') :\n"
                    "- 'right' : champ texte classique (ex: 'Nom :', 'Ville :').\n"
                    "- 'check' : pour cocher une case (ex: '□ M.'). La valeur doit être 'X'.\n"
                    "- 'column' : pour un tableau de dépenses (ex: étiquette 'Train' ou 'Avion', provenant des JUSTIFICATIFS). col_header doit être l'en-tête de colonne ('Montant').\n\n"
                    "Exemple de réponse attendue :\n"
                    "{\n"
                    "  \"Nom :\": {\"val\": \"Dupont\", \"type\": \"right\"},\n"
                    "  \"□ M.\": {\"val\": \"X\", \"type\": \"check\"},\n"
                    "  \"Train\": {\"val\": \"350\", \"type\": \"column\", \"col_header\": \"Montant\"}\n"
                    "}\n\n"
                    f"TEXTE DU PDF À ANALYSER :\n{pdf_text[:12000]}"
                )
                
                response = claude.messages.create(
                    model=CLAUDE_MODEL,
                    max_tokens=2048,
                    temperature=0.0,
                    system="Tu es un robot JSON.",
                    messages=[{"role": "user", "content": prompt}]
                )
                
                reponse_texte = response.content[0].text.strip()
                if reponse_texte.startswith("```json"):
                    reponse_texte = reponse_texte.replace("```json", "", 1)
                if reponse_texte.endswith("```"):
                    reponse_texte = reponse_texte[:-3]
                
                ai_mapping = json.loads(reponse_texte.strip())
                print("MAPPING PDF INTELLIGENT TROUVÉ: ", ai_mapping)
                
                for page_num in range(doc_pdf.page_count):
                    page = doc_pdf[page_num]
                    for label, config in ai_mapping.items():
                        if not isinstance(config, dict) or "val" not in config:
                            continue
                        
                        tval = config.get("val", "")
                        atype = config.get("type", "right")
                        if not tval: continue
                        
                        rects = page.search_for(str(label))
                        if rects:
                            # TRES IMPORTANT: On ne prend QUE la dernière occurrence du mot sur la page.
                            # Les formulaires placent quasiment toujours les instructions (textes explicatifs) 
                            # en haut de la page, et les vrais champs de saisie ou tableaux tout en bas.
                            r = rects[-1]
                            
                            if atype == "check":
                                page.insert_text((r.x0 + 1, r.y1 - 2), "X", fontsize=12, color=(0, 0, 0.6))
                            elif atype == "column":
                                col_header = config.get("col_header", "")
                                col_x = r.x1 + 80
                                if col_header:
                                    h_rects = page.search_for(str(col_header))
                                    if h_rects:
                                        # Prendre aussi le dernier header s'il y en a plusieurs
                                        h_r = h_rects[-1]
                                        col_x = (h_r.x0 + h_r.x1) / 2 - 10
                                page.insert_text((col_x, r.y1 - 1), str(tval), fontsize=10, color=(0, 0, 0.6))
                            else:
                                # increased padding to +35 to avoid colon / space overlap if label missed dots
                                page.insert_text((r.x1 + 35, r.y1 - 1), str(tval), fontsize=10, color=(0, 0, 0.6))
                                
                pdf_buf = io.BytesIO(doc_pdf.write())
                doc_pdf.close()
                pdf_buf.seek(0)
                
                return send_file(
                    pdf_buf,
                    mimetype="application/pdf",
                    as_attachment=True,
                    download_name=nom_fich
                )
            except Exception as err:
                print(f"WARNING: Erreur de surimpression PDF: {err}")
                return jsonify({
                    "error": f"Impossible d'imprimer sur ce PDF nativement : {err}"
                }), 422
        elif ext == ".doc":
            print(f"Conversion .doc → .docx : {nom_fich}")
            try:
                docx_bytes = _convert_doc_to_docx(raw_bytes, nom_fich)
                nom_fich   = os.path.splitext(nom_fich)[0] + ".docx"
            except Exception as conv_err:
                print(f"WARNING: Conversion échouée: {conv_err}")
                return jsonify({
                    "error": "Conversion .doc impossible. Ouvrez le fichier dans Word et enregistrez-le en .docx."
                }), 422
        else:
            docx_bytes = raw_bytes

        doc = DocxDoc(io.BytesIO(docx_bytes))

        # ── Décomposer l'adresse en numéro + rue ─────────────────────────────
        adresse_brute = profil.get("adresse", "")
        num_rue, nom_rue = _decomposer_adresse(adresse_brute)

        profil_enrichi = dict(profil)
        profil_enrichi["numero_rue"]       = num_rue
        profil_enrichi["nom_rue"]          = nom_rue
        profil_enrichi["adresse_complete"] = (
            f"{adresse_brute}, {profil.get('code_postal','')} "
            f"{profil.get('ville','')}".strip(", ")
        )

        # ── Enrichir profil avec les infos des justificatifs ──────────────────
        billets_train = [j for j in justifs if j.get("type_document") == "billet_train"]
        rib = next((j for j in justifs if j.get("type_document") == "rib"), None)

        if billets_train:
            total_train = sum(
                float(j.get("informations", {}).get("montant") or 0)
                for j in billets_train
            )
            dates_train = [
                j.get("informations", {}).get("date", "")
                for j in billets_train
                if j.get("informations", {}).get("date")
            ]
            profil_enrichi["frais_train"] = str(total_train)
            if dates_train:
                profil_enrichi["date_depart"] = dates_train[0]
                profil_enrichi["date_arrivee"] = dates_train[-1]

        if rib:
            info_rib = rib.get("informations", {})
            profil_enrichi["iban"]             = info_rib.get("iban", profil.get("iban", ""))
            profil_enrichi["bic"]              = info_rib.get("bic", profil.get("bic", ""))
            profil_enrichi["titulaire_compte"] = info_rib.get("titulaire", "")
            profil_enrichi["banque_nom"]       = info_rib.get("banque", "La Banque Postale")
            profil_enrichi["banque_adresse"]   = info_rib.get("adresse", "")

        # ── Construire la liste numérotée des paragraphes ─────────────────────
        champs_doc = []
        for para in doc.paragraphs:
            if para.text.strip():
                champs_doc.append(para.text)
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    for para in cell.paragraphs:
                        if para.text.strip():
                            champs_doc.append(para.text)

        champs_str = "\n".join(
            f"[{i}] {t}" for i, t in enumerate(champs_doc)
        )[:5000]

        # Construire le bloc justificatifs pour le prompt
        justif_bloc = ""
        if justifs:
            justif_lines = "\n".join([
                f"- {j.get('type_document','')}: "
                f"{j.get('resume', '')} "
                f"| montant: {j.get('informations',{}).get('montant','')}"
                f" | date: {j.get('informations',{}).get('date','')}"
                f" | trajet: {j.get('informations',{}).get('description','')}"
                f" | IBAN: {j.get('informations',{}).get('iban','')}"
                f" | BIC: {j.get('informations',{}).get('bic','')}"
                for j in justifs
            ])
            justif_bloc = f"""
Informations extraites des justificatifs fournis :
{justif_lines}
Utilise ces informations pour remplir les champs correspondants :
- Dates de départ/arrivée → champs date
- Montants train/avion/taxi → colonnes montant
- Trajet → champs origine/destination
- IBAN/BIC → coordonnées bancaires
- Total → montant total du tableau
"""

        system = (
            "Tu es un assistant expert en administration française. "
            "Tu analyses des documents administratifs et identifies "
            "les champs personnels à remplir pour un intervenant. "
            "Tu retournes UNIQUEMENT un objet JSON valide, "
            "sans texte avant ni après, sans balises markdown."
        )

        user = f"""Voici les lignes numérotées d'un document administratif :

{champs_str}

Voici le profil de la personne :
{json.dumps(profil_enrichi, ensure_ascii=False, indent=2)}
{justif_bloc}
Retourne un JSON où :
- la clé est le numéro de ligne (ex: "4")
- la valeur est le texte COMPLET de la ligne après remplissage

Règles IMPORTANTES pour identifier les champs à remplir :
✓ REMPLIR uniquement si la cellule/ligne contient :
  - Des pointillés : ............. ou …………………
  - Des tirets : ____________
  - Est complètement vide
  - Contient uniquement (jj/mm/aa) ou similaire
✗ NE JAMAIS modifier si la cellule contient :
  - Un label descriptif : 'Départ :', 'Train', 'Avion', 'Taxi', 'Catégories', 'Montant'
  - Du texte qui décrit le champ lui-même
  - Des instructions ou explications
  - Des titres de colonnes ou de lignes

La règle clé : si supprimer le texte existant rendrait le document incompréhensible,
NE PAS modifier cette cellule.

Règles supplémentaires :
3. Pour le champ Adresse qui contient "N°... Bât... Rue..." :
   - Utilise "numero_rue" pour le N°
   - Utilise "nom_rue" pour la Rue
   - Laisse Bât vide s'il n'est pas renseigné
   Ex: "Adresse : N°50   Bât. :    Rue : rue saint jacques"

4. Pour le champ "Code Postal... Ville..." :
   - Mets UNIQUEMENT le code postal et la ville
   - Ne mets PAS le téléphone dans ce champ
   Ex: "Code Postal : 75005  Ville : Paris"

5. Le téléphone a SA PROPRE ligne dans le document.
   Cherche une ligne avec "Tél" ou "Téléphone" et mets-y le numéro.

6. Pour les cases à cocher □ :
   - Coche la bonne case avec ☑
   - Laisse les autres □
   - Conserve l'espacement exact de la ligne

7. N'invente aucune information absente du profil.
   Si une info manque, laisse la ligne inchangée.

Exemple :
{{
  "4": "Nom d'usage : DUPONT",
  "6": "Prénom(s) : Marie",
  "8": "Né(e) le : 15/03/1985 à Paris",
  "10": "Adresse : N°12   Bât. :    Rue : rue de la Paix",
  "11": "Code Postal : 75001  Ville : Paris",
  "12": "Tél : 06 12 34 56 78",
  "13": "Mail : marie@example.com",
  "15": "Nationalité : Française",
  "18": "Situation de famille : Célibataire",
  "20": "□ Personnel du secteur privé   □ Personnel du secteur public   ☑ Travailleur non-salarié   □ Retraité"
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
            index_map = json.loads(raw)
        except json.JSONDecodeError:
            print(f"WARNING: JSON invalide de Claude: {raw[:200]}")
            index_map = {}

        # ── Appliquer les modifications par index ─────────────────────────────
        doc_out   = DocxDoc(io.BytesIO(docx_bytes))
        paras_out = []
        for para in doc_out.paragraphs:
            if para.text.strip():
                paras_out.append(para)
        for table in doc_out.tables:
            for row in table.rows:
                for cell in row.cells:
                    for para in cell.paragraphs:
                        if para.text.strip():
                            paras_out.append(para)

        for idx_str, nouveau_texte in index_map.items():
            try:
                idx = int(idx_str)
                if 0 <= idx < len(paras_out):
                    para = paras_out[idx]
                    _fusionner_runs(para)
                    if para.runs:
                        para.runs[0].text = nouveau_texte
                    else:
                        from docx.oxml import OxmlElement
                        r = OxmlElement('w:r')
                        t = OxmlElement('w:t')
                        t.text = nouveau_texte
                        r.append(t)
                        para._p.append(r)
            except (ValueError, IndexError) as e:
                print(f"WARNING: index {idx_str} invalide: {e}")

        # ── Sauvegarder et retourner ──────────────────────────────────────────
        buf = io.BytesIO()
        doc_out.save(buf)
        buf.seek(0)
        
        try:
            pdf_bytes = _convert_docx_to_pdf(buf.getvalue(), nom_fich)
            pdf_buf = io.BytesIO(pdf_bytes)
        except Exception as conv_pdf_err:
             return jsonify({"error": f"Erreur lors de la génération du PDF final: {conv_pdf_err}"}), 500

        p = profil
        if p.get("nom") and p.get("prenom"):
            nom_sortie = f"document-{p['prenom'].lower()}-{p['nom'].lower()}.pdf"
        else:
            nom_sortie = f"{os.path.splitext(nom_fich)[0]}-complété.pdf"

        return send_file(
            pdf_buf,
            mimetype="application/pdf",
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


# ══════════════════════════════════════════════════════════════════════════════
# IA — ANALYSER UN DOCUMENT ET IDENTIFIER LES CHAMPS MANQUANTS
# ══════════════════════════════════════════════════════════════════════════════

@app.route("/ai/analyser-document", methods=["POST"])
def ai_analyser_document():
    """
    Analyse un document et identifie les champs personnels à remplir.
    Compare avec le profil existant pour ne demander que ce qui manque.

    Corps attendu :
    {
      "texte_document": "texte extrait du document",
      "profil_existant": {
        "nom": "...", "prenom": "...", "email": "...", ...
      }
    }

    Réponse :
    {
      "champs_manquants": [
        { "id": "num_secu", "label": "N° de sécurité sociale",
          "type": "texte", "exemple": "1 85 03 75 056 789 42" },
        { "id": "situation_famille", "label": "Situation familiale",
          "type": "choix",
          "options": ["Célibataire", "Marié(e)", "Pacsé(e)", "Divorcé(e)", "Veuf/Veuve"] }
      ],
      "champs_trouves": ["nom", "prenom", "email", "adresse"]
    }
    """
    if not CLAUDE_OK:
        return jsonify({"error": "Service Claude non disponible"}), 503
    raw = ""
    try:
        d      = request.get_json(force=True) or {}
        texte  = d.get("texte_document", "")
        profil = d.get("profil_existant", {})
        justifs = d.get("justificatifs", [])

        if not texte:
            return jsonify({"error": "texte_document manquant"}), 400

        # Extraire les infos utiles des justificatifs
        infos_justifs = {}
        for j in justifs:
            info = j.get("informations", {})
            type_doc = j.get("type_document", "")
            if type_doc == "billet_train":
                if info.get("montant") is not None:
                    infos_justifs["frais_train"] = info.get("montant")
                if info.get("date"):
                    infos_justifs["date_depart"] = info.get("date")
            if type_doc == "rib":
                if info.get("iban"):
                    infos_justifs["iban"] = info.get("iban")
                if info.get("bic"):
                    infos_justifs["bic"] = info.get("bic")
                if info.get("titulaire"):
                    infos_justifs["titulaire_compte"] = info.get("titulaire")
                infos_justifs["banque_nom"] = info.get("banque", "La Banque Postale")

        # Combiner profil + infos justificatifs
        profil_complet = {**profil, **infos_justifs}

        # Déterminer quels champs du profil complet sont déjà renseignés
        champs_profil_renseignes = {
            k: v for k, v in profil_complet.items()
            if v and str(v).strip() not in ("", "null", "None")
        }

        system = (
            "Tu es un assistant administratif expert. "
            "Tu analyses des documents administratifs français et identifies "
            "les champs personnels qui doivent être remplis. "
            "Tu retournes UNIQUEMENT un objet JSON valide, "
            "sans texte avant ni après, sans balises markdown."
        )

        user = f"""Voici le texte d'un document administratif à remplir :

{texte[:4000]}

Voici les informations déjà disponibles dans le profil de l'utilisateur :
{json.dumps(champs_profil_renseignes, ensure_ascii=False, indent=2)}

Identifie tous les champs personnels présents dans ce document.
Pour chaque champ, indique s'il est déjà disponible dans le profil.

Retourne ce JSON :
{{
  "champs_manquants": [
    {{
      "id": "identifiant_snake_case",
      "label": "Libellé clair pour l'utilisateur",
      "type": "texte" ou "choix" ou "date",
      "exemple": "exemple de valeur (optionnel)",
      "options": ["opt1", "opt2"]
    }}
  ],
  "champs_trouves": ["liste des champs déjà disponibles dans le profil"]
}}

Champs personnels à détecter dans le document :
- nom (nom d'usage)
- prenom
- nom_naissance (nom de naissance / nom patronymique)
- date_naissance
- lieu_naissance
- departement_naissance
- pays_naissance
- nationalite
- adresse
- code_postal
- ville
- telephone_domicile
- telephone_portable
- email
- num_secu (numéro de sécurité sociale)
- iban
- bic
- siret
- situation_famille (Célibataire/Marié(e)/Pacsé(e)/Divorcé(e)/Veuf/Concubinage)
- situation_pro (secteur public/privé/indépendant/retraité/étudiant)
- employeur_nom
- employeur_adresse

Pour "situation_famille", utilise type "choix" avec options :
["Célibataire", "Marié(e)", "Pacsé(e)", "Divorcé(e)", "Veuf/Veuve", "Concubinage"]

Pour "situation_pro", utilise type "choix" avec options :
["Secteur public", "Secteur privé salarié", "Travailleur non-salarié / Auto-entrepreneur", "Intermittent du spectacle", "Étudiant(e) 3ème cycle", "Retraité(e)"]

Ne liste dans champs_manquants QUE les champs présents dans le document
ET absents (ou vides) dans le profil fourni.

JSON :"""

        raw = claude_text(system, user, max_tokens=1500)
        raw = raw.strip()
        if raw.startswith("```"):
            parts = raw.split("```")
            raw = parts[1] if len(parts) > 1 else raw
            if raw.startswith("json"):
                raw = raw[4:]
        raw = raw.strip()

        result = json.loads(raw)

        # S'assurer que les deux clés sont présentes
        if "champs_manquants" not in result:
            result["champs_manquants"] = []
        if "champs_trouves" not in result:
            result["champs_trouves"] = list(champs_profil_renseignes.keys())

        return jsonify(result)

    except json.JSONDecodeError as e:
        print(f"WARNING: JSON invalide de Claude: {raw[:200]}")
        # Fallback : retourner les champs les plus courants comme manquants
        return jsonify({
            "champs_manquants": [
                {"id": "num_secu",          "label": "N° de sécurité sociale",  "type": "texte", "exemple": "1 85 03 75 056 789 42"},
                {"id": "iban",              "label": "IBAN",                      "type": "texte", "exemple": "FR76 1234 5678 9012 3456 7890 123"},
                {"id": "situation_famille", "label": "Situation familiale",        "type": "choix",
                 "options": ["Célibataire", "Marié(e)", "Pacsé(e)", "Divorcé(e)", "Veuf/Veuve", "Concubinage"]},
            ],
            "champs_trouves": list(profil.keys())
        })
    except Exception as e:
        traceback.print_exc()
        return jsonify({"error": str(e)}), 500


# ══════════════════════════════════════════════════════════════════════════════
# PDF DEVIS
# ══════════════════════════════════════════════════════════════════════════════

try:
    from generate_devis import generate_devis_pdf
    DEVIS_OK = True
except ImportError as e:
    DEVIS_OK = False
    print(f"WARNING: generate_devis non trouvé: {e}")


@app.route("/generate-devis", methods=["POST"])
def generate_devis():
    """
    Génère le PDF d'un devis.
    Corps : { emetteur, client, devis, lignes, totaux }
    """
    if not DEVIS_OK:
        return jsonify({"error": "Module generate_devis non disponible"}), 503
    try:
        data   = request.get_json(force=True) or {}
        pdf    = generate_devis_pdf(data)
        dv     = data.get("devis", {})
        cl     = data.get("client", {})
        numero = dv.get("numero", "devis").replace(" ", "-").replace("/", "-")
        client = cl.get("nom", "client")[:20].replace(" ", "-")
        return send_file(
            io.BytesIO(pdf),
            mimetype="application/pdf",
            as_attachment=True,
            download_name=f"Devis-{numero}-{client}.pdf",
        )
    except Exception as e:
        traceback.print_exc()
        return jsonify({"error": str(e)}), 500


# ══════════════════════════════════════════════════════════════════════════════
# PDF ORDRE DE MISSION
# ══════════════════════════════════════════════════════════════════════════════

try:
    from generate_ordre_mission import generate_ordre_mission_pdf
    MISSION_OK = True
except ImportError as e:
    MISSION_OK = False
    print(f"WARNING: generate_ordre_mission non trouvé: {e}")


@app.route("/generate-ordre-mission", methods=["POST"])
def generate_ordre_mission():
    """
    Génère le PDF de la demande d'autorisation de déplacement.
    Corps : { missionnaire, mission, transport, frais, etranger }
    """
    if not MISSION_OK:
        return jsonify({"error": "Module generate_ordre_mission non disponible"}), 503
    try:
        data = request.get_json(force=True) or {}
        pdf  = generate_ordre_mission_pdf(data)
        mis  = data.get("missionnaire", {})
        nom  = mis.get("nom", "mission").replace(" ", "-").lower()
        msn  = data.get("mission", {})
        date = str(msn.get("date_debut", ""))[:10].replace("-", "")
        return send_file(
            io.BytesIO(pdf),
            mimetype="application/pdf",
            as_attachment=True,
            download_name=f"OrdeMission-{nom}-{date}.pdf",
        )
    except Exception as e:
        traceback.print_exc()
        return jsonify({"error": str(e)}), 500
