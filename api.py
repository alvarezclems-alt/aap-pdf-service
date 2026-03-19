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

def _fusionner_runs(para):
    """
    Fusionne tous les runs d'un paragraphe en un seul pour faciliter
    les remplacements. Conserve le formatage du premier run.
    """
    if len(para.runs) <= 1:
        return
    texte_complet = para.text
    if not texte_complet.strip():
        return
    # Mettre tout le texte dans le premier run, vider les autres
    para.runs[0].text = texte_complet
    for run in para.runs[1:]:
        run.text = ""


def _remplacer_dans_para(para, ancien, nouveau):
    """Remplace une chaîne dans un paragraphe, en fusionnant les runs si nécessaire."""
    if ancien not in para.text:
        return False
    # Essai direct sur les runs existants
    for run in para.runs:
        if ancien in run.text:
            run.text = run.text.replace(ancien, nouveau)
            return True
    # Fusionner les runs puis remplacer
    _fusionner_runs(para)
    for run in para.runs:
        if ancien in run.text:
            run.text = run.text.replace(ancien, nouveau)
            return True
    return False


def _appliquer_remplacements(doc, remplacements: dict):
    """Remplace les chaînes dans tous les paragraphes et tableaux.
    Fusionne les runs si nécessaire pour que les substitutions fonctionnent."""
    for ancien, nouveau in remplacements.items():
        if not isinstance(ancien, str) or not isinstance(nouveau, str):
            continue
        # Paragraphes directs
        for para in doc.paragraphs:
            _remplacer_dans_para(para, ancien, nouveau)
        # Paragraphes dans les tableaux
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    for para in cell.paragraphs:
                        _remplacer_dans_para(para, ancien, nouveau)


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


def _convert_doc_to_docx(doc_bytes: bytes, nom_fich: str) -> bytes:
    """Convertit un .doc en .docx via LibreOffice headless."""
    import subprocess, tempfile, glob
    with tempfile.TemporaryDirectory() as tmpdir:
        input_path = os.path.join(tmpdir, nom_fich)
        with open(input_path, "wb") as f:
            f.write(doc_bytes)
        result = subprocess.run(
            ["libreoffice", "--headless", "--convert-to", "docx",
             "--outdir", tmpdir, input_path],
            capture_output=True, text=True, timeout=60
        )
        if result.returncode != 0:
            raise RuntimeError(f"LibreOffice erreur: {result.stderr}")
        docx_files = glob.glob(os.path.join(tmpdir, "*.docx"))
        if not docx_files:
            raise RuntimeError("Conversion LibreOffice : aucun .docx produit")
        with open(docx_files[0], "rb") as f:
            return f.read()


@app.route("/ai/remplir-document", methods=["POST"])
def ai_remplir_document():
    """
    Reçoit un DOC ou DOCX en base64 + le profil utilisateur.
    Convertit automatiquement les .doc en .docx via LibreOffice.
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

        raw_bytes = base64.b64decode(b64)

        # ── Conversion automatique .doc → .docx ──────────────────────────────
        ext = os.path.splitext(nom_fich)[1].lower()
        if ext == ".doc":
            print(f"Conversion .doc → .docx : {nom_fich}")
            try:
                docx_bytes = _convert_doc_to_docx(raw_bytes, nom_fich)
                nom_fich   = os.path.splitext(nom_fich)[0] + ".docx"
            except Exception as conv_err:
                print(f"WARNING: Conversion échouée: {conv_err}")
                return jsonify({
                    "error": "Conversion .doc impossible. "
                             "Ouvrez le fichier dans Word et enregistrez-le en .docx."
                }), 422
        else:
            docx_bytes = raw_bytes

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

        # ── Extraire les paragraphes champ par champ ──────────────────────────
        # Construire un mapping : texte_para → para_index pour chaque cellule
        champs_doc = []
        for para in doc.paragraphs:
            if para.text.strip():
                champs_doc.append({"texte": para.text, "source": "para"})
        for ti, table in enumerate(doc.tables):
            for ri, row in enumerate(table.rows):
                for ci, cell in enumerate(row.cells):
                    for pi, para in enumerate(cell.paragraphs):
                        if para.text.strip():
                            champs_doc.append({
                                "texte": para.text,
                                "source": f"table_{ti}_{ri}_{ci}_{pi}"
                            })

        # Limiter à 5000 chars
        champs_str = "\n".join(
            f"[{i}] {c['texte']}" for i, c in enumerate(champs_doc)
        )[:5000]

        system = (
            "Tu es un assistant expert en administration française. "
            "Tu analyses des documents administratifs et identifies "
            "les champs à remplir. "
            "Tu retournes UNIQUEMENT un objet JSON valide, "
            "sans texte avant ni après, sans balises markdown."
        )

        user = f"""Voici les lignes d'un document administratif numérotées :

{champs_str}

Voici le profil de la personne :
{json.dumps(profil, ensure_ascii=False, indent=2)}

Pour chaque ligne qui contient un champ vide (avec des ……, des □, ou du texte à compléter),
retourne un JSON avec :
- clé = le numéro de ligne entre crochets (ex: "3")
- valeur = le texte COMPLET de la ligne tel qu'il doit apparaître après remplissage

Ne retourne que les lignes à modifier. Pour les cases à cocher □,
coche la bonne case avec ☑ et laisse les autres □.

Exemple de réponse :
{{
  "4": "Nom d'usage : DUPONT",
  "8": "Prénom(s) : Marie",
  "10": "Né(e) le : 06/10/2001 à Toulouse",
  "15": "☑ Travailleur non-salarié"
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
            print(f"WARNING: JSON invalide de Claude, utilisation fallback")
            index_map = {}

        # ── Appliquer par index de paragraphe ─────────────────────────────────
        doc_out = DocxDoc(io.BytesIO(docx_bytes))

        # Reconstruire la même liste ordonnée sur doc_out
        paras_out = []
        for para in doc_out.paragraphs:
            if para.text.strip():
                paras_out.append(para)
        for ti, table in enumerate(doc_out.tables):
            for ri, row in enumerate(table.rows):
                for ci, cell in enumerate(row.cells):
                    for pi, para in enumerate(cell.paragraphs):
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
                        # Pas de run : ajouter un nouveau
                        from docx.oxml.ns import qn
                        from docx.oxml import OxmlElement
                        r = OxmlElement('w:r')
                        t = OxmlElement('w:t')
                        t.text = nouveau_texte
                        r.append(t)
                        para._p.append(r)
            except (ValueError, IndexError) as e:
                print(f"WARNING: index {idx_str} invalide: {e}")

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
