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

Retourne un JSON où :
- la clé est le numéro de ligne (ex: "4")
- la valeur est le texte COMPLET de la ligne après remplissage

Règles importantes :
1. Remplis UNIQUEMENT les champs d'identification personnelle :
   nom, prénom, adresse, date/lieu de naissance, email, téléphone,
   nationalité, situation professionnelle, N° sécu, IBAN, SIRET,
   situation de famille.

2. Ne modifie JAMAIS les champs institutionnels :
   structure/composante/direction, N° dossier, visa administratif,
   cachet, signatures de responsables, intitulé du poste,
   service RH, dates d'intervention fixées par l'institution.

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
