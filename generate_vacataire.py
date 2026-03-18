"""
generate_vacataire.py
Module ReportLab pour générer le dossier vacataire officiel.
À ajouter dans le repo GitHub aap-pdf-service.
"""
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.units import mm
from reportlab.lib.styles import ParagraphStyle
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle,
    HRFlowable, PageBreak
)
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.lib.enums import TA_LEFT, TA_CENTER, TA_JUSTIFY
import io
import os
import re
from datetime import datetime

# ── Polices ────────────────────────────────────────────────────────────────────
BASE_FONT = "/usr/share/fonts/truetype/liberation/"
try:
    pdfmetrics.registerFont(TTFont("Cal",    BASE_FONT + "LiberationSans-Regular.ttf"))
    pdfmetrics.registerFont(TTFont("Cal-B",  BASE_FONT + "LiberationSans-Bold.ttf"))
    pdfmetrics.registerFont(TTFont("Cal-I",  BASE_FONT + "LiberationSans-Italic.ttf"))
    pdfmetrics.registerFont(TTFont("Cal-BI", BASE_FONT + "LiberationSans-BoldItalic.ttf"))
    pdfmetrics.registerFontFamily("Cal", normal="Cal", bold="Cal-B",
                                   italic="Cal-I", boldItalic="Cal-BI")
    FONT = "Cal"
except Exception:
    FONT = "Helvetica"

# ── Couleurs ───────────────────────────────────────────────────────────────────
GREY_HEADER  = colors.Color(0.88, 0.88, 0.88)
GREY_LIGHT   = colors.Color(0.96, 0.96, 0.96)
BLUE_HEADER  = colors.Color(0.18, 0.32, 0.52)
BORDER       = colors.Color(0.45, 0.45, 0.45)
BLACK        = colors.black
WHITE        = colors.white

# ── Dimensions ─────────────────────────────────────────────────────────────────
W, H = A4
LM, RM, TM, BM = 20*mm, 20*mm, 22*mm, 20*mm
CW = W - LM - RM

# ── Styles ─────────────────────────────────────────────────────────────────────
def S(name, size=9.5, bold=False, italic=False, align=TA_LEFT,
       sb=0, sa=2, color=BLACK, leading=None):
    fn = FONT + ("-B" if bold and not italic else
                 "-I" if italic and not bold else
                 "-BI" if bold and italic else "")
    if leading is None:
        leading = size * 1.3
    return ParagraphStyle(name, fontName=fn, fontSize=size,
                          leading=leading, alignment=align,
                          spaceBefore=sb, spaceAfter=sa, textColor=color)

sNormal    = S("normal")
sNormalJ   = S("normalJ", align=TA_JUSTIFY)
sBold      = S("bold", bold=True)
sSmall     = S("small", size=8.5)
sSmallI    = S("smallI", size=8.5, italic=True)
sCenter    = S("center", align=TA_CENTER)
sCenterB   = S("centerB", bold=True, align=TA_CENTER, size=11)
sTitle     = S("title", bold=True, size=14, align=TA_CENTER, sb=4, sa=4)
sSection   = S("section", bold=True, size=10, color=WHITE)
sH2        = S("h2", bold=True, size=9.5, sb=4, sa=2)
sCellL     = S("cellL", bold=True, size=9)
sCellV     = S("cellV", size=9)
sBullet    = ParagraphStyle("bullet", fontName=FONT, fontSize=9,
                             leftIndent=12, firstLineIndent=-10,
                             spaceBefore=1, spaceAfter=1,
                             alignment=TA_LEFT)


def P(txt, style=None):
    txt = clean(txt)
    return Paragraph(txt, style or sNormal)


def SP(h=3):
    return Spacer(1, h * mm)


def clean(txt):
    if not txt:
        return ""
    txt = str(txt)
    txt = txt.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
    return txt.strip()


def fmt_date(d):
    """Formate une date ISO en DD/MM/YYYY."""
    if not d:
        return ""
    try:
        return datetime.strptime(str(d)[:10], "%Y-%m-%d").strftime("%d/%m/%Y")
    except Exception:
        return str(d)


def fmt_eur(val):
    """Formate un montant en euros."""
    try:
        return f"{float(val):,.2f} €".replace(",", " ").replace(".", ",")
    except Exception:
        return str(val)


# ── Composants de mise en page ─────────────────────────────────────────────────

def section_header(title):
    """Bandeau de titre de section (fond bleu foncé, texte blanc)."""
    t = Table([[P(title.upper(), sSection)]], colWidths=[CW])
    t.setStyle(TableStyle([
        ("BACKGROUND",    (0, 0), (-1, -1), BLUE_HEADER),
        ("TOPPADDING",    (0, 0), (-1, -1), 5),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
        ("LEFTPADDING",   (0, 0), (-1, -1), 6),
    ]))
    return t


def info_table(rows, c1w=62*mm):
    """Tableau 2 colonnes label / valeur."""
    c2w = CW - c1w
    data = [[P(lab, sCellL), P(val, sCellV)] for lab, val in rows]
    t = Table(data, colWidths=[c1w, c2w])
    t.setStyle(TableStyle([
        ("GRID",           (0, 0), (-1, -1), 0.4, BORDER),
        ("BACKGROUND",     (0, 0), (0, -1),  GREY_LIGHT),
        ("VALIGN",         (0, 0), (-1, -1), "TOP"),
        ("TOPPADDING",     (0, 0), (-1, -1), 3),
        ("BOTTOMPADDING",  (0, 0), (-1, -1), 3),
        ("LEFTPADDING",    (0, 0), (-1, -1), 5),
        ("RIGHTPADDING",   (0, 0), (-1, -1), 5),
    ]))
    return t


def check_table(items_checked):
    """Tableau de cases à cocher avec état (✓ ou ☐)."""
    data = []
    for label, checked in items_checked:
        mark = "✓" if checked else "☐"
        data.append([P(f"{mark}  {label}", sCellV)])
    t = Table(data, colWidths=[CW])
    t.setStyle(TableStyle([
        ("GRID",          (0, 0), (-1, -1), 0.4, BORDER),
        ("TOPPADDING",    (0, 0), (-1, -1), 3),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
        ("LEFTPADDING",   (0, 0), (-1, -1), 8),
    ]))
    return t


def montant_card(heures, taux, montant):
    """Card mise en évidence pour le calcul financier."""
    data = [
        [P("Heures équivalent TD :", sCellL), P(f"{heures} h", sCellV)],
        [P("Taux horaire brut :",    sCellL), P(fmt_eur(taux), sCellV)],
        [P("MONTANT BRUT ESTIMÉ :",
           S("mbe", bold=True, size=10, color=BLUE_HEADER)),
         P(fmt_eur(montant),
           S("mbv", bold=True, size=10, color=BLUE_HEADER))],
    ]
    t = Table(data, colWidths=[80*mm, CW - 80*mm])
    t.setStyle(TableStyle([
        ("BOX",            (0, 0), (-1, -1), 1.0, BLUE_HEADER),
        ("GRID",           (0, 0), (-1, -1), 0.4, BORDER),
        ("BACKGROUND",     (0, 0), (-1, 1),  GREY_LIGHT),
        ("BACKGROUND",     (0, 2), (-1, 2),  colors.Color(0.9, 0.93, 0.97)),
        ("TOPPADDING",     (0, 0), (-1, -1), 4),
        ("BOTTOMPADDING",  (0, 0), (-1, -1), 4),
        ("LEFTPADDING",    (0, 0), (-1, -1), 6),
    ]))
    return t


def signature_table():
    """Bloc signatures 3 colonnes."""
    sig_style = S("sig", size=8.5, align=TA_CENTER)
    data = [[
        P("L'intervenant(e)", sig_style),
        P("Le/La responsable pédagogique", sig_style),
        P("La Direction des Ressources Humaines", sig_style),
    ]]
    t = Table(data, colWidths=[CW/3, CW/3, CW/3])
    t.setStyle(TableStyle([
        ("BOX",           (0, 0), (-1, -1), 0.5, BORDER),
        ("INNERGRID",     (0, 0), (-1, -1), 0.5, BORDER),
        ("TOPPADDING",    (0, 0), (-1, -1), 6),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 40),
        ("LEFTPADDING",   (0, 0), (-1, -1), 4),
        ("ALIGN",         (0, 0), (-1, -1), "CENTER"),
    ]))
    return t


def on_page(canvas, doc):
    """Numérotation + pied de page."""
    canvas.saveState()
    canvas.setFont(FONT, 8)
    canvas.setFillColor(colors.Color(0.5, 0.5, 0.5))
    today = datetime.now().strftime("%d/%m/%Y")
    canvas.drawString(LM, 14*mm,
                      f"Document généré par DocAdmin le {today}")
    canvas.drawRightString(W - RM, 14*mm,
                           f"Page {canvas.getPageNumber()}")
    canvas.restoreState()


# ══════════════════════════════════════════════════════════════════════════════
# FONCTION PRINCIPALE
# ══════════════════════════════════════════════════════════════════════════════

def generate_vacataire_pdf(data: dict) -> bytes:
    """
    Génère le PDF du dossier vacataire à partir du dict `data`.
    Retourne les bytes du PDF.

    Structure attendue de `data` :
    {
      "intervenant": { nom, prenom, date_naissance, lieu_naissance,
                       adresse, telephone, email, situation_pro,
                       siret, employeur_nom, employeur_adresse,
                       employeur_tel, regime_retraite },
      "intervention": { etablissement_nom, etablissement_adresse,
                        composante, responsable_nom, responsable_email,
                        nature, intitule, niveau, heures_etd,
                        date_debut, date_fin, taux_horaire, montant_brut },
      "justificatifs": { pieces_fournies: [...], justification_texte }
    }
    """
    buf = io.BytesIO()
    doc = SimpleDocTemplate(
        buf,
        pagesize=A4,
        leftMargin=LM, rightMargin=RM,
        topMargin=TM,  bottomMargin=BM + 10*mm,
        title="Dossier vacataire",
    )

    i   = data.get("intervenant", {})
    inv = data.get("intervention", {})
    j   = data.get("justificatifs", {})

    nom_complet = f"{i.get('prenom', '')} {i.get('nom', '')}".strip()
    heures  = i.get("heures_etd") or inv.get("heures_etd", 0)
    taux    = i.get("taux_horaire") or inv.get("taux_horaire", 0)
    montant = i.get("montant_brut") or inv.get("montant_brut", 0)
    if not montant and heures and taux:
        try:
            montant = float(heures) * float(taux)
        except Exception:
            montant = 0

    story = []

    # ── EN-TÊTE ────────────────────────────────────────────────────────────────
    story.append(P("DOSSIER DE VACATION", sTitle))
    story.append(P(
        f"Établissement d'accueil : "
        f"{clean(inv.get('etablissement_nom', '…'))}",
        sCenterB
    ))
    story.append(HRFlowable(width=CW, thickness=1.5,
                             color=BLUE_HEADER, spaceAfter=6))
    story.append(SP(2))

    # ── SECTION 1 : INTERVENANT ────────────────────────────────────────────────
    story.append(section_header("1 — Identification de l'intervenant(e)"))
    story.append(info_table([
        ("Nom et Prénom :",          nom_complet),
        ("Date de naissance :",      fmt_date(i.get("date_naissance"))),
        ("Lieu de naissance :",      i.get("lieu_naissance", "")),
        ("Adresse :",                i.get("adresse", "")),
        ("Téléphone :",              i.get("telephone", "")),
        ("Email :",                  i.get("email", "")),
        ("Situation professionnelle :", i.get("situation_pro", "")),
        ("N° SIRET (si applicable) :", i.get("siret", "") or "—"),
        ("Régime de retraite :",     i.get("regime_retraite", "")),
    ]))

    # Employeur principal (si salarié)
    if i.get("employeur_nom"):
        story.append(SP(2))
        story.append(P("Employeur principal :", sH2))
        story.append(info_table([
            ("Nom de l'employeur :",    i.get("employeur_nom", "")),
            ("Adresse :",              i.get("employeur_adresse", "")),
            ("Téléphone :",            i.get("employeur_tel", "")),
        ]))

    story.append(SP(4))

    # ── SECTION 2 : INTERVENTION ───────────────────────────────────────────────
    story.append(section_header("2 — Nature et volume de l'intervention"))
    story.append(info_table([
        ("Établissement d'accueil :",    inv.get("etablissement_nom", "")),
        ("Adresse :",                   inv.get("etablissement_adresse", "")),
        ("Composante / Département :",  inv.get("composante", "")),
        ("Responsable pédagogique :",   inv.get("responsable_nom", "")),
        ("Email responsable :",         inv.get("responsable_email", "")),
        ("Nature de l'intervention :",  inv.get("nature", "")),
        ("Intitulé :",                  inv.get("intitule", "")),
        ("Niveau :",                    inv.get("niveau", "")),
        ("Nombre d'heures ETD :",
         f"{inv.get('heures_etd', '')} heures équivalent TD"),
        ("Période :",
         f"Du {fmt_date(inv.get('date_debut'))} "
         f"au {fmt_date(inv.get('date_fin'))}"),
    ]))
    story.append(SP(3))

    # Calcul financier
    story.append(P("Éléments financiers :", sH2))
    story.append(montant_card(
        inv.get("heures_etd", heures),
        inv.get("taux_horaire", taux),
        montant
    ))
    story.append(P(
        "<i>Montant indicatif avant application des charges et retenues "
        "réglementaires. Le montant net perçu dépend du régime fiscal "
        "et social de l'intervenant(e).</i>",
        sSmallI
    ))
    story.append(SP(4))

    # ── SECTION 3 : PIÈCES JUSTIFICATIVES ─────────────────────────────────────
    story.append(section_header("3 — Pièces justificatives"))
    story.append(SP(1))

    pieces_standard = [
        "Pièce d'identité (CNI ou passeport)",
        "RIB / IBAN",
        "CV et/ou liste de publications",
        "Attestation URSSAF (auto-entrepreneurs)",
        "Justificatif de l'employeur principal",
        "Attestation de non-cumul (fonctionnaires)",
        "Copie du diplôme le plus élevé",
    ]
    pieces_fournies = j.get("pieces_fournies", [])
    if isinstance(pieces_fournies, list):
        pieces_fournies_lower = [p.lower() for p in pieces_fournies]
    else:
        pieces_fournies_lower = []

    items = []
    for piece in pieces_standard:
        checked = any(piece.lower()[:20] in p for p in pieces_fournies_lower)
        items.append((piece, checked))

    # Pièces supplémentaires éventuelles
    for p in pieces_fournies:
        if not any(s.lower()[:20] in p.lower() for s in pieces_standard):
            items.append((p, True))

    story.append(check_table(items))
    story.append(SP(4))

    # ── SECTION 4 : JUSTIFICATION ──────────────────────────────────────────────
    story.append(section_header(
        "4 — Expérience et compétences justifiant l'intervention"
    ))
    justif = j.get("justification_texte", "")
    if justif:
        story.append(SP(2))
        story.append(P(justif, sNormalJ))
    else:
        story.append(SP(2))
        story.append(P(
            "<i>Aucune justification renseignée.</i>", sSmallI
        ))
    story.append(SP(6))

    # ── SECTION 5 : SIGNATURES ─────────────────────────────────────────────────
    story.append(section_header("5 — Signatures"))
    story.append(SP(2))
    story.append(P(
        f"Fait à __________________, le __________________",
        sNormal
    ))
    story.append(SP(4))
    story.append(signature_table())
    story.append(SP(3))
    story.append(P(
        "<i>* En l'absence de signature(s), le dossier ne sera pas traité.</i>",
        sSmallI
    ))

    # ── BUILD ──────────────────────────────────────────────────────────────────
    doc.build(story, onFirstPage=on_page, onLaterPages=on_page)
    return buf.getvalue()
