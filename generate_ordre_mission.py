"""
generate_ordre_mission.py
Générateur PDF pour demande d'autorisation de déplacement
Fidèle au modèle demande_d_autorisation_de_deplacement_-v3-2022.docx
"""
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.units import mm
from reportlab.lib.styles import ParagraphStyle
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, HRFlowable
)
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.lib.enums import TA_LEFT, TA_CENTER, TA_RIGHT
import io
from datetime import datetime

# ── Polices ────────────────────────────────────────────────────────────────────
BASE_FONT = "/usr/share/fonts/truetype/liberation/"
try:
    pdfmetrics.registerFont(TTFont("Cal",   BASE_FONT + "LiberationSans-Regular.ttf"))
    pdfmetrics.registerFont(TTFont("Cal-B", BASE_FONT + "LiberationSans-Bold.ttf"))
    pdfmetrics.registerFont(TTFont("Cal-I", BASE_FONT + "LiberationSans-Italic.ttf"))
    FONT = "Cal"
except Exception:
    FONT = "Helvetica"

# ── Couleurs ───────────────────────────────────────────────────────────────────
NOIR        = colors.Color(0.13, 0.12, 0.12)   # 221E1F
ORANGE_ETR  = colors.Color(0.60, 0.28, 0.02)   # 984806 — section étranger
VERT        = colors.Color(0.0,  0.69, 0.31)   # 00B050
JAUNE       = colors.Color(0.80, 0.70, 0.0)    # CBB301
ORANGE_W    = colors.Color(0.89, 0.42, 0.04)   # E36C0A
ROUGE       = colors.Color(1.0,  0.0,  0.0)    # FF0000
GRIS_HDR    = colors.Color(0.85, 0.85, 0.85)
GRIS_LIGHT  = colors.Color(0.95, 0.95, 0.95)
BLANC       = colors.white

# ── Dimensions ─────────────────────────────────────────────────────────────────
W, H   = A4
LM = RM = 18 * mm
TM = BM = 15 * mm
CW = W - LM - RM


# ── Helpers styles ─────────────────────────────────────────────────────────────
def S(name, size=9, bold=False, italic=False,
      align=TA_LEFT, color=NOIR, sb=0, sa=1):
    fn = (FONT + "-B" if bold and not italic else
          FONT + "-I" if italic else FONT)
    return ParagraphStyle(name, fontName=fn, fontSize=size,
                          leading=size * 1.35, alignment=align,
                          spaceBefore=sb, spaceAfter=sa, textColor=color)

sN   = S("n")
sB   = S("b", bold=True)
sT   = S("t", size=11, bold=True)
sTT  = S("tt", size=13, bold=True)
sSub = S("sub", size=8, italic=True)
sC   = S("c", align=TA_CENTER)
sCB  = S("cb", bold=True, align=TA_CENTER)
sR   = S("r", align=TA_RIGHT)
sOr  = S("or", bold=True, color=ORANGE_ETR, size=9)
sV   = S("v", bold=True, color=VERT, size=9)
sJ   = S("j", bold=True, color=JAUNE, size=9)
sOW  = S("ow", bold=True, color=ORANGE_W, size=9)
sRg  = S("rg", bold=True, color=ROUGE, size=9)


def P(txt, style=None):
    txt = str(txt).replace("&","&amp;").replace("<","&lt;").replace(">","&gt;")
    return Paragraph(txt, style or sN)


def SP(h=3):
    return Spacer(1, h * mm)


def fmt_date(d):
    if not d:
        return ""
    try:
        return datetime.strptime(str(d)[:10], "%Y-%m-%d").strftime("%d/%m/%Y")
    except Exception:
        return str(d)


def check(val, label, checked_style=None, unchecked_style=None):
    """Retourne ☑ label ou ☐ label selon val."""
    mark = "☑" if val else "☐"
    style = (checked_style or sB) if val else (unchecked_style or sN)
    return P(f"{mark}  {label}", style)


def section_box(title, content_rows, col_widths=None):
    """Crée un tableau avec titre en en-tête gris."""
    if col_widths is None:
        col_widths = [CW]
    data = [[P(title, sB)]] + content_rows
    t = Table(data, colWidths=col_widths)
    n = len(data)
    t.setStyle(TableStyle([
        ("BACKGROUND",    (0,0), (-1,0), GRIS_HDR),
        ("BOX",           (0,0), (-1,-1), 0.5, NOIR),
        ("INNERGRID",     (0,0), (-1,-1), 0.3, colors.Color(0.7,0.7,0.7)),
        ("TOPPADDING",    (0,0), (-1,-1), 3),
        ("BOTTOMPADDING", (0,0), (-1,-1), 3),
        ("LEFTPADDING",   (0,0), (-1,-1), 5),
        ("VALIGN",        (0,0), (-1,-1), "TOP"),
        ("SPAN",          (0,0), (-1,0)),
    ]))
    return t


def on_page(canvas, doc):
    canvas.saveState()
    canvas.setFont(FONT, 7)
    canvas.setFillColor(colors.Color(0.5,0.5,0.5))
    canvas.drawCentredString(W/2, 10*mm, f"Page {canvas.getPageNumber()}")
    canvas.restoreState()


# ══════════════════════════════════════════════════════════════════════════════
# FONCTION PRINCIPALE
# ══════════════════════════════════════════════════════════════════════════════

def generate_ordre_mission_pdf(data: dict) -> bytes:
    """
    Génère le PDF de la demande d'autorisation de déplacement.

    Structure attendue :
    {
      "missionnaire": {
        "nom": "DUPONT", "prenom": "Marie",
        "composante": "GERiiCO / Université de Lille"
      },
      "mission": {
        "objet": "Conférence internationale sur l'IA en éducation",
        "destination": "Paris", "pays": "France",
        "date_debut": "2026-04-15", "heure_debut": "08:00",
        "date_fin": "2026-04-17",   "heure_fin": "18:00",
        "etranger": false
      },
      "transport": {
        "type_deplacement": "avec_frais",
        "train": true, "avion": false,
        "transports_commun": false, "autre": "",
        "remboursement_direct": true,
        "bon_commande": false,
        "vehicule_personnel": false,
        "motif_vehicule": ""
      },
      "frais": {
        "nb_repas": 2, "nb_nuitees": 1,
        "autres_depenses": "",
        "avance": false, "montant_avance": 0,
        "centre_cout": "", "otp": "", "domaine": ""
      },
      "etranger": {
        "niveau_securite": "vert",
        "avis_medecin": ""
      }
    }
    """
    buf = io.BytesIO()
    doc = SimpleDocTemplate(
        buf, pagesize=A4,
        leftMargin=LM, rightMargin=RM,
        topMargin=TM, bottomMargin=BM + 8*mm,
        title="Demande d'autorisation de déplacement",
    )

    mis  = data.get("missionnaire", {})
    msn  = data.get("mission", {})
    trp  = data.get("transport", {})
    fra  = data.get("frais", {})
    etr  = data.get("etranger", {})
    etranger = msn.get("etranger", False)

    story = []

    # ── TITRE ─────────────────────────────────────────────────────────────────
    story.append(P("DEMANDE D'AUTORISATION DE DÉPLACEMENT", sTT))
    story.append(P("Déplacement en France ou à l'étranger quelle que soit la durée", sT))
    story.append(P("(en dehors des congés réguliers)", S("sub2", size=9, italic=True)))
    story.append(SP(3))
    story.append(HRFlowable(width=CW, thickness=1.5, color=NOIR, spaceAfter=4))

    # ── TYPE DE DÉPLACEMENT ────────────────────────────────────────────────────
    type_depl = trp.get("type_deplacement", "avec_frais")
    type_data = [[
        check(type_depl == "avec_frais",   "AVEC FRAIS pour l'Université de Lille"),
        check(type_depl == "partiel",       "PARTIEL"),
        check(type_depl == "sans_frais",   "SANS FRAIS pour l'Université de Lille"),
    ]]
    type_t = Table(type_data, colWidths=[CW*0.4, CW*0.2, CW*0.4])
    type_t.setStyle(TableStyle([
        ("BOX",          (0,0), (-1,-1), 0.5, NOIR),
        ("TOPPADDING",   (0,0), (-1,-1), 3),
        ("BOTTOMPADDING",(0,0), (-1,-1), 3),
        ("LEFTPADDING",  (0,0), (-1,-1), 8),
        ("BACKGROUND",   (0,0), (-1,-1), GRIS_LIGHT),
    ]))
    story.append(P("Type de déplacement :", sB))
    story.append(type_t)
    story.append(SP(3))

    # ── RENSEIGNEMENTS MISSIONNAIRE ────────────────────────────────────────────
    nom_complet = f"{mis.get('nom','').upper()}  {mis.get('prenom','')}"
    composante  = mis.get("composante", "")
    miss_data = [[
        P(f"NOM : {mis.get('nom','').upper()}", sB),
        P(f"Prénom : {mis.get('prenom','')}", sB),
        P(f"Composante / Service / Labo : {composante}", sB),
    ]]
    miss_t = Table(miss_data, colWidths=[CW*0.28, CW*0.28, CW*0.44])
    miss_t.setStyle(TableStyle([
        ("BOX",          (0,0), (-1,-1), 0.5, NOIR),
        ("INNERGRID",    (0,0), (-1,-1), 0.3, colors.Color(0.7,0.7,0.7)),
        ("BACKGROUND",   (0,0), (0,0),   GRIS_HDR),
        ("TOPPADDING",   (0,0), (-1,-1), 4),
        ("BOTTOMPADDING",(0,0), (-1,-1), 4),
        ("LEFTPADDING",  (0,0), (-1,-1), 5),
    ]))
    story.append(P("Renseignements concernant le missionnaire :", sB))
    story.append(miss_t)
    story.append(SP(3))

    # ── OBJET DE LA MISSION ────────────────────────────────────────────────────
    story.append(P(f"Déroulement de la mission / Objet : {msn.get('objet','')}", sB))
    story.append(SP(3))

    # ── TABLEAUX ALLER / RETOUR ────────────────────────────────────────────────
    def trajet_table(titre, date_dep, heure_dep, lieu_dep, date_arr, heure_arr, lieu_arr):
        hdr  = [[P(titre, sCB), P(""), P(""), P(""), P(""), P("")]]
        row1 = [[P("Départ :", sB), P(""), P(""), P("Arrivée :", sB), P(""), P("")]]
        row2 = [[P("Date :", sB), P("Heure :", sB), P("Lieu :", sB),
                 P("Date :", sB), P("Heure :", sB), P("Lieu :", sB)]]
        row3 = [[P(date_dep, sN), P(heure_dep, sN), P(lieu_dep, sN),
                 P(date_arr, sN), P(heure_arr, sN), P(lieu_arr, sN)]]
        data = hdr + row1 + row2 + row3
        cw   = [CW/6] * 6
        t    = Table(data, colWidths=cw)
        t.setStyle(TableStyle([
            ("BOX",          (0,0), (-1,-1), 0.5, NOIR),
            ("INNERGRID",    (0,0), (-1,-1), 0.3, colors.Color(0.7,0.7,0.7)),
            ("SPAN",         (0,0), (-1,0)),
            ("SPAN",         (0,1), (2,1)),
            ("SPAN",         (3,1), (5,1)),
            ("BACKGROUND",   (0,0), (-1,0),  GRIS_HDR),
            ("BACKGROUND",   (0,1), (-1,1),  GRIS_LIGHT),
            ("TOPPADDING",   (0,0), (-1,-1), 3),
            ("BOTTOMPADDING",(0,0), (-1,-1), 3),
            ("LEFTPADDING",  (0,0), (-1,-1), 4),
        ]))
        return t

    story.append(trajet_table(
        "ALLER :",
        fmt_date(msn.get("date_debut")),
        msn.get("heure_debut", ""),
        "Lille",
        fmt_date(msn.get("date_debut")),
        msn.get("heure_debut", ""),
        msn.get("destination", ""),
    ))
    story.append(SP(2))
    story.append(trajet_table(
        "RETOUR :",
        fmt_date(msn.get("date_fin")),
        msn.get("heure_fin", ""),
        msn.get("destination", ""),
        fmt_date(msn.get("date_fin")),
        msn.get("heure_fin", ""),
        "Lille",
    ))
    story.append(SP(3))

    # ── DÉPLACEMENT À L'ÉTRANGER (si applicable) ──────────────────────────────
    if etranger:
        niv = etr.get("niveau_securite", "vert")
        etr_data = [
            [P("* Niveau de sécurité (diplomatie.gouv.fr)", sOr)],
            [check(niv == "vert",   "Vigilance normale (vert)",                    sV, sN)],
            [check(niv == "jaune",  "Vigilance renforcée (jaune)",                 sJ, sN)],
            [check(niv == "orange", "Déconseillé sauf raison impérative (orange)", sOW, sN)],
            [check(niv == "rouge",  "Formellement déconseillé (rouge)",            sRg, sN)],
            [P("* Risques sanitaires (pasteur.fr)", sOr)],
            [P(f"Avis du Médecin : {etr.get('avis_medecin','')}", sOr)],
        ]
        story.append(section_box("Déplacement à l'étranger :", etr_data))
        story.append(SP(3))

    # ── MOYENS DE TRANSPORT ────────────────────────────────────────────────────
    transport_data = [
        [
            P("Moyens de transport", sB),
            P("Remboursement à l'agent", sCB),
            P("Billets par bon de commande", sCB),
            P("Pas de remboursement", sCB),
        ],
        [
            check(trp.get("train"), "Train"),
            check(trp.get("remboursement_direct") and trp.get("train"), "☑" if (trp.get("remboursement_direct") and trp.get("train")) else ""),
            check(trp.get("bon_commande") and trp.get("train"), ""),
            P(""),
        ],
        [
            check(trp.get("avion"), "Avion"),
            P(""),  P(""),  P(""),
        ],
        [
            check(trp.get("transports_commun"), "Transports en commun"),
            P(""),  P(""),  P(""),
        ],
        [
            P(f"Autres : {trp.get('autre', '__________________')}", sN),
            P(""),  P(""),  P(""),
        ],
    ]
    tr_t = Table(transport_data, colWidths=[CW*0.35, CW*0.25, CW*0.25, CW*0.15])
    tr_t.setStyle(TableStyle([
        ("BOX",          (0,0), (-1,-1), 0.5, NOIR),
        ("INNERGRID",    (0,0), (-1,-1), 0.3, colors.Color(0.7,0.7,0.7)),
        ("BACKGROUND",   (0,0), (-1,0),  GRIS_HDR),
        ("TOPPADDING",   (0,0), (-1,-1), 3),
        ("BOTTOMPADDING",(0,0), (-1,-1), 3),
        ("LEFTPADDING",  (0,0), (-1,-1), 5),
    ]))
    story.append(tr_t)
    story.append(SP(3))

    # ── VÉHICULE PERSONNEL ─────────────────────────────────────────────────────
    if trp.get("vehicule_personnel"):
        motif = trp.get("motif_vehicule", "")
        vp_data = [[
            check("convenance" in motif.lower(),  "Convenance personnelle"),
            check("commun" in motif.lower(),       "Absence transport en commun"),
            check("materiel" in motif.lower(),     "Transport matériel fragile/lourd"),
            check("temps" in motif.lower() or "économie" in motif.lower(), "Gain de temps / économie"),
        ]]
        vp_t = Table(vp_data, colWidths=[CW*0.28]*3 + [CW*0.16])
        vp_t.setStyle(TableStyle([
            ("BOX",  (0,0), (-1,-1), 0.5, NOIR),
            ("INNERGRID", (0,0), (-1,-1), 0.3, colors.Color(0.7,0.7,0.7)),
            ("TOPPADDING",   (0,0), (-1,-1), 3),
            ("BOTTOMPADDING",(0,0), (-1,-1), 3),
            ("LEFTPADDING",  (0,0), (-1,-1), 5),
        ]))
        story.append(P("Motif d'utilisation du véhicule personnel :", sB))
        story.append(vp_t)
        story.append(SP(3))

    # ── FRAIS DE SÉJOUR ────────────────────────────────────────────────────────
    nb_repas   = fra.get("nb_repas", "")
    nb_nuitees = fra.get("nb_nuitees", "")
    autres_dep = fra.get("autres_depenses", "")
    avance     = fra.get("avance", False)
    montant_av = fra.get("montant_avance", "")

    frais_data = [
        [P(f"Nombre de repas à rembourser : {nb_repas or '……'}    "
           f"Nombre de nuitées à rembourser : {nb_nuitees or '……'}", sN)],
        [P(f"Autres dépenses (nature et montant) : {autres_dep}", sN)],
        [P(f"Demande d'avance : {'OUI' if avance else 'NON'}    "
           f"Montant : {montant_av or '……'}", sB if avance else sN)],
    ]
    story.append(section_box("Frais de séjour (estimation) :", frais_data))
    story.append(SP(3))

    # ── IMPUTATION ─────────────────────────────────────────────────────────────
    cc  = fra.get("centre_cout", "……………")
    otp = fra.get("otp", "……………")
    dom = fra.get("domaine", "……………")
    story.append(P(
        f"Imputation des frais — Centre de coût : {cc}    "
        f"Élément d'OTP : {otp}    Domaine fonctionnel : {dom}", sB))
    story.append(SP(4))

    # ── SIGNATURES ────────────────────────────────────────────────────────────
    date_sig = fmt_date(msn.get("date_debut")) or "______________"
    sig_data = [
        [
            P("Signature du Missionnaire", sCB),
            P("Avis du Chef de service /\nResponsable de crédits", sCB),
            P("Décision du Doyen\nou Directeur", sCB),
        ],
        [
            P(f"À Lille, le {date_sig}\n\n\n\n", sN),
            P("☐ favorable    ☐ défavorable\n\nNom / Prénom / Signature :\n\n\n", sN),
            P("☐ autorisation accordée\n☐ autorisation refusée\n\nÀ Lille, le ______________\n", sN),
        ],
    ]
    sig_t = Table(sig_data, colWidths=[CW/3, CW/3, CW/3])
    sig_t.setStyle(TableStyle([
        ("BOX",          (0,0), (-1,-1), 0.5, NOIR),
        ("INNERGRID",    (0,0), (-1,-1), 0.5, NOIR),
        ("BACKGROUND",   (0,0), (-1,0),  GRIS_HDR),
        ("TOPPADDING",   (0,0), (-1,-1), 4),
        ("BOTTOMPADDING",(0,0), (-1,-1), 4),
        ("LEFTPADDING",  (0,0), (-1,-1), 6),
        ("VALIGN",       (0,0), (-1,-1), "TOP"),
    ]))
    story.append(sig_t)
    story.append(SP(2))
    story.append(P(
        "* La signature n'est pas obligatoire si l'OM est signé "
        "par le responsable des crédits.", sSub))

    # ── BUILD ──────────────────────────────────────────────────────────────────
    doc.build(story, onFirstPage=on_page, onLaterPages=on_page)
    return buf.getvalue()
