"""
generate_ordre_mission.py
Générateur PDF pour demande d'autorisation de déplacement
Compatible avec les modèles Université de Lille et GERiiCO
"""
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.units import mm
from reportlab.lib.styles import ParagraphStyle
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle,
    HRFlowable, Image
)
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.lib.enums import TA_LEFT, TA_CENTER, TA_RIGHT
import io, os
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
NOIR       = colors.Color(0.13, 0.12, 0.12)
ORANGE_ETR = colors.Color(0.60, 0.28, 0.02)
VERT       = colors.Color(0.0,  0.69, 0.31)
JAUNE      = colors.Color(0.80, 0.70, 0.0)
ORANGE_W   = colors.Color(0.89, 0.42, 0.04)
ROUGE      = colors.Color(1.0,  0.0,  0.0)
GRIS_HDR   = colors.Color(0.85, 0.85, 0.85)
GRIS_LIGHT = colors.Color(0.95, 0.95, 0.95)
BLANC      = colors.white

# ── Dimensions ─────────────────────────────────────────────────────────────────
W, H = A4
LM = RM = 18*mm
TM = BM = 15*mm
CW = W - LM - RM

# ── Styles ─────────────────────────────────────────────────────────────────────
def S(name, size=9, bold=False, italic=False,
      align=TA_LEFT, color=NOIR, sb=0, sa=1):
    fn = (FONT+"-B" if bold and not italic else
          FONT+"-I" if italic else FONT)
    return ParagraphStyle(name, fontName=fn, fontSize=size,
                          leading=size*1.35, alignment=align,
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
sV   = S("v",  bold=True, color=VERT, size=9)
sJ   = S("j",  bold=True, color=JAUNE, size=9)
sOW  = S("ow", bold=True, color=ORANGE_W, size=9)
sRg  = S("rg", bold=True, color=ROUGE, size=9)


def P(txt, style=None):
    txt = str(txt).replace("&","&amp;").replace("<","&lt;").replace(">","&gt;")
    return Paragraph(txt, style or sN)


def SP(h=3):
    return Spacer(1, h*mm)


def fmt_date(d):
    if not d: return ""
    try:
        return datetime.strptime(str(d)[:10], "%Y-%m-%d").strftime("%d/%m/%Y")
    except Exception:
        return str(d)


def check(val, label, s_on=None, s_off=None):
    mark  = "☑" if val else "☐"
    style = (s_on or sB) if val else (s_off or sN)
    return P(f"{mark}  {label}", style)


def on_page(canvas, doc):
    canvas.saveState()
    canvas.setFont(FONT, 7)
    canvas.setFillColor(colors.Color(0.5,0.5,0.5))
    canvas.drawCentredString(W/2, 10*mm, f"Page {canvas.getPageNumber()}")
    canvas.restoreState()


# ══════════════════════════════════════════════════════════════════════════════
def generate_ordre_mission_pdf(data: dict) -> bytes:
    """
    Génère le PDF de la demande d'autorisation de déplacement.

    data = {
      "template": "gerrico" | "ulille",   ← optionnel, défaut ulille
      "logo_path": "/chemin/logo.jpg",    ← optionnel
      "missionnaire": {
          "nom", "prenom", "composante"
      },
      "mission": {
          "objet", "destination", "pays",
          "date_debut", "heure_debut",
          "date_fin",   "heure_fin",
          "etranger": bool
      },
      "transport": {
          "type_deplacement": "avec_frais"|"partiel"|"sans_frais",
          "co_financeur": "",
          "train": bool, "avion": bool,
          "transports_commun": bool, "autre": "",
          "remboursement_direct": bool,
          "bon_commande": bool,
          "vehicule_personnel": bool,
          "motif_vehicule": "",
          "remboursement_km": bool,
          "remboursement_sncf": bool,
          "parking": bool,
          "peage": bool
      },
      "frais": {
          "nb_repas": 0, "nb_nuitees": 0,
          "autres_depenses": "",
          "avance": bool, "montant_avance": 0,
          "centre_cout": "", "otp": "", "domaine": ""
      },
      "etranger": {
          "niveau_securite": "vert"|"jaune"|"orange"|"rouge",
          "avis_medecin": ""
      }
    }
    """
    buf     = io.BytesIO()
    template = data.get("template", "ulille")
    doc = SimpleDocTemplate(
        buf, pagesize=A4,
        leftMargin=LM, rightMargin=RM,
        topMargin=TM, bottomMargin=BM+8*mm,
        title="Demande d'autorisation de déplacement",
    )

    mis = data.get("missionnaire", {})
    msn = data.get("mission",      {})
    trp = data.get("transport",    {})
    fra = data.get("frais",        {})
    etr = data.get("etranger",     {})
    etranger = msn.get("etranger", False)

    story = []

    # ── EN-TÊTE ────────────────────────────────────────────────────────────────
    # Logo si disponible
    logo_path = data.get("logo_path", "")
    if not logo_path:
        # Chercher logos connus dans le répertoire
        for candidate in [
            "/app/logo_gerrico.jpg",
            "/home/claude/logo_gerrico.jpg",
        ]:
            if os.path.exists(candidate):
                logo_path = candidate
                break

    titre_content = [
        P("DEMANDE D'AUTORISATION DE DÉPLACEMENT", sTT),
        P("Déplacement en France ou à l'étranger quelle que soit la durée", sT),
        P("(en dehors des congés réguliers)", sSub),
    ]

    if logo_path and os.path.exists(logo_path):
        try:
            logo = Image(logo_path, width=35*mm, height=20*mm, kind='proportional')
            header_data = [[logo, titre_content]]
            header_t = Table(header_data, colWidths=[40*mm, CW-40*mm])
            header_t.setStyle(TableStyle([
                ("VALIGN",       (0,0), (-1,-1), "MIDDLE"),
                ("LEFTPADDING",  (0,0), (-1,-1), 0),
                ("RIGHTPADDING", (0,0), (-1,-1), 0),
                ("TOPPADDING",   (0,0), (-1,-1), 0),
                ("BOTTOMPADDING",(0,0), (-1,-1), 0),
            ]))
            story.append(header_t)
        except Exception:
            for p in titre_content:
                story.append(p)
    else:
        for p in titre_content:
            story.append(p)

    story.append(SP(2))
    story.append(HRFlowable(width=CW, thickness=1.5, color=NOIR, spaceAfter=4))

    # ── TYPE DE DÉPLACEMENT ────────────────────────────────────────────────────
    type_depl   = trp.get("type_deplacement", "avec_frais")
    co_financeur = trp.get("co_financeur", "")

    type_data = [[
        check(type_depl == "avec_frais",
              "AVEC FRAIS pour l'Université de Lille"),
        check(type_depl == "partiel",   "PARTIEL"),
        check(type_depl == "sans_frais",
              "SANS FRAIS pour l'Université de Lille"),
    ]]
    type_t = Table(type_data, colWidths=[CW*0.40, CW*0.18, CW*0.42])
    type_t.setStyle(TableStyle([
        ("BOX",          (0,0), (-1,-1), 0.5, NOIR),
        ("INNERGRID",    (0,0), (-1,-1), 0.3, colors.Color(0.7,0.7,0.7)),
        ("BACKGROUND",   (0,0), (-1,-1), GRIS_LIGHT),
        ("TOPPADDING",   (0,0), (-1,-1), 3),
        ("BOTTOMPADDING",(0,0), (-1,-1), 3),
        ("LEFTPADDING",  (0,0), (-1,-1), 8),
    ]))
    story.append(P("Type de déplacement :", sB))
    story.append(type_t)

    if type_depl in ("partiel", "sans_frais") and co_financeur:
        story.append(SP(1))
        story.append(P(
            f"Organisme co-financeur / nature des frais : {co_financeur}", sN))
    story.append(SP(3))

    # ── MISSIONNAIRE ──────────────────────────────────────────────────────────
    composante = mis.get("composante", "Laboratoire GERiiCO (ULR 4073)")
    miss_data = [[
        P(f"NOM : {mis.get('nom','').upper()}", sB),
        P(f"Prénom : {mis.get('prenom','')}", sB),
        P(f"Composante / Service : {composante}", sB),
    ]]
    miss_t = Table(miss_data, colWidths=[CW*0.25, CW*0.25, CW*0.50])
    miss_t.setStyle(TableStyle([
        ("BOX",          (0,0), (-1,-1), 0.5, NOIR),
        ("INNERGRID",    (0,0), (-1,-1), 0.3, colors.Color(0.7,0.7,0.7)),
        ("BACKGROUND",   (0,0), (-1,-1), GRIS_HDR),
        ("TOPPADDING",   (0,0), (-1,-1), 4),
        ("BOTTOMPADDING",(0,0), (-1,-1), 4),
        ("LEFTPADDING",  (0,0), (-1,-1), 5),
    ]))
    story.append(P("Renseignements concernant le missionnaire :", sB))
    story.append(miss_t)
    story.append(SP(3))

    # ── MISSION ───────────────────────────────────────────────────────────────
    story.append(P("Renseignements concernant la mission :", sB))
    dates_data = [[
        P(f"Date de départ : {fmt_date(msn.get('date_debut'))}", sN),
        P(f"Heure de départ : {msn.get('heure_debut','')}", sN),
        P(f"Date de retour : {fmt_date(msn.get('date_fin'))}", sN),
        P(f"Heure de retour : {msn.get('heure_fin','')}", sN),
    ],[
        P(f"Lieu : {msn.get('destination','')}", sN),
        P(""),
        P(f"Pays : {msn.get('pays','France')}", sN),
        P(""),
    ]]
    dates_t = Table(dates_data, colWidths=[CW*0.28, CW*0.22, CW*0.28, CW*0.22])
    dates_t.setStyle(TableStyle([
        ("BOX",          (0,0), (-1,-1), 0.5, NOIR),
        ("INNERGRID",    (0,0), (-1,-1), 0.3, colors.Color(0.7,0.7,0.7)),
        ("TOPPADDING",   (0,0), (-1,-1), 3),
        ("BOTTOMPADDING",(0,0), (-1,-1), 3),
        ("LEFTPADDING",  (0,0), (-1,-1), 5),
    ]))
    story.append(dates_t)
    story.append(SP(2))
    story.append(P(f"OBJET : {msn.get('objet','')}", sB))
    story.append(SP(3))

    # ── SECTION ÉTRANGER ──────────────────────────────────────────────────────
    if etranger:
        niv = etr.get("niveau_securite", "vert")
        etr_rows = [
            [P("Niveau de sécurité (diplomatie.gouv.fr) :", sOr)],
            [check(niv=="vert",   "Vigilance normale (vert)",                    sV,  sN)],
            [check(niv=="jaune",  "Vigilance renforcée (jaune)",                 sJ,  sN)],
            [check(niv=="orange", "Déconseillé sauf raison impérative (orange)", sOW, sN)],
            [check(niv=="rouge",  "Formellement déconseillé (rouge)",            sRg, sN)],
            [P("Risques sanitaires (pasteur.fr) :", sOr)],
            [P("Avant tout déplacement à risque sanitaire, se rapprocher du Médecin"
               " de prévention 2 mois avant le départ.", sSub)],
            [P(f"Avis du Médecin : {etr.get('avis_medecin','FAVORABLE AU DEPART  /  DEFAVORABLE AU DEPART')}", sOr)],
        ]
        etr_t = Table([[P("Déplacement à l'étranger :", sB)]] + etr_rows,
                      colWidths=[CW])
        etr_t.setStyle(TableStyle([
            ("BOX",          (0,0), (-1,-1), 0.5, NOIR),
            ("INNERGRID",    (0,0), (-1,-1), 0.2, colors.Color(0.8,0.8,0.8)),
            ("BACKGROUND",   (0,0), (-1,0),  GRIS_HDR),
            ("TOPPADDING",   (0,0), (-1,-1), 2),
            ("BOTTOMPADDING",(0,0), (-1,-1), 2),
            ("LEFTPADDING",  (0,0), (-1,-1), 8),
            ("SPAN",         (0,0), (-1,0)),
        ]))
        story.append(etr_t)
        story.append(SP(3))

    # ── MOYENS DE TRANSPORT ───────────────────────────────────────────────────
    story.append(P("Moyens de transport :", sB))
    tr_data = [
        [P("Moyens de transport", sB),
         P("Remboursement à l'agent", sCB),
         P("Billets par bon de commande", sCB),
         P("Pas de remboursement", sCB)],
        [check(trp.get("train"), "Train"),
         P("☑" if trp.get("remboursement_direct") and trp.get("train") else "☐", sC),
         P("☑" if trp.get("bon_commande") and trp.get("train") else "☐", sC),
         P("☑" if not trp.get("remboursement_direct") and not trp.get("bon_commande") and trp.get("train") else "☐", sC)],
        [check(trp.get("avion"), "Avion"),
         P("☑" if trp.get("remboursement_direct") and trp.get("avion") else "☐", sC),
         P("☑" if trp.get("bon_commande") and trp.get("avion") else "☐", sC),
         P("", sC)],
        [check(trp.get("transports_commun"), "Transports en commun"),
         P("☑" if trp.get("remboursement_direct") and trp.get("transports_commun") else "☐", sC),
         P("", sC), P("", sC)],
        [P(f"Autres : {trp.get('autre','') or '_______________'}", sN),
         P(""), P(""), P("")],
    ]
    tr_t = Table(tr_data, colWidths=[CW*0.35, CW*0.23, CW*0.25, CW*0.17])
    tr_t.setStyle(TableStyle([
        ("BOX",          (0,0), (-1,-1), 0.5, NOIR),
        ("INNERGRID",    (0,0), (-1,-1), 0.3, colors.Color(0.7,0.7,0.7)),
        ("BACKGROUND",   (0,0), (-1,0),  GRIS_HDR),
        ("TOPPADDING",   (0,0), (-1,-1), 3),
        ("BOTTOMPADDING",(0,0), (-1,-1), 3),
        ("LEFTPADDING",  (0,0), (-1,-1), 5),
    ]))
    story.append(tr_t)
    story.append(SP(2))

    # ── VÉHICULE PERSONNEL ────────────────────────────────────────────────────
    if trp.get("vehicule_personnel"):
        motif = trp.get("motif_vehicule", "").lower()
        vp_motifs = [[
            check("convenance" in motif,  "Convenance personnelle"),
            check("commun"     in motif,  "Absence de transports en commun"),
            check("matériel"   in motif or "materiel" in motif,
                                          "Transport de matériel fragile/lourd"),
            check("temps"      in motif or "économie" in motif or "economie" in motif,
                                          "Gain de temps / économie"),
        ]]
        rembours = [[
            check(trp.get("remboursement_km"),   "Indemnité kilométrique"),
            check(trp.get("remboursement_sncf"),  "Tarif SNCF"),
            check(trp.get("parking"),             "Parking"),
            check(trp.get("peage"),               "Péage"),
        ]]
        vp_data = [
            [P("Motif utilisation véhicule personnel :", sB),
             P("Remboursements :", sB)],
            [Table(vp_motifs, colWidths=[CW*0.24]*4),
             Table(rembours,  colWidths=[CW*0.12]*4)],
        ]
        # Affichage simplifié
        vp_t = Table([
            [P("Motif d'utilisation du véhicule personnel :", sB),
             P("Remboursements :", sB)],
            [
                "\n".join([
                    ("☑" if "convenance" in motif else "☐") + "  Convenance personnelle",
                    ("☑" if "commun" in motif else "☐") + "  Absence transports en commun",
                    ("☑" if "matériel" in motif or "materiel" in motif else "☐") + "  Matériel fragile/lourd",
                    ("☑" if "temps" in motif or "économie" in motif else "☐") + "  Gain de temps/économie",
                ]),
                "\n".join([
                    ("☑" if trp.get("remboursement_km") else "☐") + "  Indemnité kilométrique",
                    ("☑" if trp.get("remboursement_sncf") else "☐") + "  Tarif SNCF",
                    ("☑" if trp.get("parking") else "☐") + "  Parking",
                    ("☑" if trp.get("peage") else "☐") + "  Péage",
                ]),
            ]
        ], colWidths=[CW*0.55, CW*0.45])
        # Reconstruire proprement
        vp_final = Table([
            [P("Motif d'utilisation du véhicule personnel :", sB),
             P("Remboursements :", sB)],
            [
                Table([[
                    check("convenance" in motif, "Convenance personnelle"),
                    check("commun" in motif,     "Absence transports en commun"),
                ],[
                    check("matériel" in motif or "materiel" in motif, "Matériel fragile"),
                    check("temps" in motif or "économie" in motif,    "Gain de temps"),
                ]], colWidths=[CW*0.28, CW*0.27]),
                Table([[
                    check(trp.get("remboursement_km"),   "Indemnité km"),
                    check(trp.get("remboursement_sncf"),  "Tarif SNCF"),
                ],[
                    check(trp.get("parking"), "Parking"),
                    check(trp.get("peage"),   "Péage"),
                ]], colWidths=[CW*0.22, CW*0.23]),
            ]
        ], colWidths=[CW*0.55, CW*0.45])
        vp_final.setStyle(TableStyle([
            ("BOX",          (0,0), (-1,-1), 0.5, NOIR),
            ("INNERGRID",    (0,0), (-1,-1), 0.3, colors.Color(0.7,0.7,0.7)),
            ("BACKGROUND",   (0,0), (-1,0),  GRIS_HDR),
            ("TOPPADDING",   (0,0), (-1,-1), 3),
            ("BOTTOMPADDING",(0,0), (-1,-1), 3),
            ("LEFTPADDING",  (0,0), (-1,-1), 5),
        ]))
        story.append(vp_final)
        story.append(SP(2))

    # ── FRAIS ─────────────────────────────────────────────────────────────────
    nb_r = fra.get("nb_repas",   "") or "______"
    nb_n = fra.get("nb_nuitees", "") or "______"
    autres = fra.get("autres_depenses", "") or "__________________________________"
    avance  = fra.get("avance", False)
    montant = fra.get("montant_avance", "") or "__________"

    frais_data = [
        [P("Frais de séjour (estimation) :", sB)],
        [P(f"Nb de repas à rembourser : {nb_r}     "
           f"Nb de nuitées à rembourser : {nb_n}", sN)],
        [P(f"Autres dépenses (nature et montant) : {autres}", sN)],
        [P(f"Demande d'avance : {'OUI' if avance else '__________'}     "
           f"Montant : {montant if avance else '__________'}", sB if avance else sN)],
    ]
    frais_t = Table(frais_data, colWidths=[CW])
    frais_t.setStyle(TableStyle([
        ("BOX",          (0,0), (-1,-1), 0.5, NOIR),
        ("INNERGRID",    (0,0), (-1,-1), 0.2, colors.Color(0.8,0.8,0.8)),
        ("BACKGROUND",   (0,0), (-1,0),  GRIS_HDR),
        ("TOPPADDING",   (0,0), (-1,-1), 3),
        ("BOTTOMPADDING",(0,0), (-1,-1), 3),
        ("LEFTPADDING",  (0,0), (-1,-1), 8),
        ("SPAN",         (0,0), (-1,0)),
    ]))
    story.append(frais_t)
    story.append(SP(2))

    # ── IMPUTATION ────────────────────────────────────────────────────────────
    cc  = fra.get("centre_cout", "__________")
    otp = fra.get("otp",         "__________")
    dom = fra.get("domaine",     "__________")
    story.append(P("Imputation des frais :", sB))
    story.append(P(
        f"Centre de coût : {cc}     "
        f"Élément d'OTP : {otp}     "
        f"Domaine fonctionnel : {dom}", sN))
    story.append(SP(4))

    # ── SIGNATURES ────────────────────────────────────────────────────────────
    date_sig = fmt_date(msn.get("date_debut")) or "______________"
    sig_data = [
        [P("Signature\ndu missionnaire", sCB),
         P("Avis Chef de service /\nResponsable de crédits", sCB),
         P("Avis Chef de service /\nResponsable de crédits", sCB),
         P("Avis Directeur\ndu CHRU", sCB)],
        [P(f"À Lille, le {date_sig}\n\n\n\n", sN),
         P("☐ favorable   ☐ défavorable\n\nNom / Prénom / Signature :\n\n\n", sN),
         P("☐ favorable   ☐ défavorable\n\nNom / Prénom / Signature :\n\n\n", sN),
         P("☐ favorable   ☐ défavorable\n\nSignature :\n\n\n", sN)],
        [P("Cadre réservé aux autorités compétentes\n(arrêté min. 3 juillet 2006)", sSub),
         P(""), P(""), P("")],
        [P("Décision Doyen / Directeur :\n"
           "☐ autorisation accordée\n"
           "☐ refusée : _______________\n"
           f"À Lille, le {date_sig}", sN),
         P("Décision Doyen / Directeur :\n"
           "☐ autorisation accordée\n"
           "☐ refusée : _______________\n"
           f"À Lille, le {date_sig}", sN),
         P("Décision Président Université :\n"
           "☐ autorisation accordée\n"
           "☐ refusée : _______________\n"
           f"À Lille, le {date_sig}", sN),
         P("Décision Président Université :\n"
           "☐ autorisation accordée\n"
           "☐ refusée : _______________\n"
           f"À Lille, le {date_sig}", sN)],
    ]
    sig_t = Table(sig_data, colWidths=[CW/4]*4)
    sig_t.setStyle(TableStyle([
        ("BOX",          (0,0), (-1,-1), 0.5, NOIR),
        ("INNERGRID",    (0,0), (-1,-1), 0.5, NOIR),
        ("BACKGROUND",   (0,0), (-1,0),  GRIS_HDR),
        ("SPAN",         (0,2), (-1,2)),
        ("TOPPADDING",   (0,0), (-1,-1), 3),
        ("BOTTOMPADDING",(0,0), (-1,-1), 3),
        ("LEFTPADDING",  (0,0), (-1,-1), 4),
        ("VALIGN",       (0,0), (-1,-1), "TOP"),
    ]))
    story.append(sig_t)
    story.append(SP(2))
    story.append(P(
        "* La signature du Chef de service n'est pas obligatoire "
        "si l'OM est signé par le responsable des crédits.", sSub))

    # ── BUILD ─────────────────────────────────────────────────────────────────
    doc.build(story, onFirstPage=on_page, onLaterPages=on_page)
    return buf.getvalue()
