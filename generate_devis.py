"""
generate_devis.py — Générateur PDF de devis Ludoscience
Style fidèle au modèle DevisISFECBretagneE01.docx
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
    pdfmetrics.registerFontFamily("Cal", normal="Cal", bold="Cal-B", italic="Cal-I")
    FONT = "Cal"
except Exception:
    FONT = "Helvetica"

# ── Couleurs ───────────────────────────────────────────────────────────────────
ORANGE      = colors.Color(1.0, 0.4, 0.0)      # FF6600 — couleur Ludoscience
GRIS_HEADER = colors.Color(0.92, 0.92, 0.92)   # en-tête tableau
GRIS_TOTAL  = colors.Color(0.97, 0.97, 0.97)   # ligne total
NOIR        = colors.black
BLANC       = colors.white

# ── Dimensions ─────────────────────────────────────────────────────────────────
W, H = A4
LM = RM = TM = BM = 25 * mm
CW = W - LM - RM


# ── Styles ─────────────────────────────────────────────────────────────────────
def S(name, size=11, bold=False, italic=False, align=TA_LEFT,
      color=NOIR, sb=0, sa=2):
    fn = (FONT + "-B" if bold and not italic else
          FONT + "-I" if italic else FONT)
    return ParagraphStyle(name, fontName=fn, fontSize=size,
                          leading=size * 1.4, alignment=align,
                          spaceBefore=sb, spaceAfter=sa, textColor=color)

sNormal  = S("normal")
sBold    = S("bold", bold=True)
sOrange  = S("orange", bold=True, color=ORANGE, size=13)
sCenter  = S("center", align=TA_CENTER)
sCenterB = S("centerB", bold=True, align=TA_CENTER)
sRight   = S("right", align=TA_RIGHT)
sSmall   = S("small", size=9)
sSmallI  = S("smallI", size=9, italic=True)
sSmallR  = S("smallR", size=9, align=TA_RIGHT)


def P(txt, style=None):
    txt = str(txt).replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
    return Paragraph(txt.strip(), style or sNormal)


def SP(h=4):
    return Spacer(1, h * mm)


def fmt_eur(val):
    try:
        return f"{float(val):,.2f} €".replace(",", " ").replace(".", ",")
    except Exception:
        return str(val)


def fmt_date(d):
    if not d:
        return ""
    try:
        return datetime.strptime(str(d)[:10], "%Y-%m-%d").strftime("%d/%m/%Y")
    except Exception:
        return str(d)


def on_page(canvas, doc):
    canvas.saveState()
    canvas.setFont(FONT, 8)
    canvas.setFillColor(colors.Color(0.6, 0.6, 0.6))
    canvas.drawCentredString(W / 2, 12 * mm,
                             f"Page {canvas.getPageNumber()}")
    canvas.restoreState()


# ══════════════════════════════════════════════════════════════════════════════
# FONCTION PRINCIPALE
# ══════════════════════════════════════════════════════════════════════════════

def generate_devis_pdf(data: dict) -> bytes:
    """
    Génère le PDF du devis.

    Structure attendue de `data` :
    {
      "emetteur": { nom, adresse, email, tel, siret },
      "client":   { nom, adresse, email },
      "devis":    { numero, intitule, date_emission, ville_emission,
                    date_prestation, lieu },
      "lignes":   [{ description, quantite, prix_unitaire, total }],
      "totaux":   { montant_ht, tva_taux, montant_tva, montant_ttc,
                    mention_tva, exonere }
    }
    """
    buf = io.BytesIO()
    doc = SimpleDocTemplate(
        buf, pagesize=A4,
        leftMargin=LM, rightMargin=RM,
        topMargin=TM, bottomMargin=BM + 8 * mm,
        title=f"Devis {data.get('devis', {}).get('numero', '')}",
    )

    em  = data.get("emetteur", {})
    cl  = data.get("client", {})
    dv  = data.get("devis", {})
    lg  = data.get("lignes", [])
    tot = data.get("totaux", {})

    story = []

    # ── BLOC CLIENT + INTITULÉ + NUMÉRO ──────────────────────────────────────
    # Tableau 2 colonnes : client à gauche, infos devis à droite
    client_lines = [P(cl.get("nom", ""), sBold)]
    for line in cl.get("adresse", "").split("\n"):
        if line.strip():
            client_lines.append(P(line, sNormal))
    if cl.get("email"):
        client_lines.append(P(cl["email"], sNormal))

    right_lines = [
        P(dv.get("intitule", ""), sBold),
        SP(2),
        P(f"Facture n° {dv.get('numero', '')}", sBold),
    ]

    header_data = [[client_lines, right_lines]]
    header_table = Table(header_data, colWidths=[CW * 0.5, CW * 0.5])
    header_table.setStyle(TableStyle([
        ("VALIGN",       (0, 0), (-1, -1), "TOP"),
        ("LEFTPADDING",  (0, 0), (-1, -1), 0),
        ("RIGHTPADDING", (0, 0), (-1, -1), 0),
        ("TOPPADDING",   (0, 0), (-1, -1), 0),
        ("BOTTOMPADDING",(0, 0), (-1, -1), 0),
    ]))
    story.append(header_table)
    story.append(SP(4))

    # ── VILLE ET DATE ─────────────────────────────────────────────────────────
    ville = dv.get("ville_emission", "Lille")
    date  = fmt_date(dv.get("date_emission")) or datetime.now().strftime("%d/%m/%Y")
    story.append(P(f"{ville}, le {date}"))
    story.append(SP(6))

    # ── LIGNE ORANGE SÉPARATRICE ───────────────────────────────────────────────
    story.append(HRFlowable(width=CW, thickness=2, color=ORANGE, spaceAfter=6))

    # ── TABLEAU DES PRESTATIONS ───────────────────────────────────────────────
    # En-tête
    col_w = [CW * 0.55, CW * 0.12, CW * 0.18, CW * 0.15]

    table_data = [
        [
            P("Description", sBold),
            P("Quantité", sBold),
            P("Prix Unitaire", sBold),
            P("TOTAL", sBold),
        ]
    ]

    # Lignes de prestation
    for ligne in lg:
        desc = ligne.get("description", "")
        qte  = str(ligne.get("quantite", ""))
        pu   = ligne.get("prix_unitaire", "")
        tot_ligne = ligne.get("total", 0)

        # Formater le total
        if isinstance(tot_ligne, (int, float)) and tot_ligne != 0:
            tot_str = f"+{fmt_eur(tot_ligne)}" if tot_ligne > 0 else fmt_eur(tot_ligne)
        else:
            tot_str = str(tot_ligne)

        table_data.append([
            P(desc, sNormal),
            P(qte, sCenter),
            P(str(pu), sCenter),
            P(tot_str, sRight),
        ])

    # Ligne TVA
    mention_tva = tot.get("mention_tva", "Montant non assujetti à la TVA")
    exonere     = tot.get("exonere", True)

    if exonere:
        table_data.append([
            P(mention_tva, sSmallI),
            P("", sSmall), P("", sSmall),
            P(fmt_eur(tot.get("montant_ht", 0)), sSmallR),
        ])
    else:
        tva_taux = tot.get("tva_taux", 20)
        table_data.append([
            P(f"Montant HT", sSmall),
            P("", sSmall), P("", sSmall),
            P(fmt_eur(tot.get("montant_ht", 0)), sSmallR),
        ])
        table_data.append([
            P(f"TVA {tva_taux}%", sSmall),
            P("", sSmall), P("", sSmall),
            P(fmt_eur(tot.get("montant_tva", 0)), sSmallR),
        ])

    # Ligne total net
    table_data.append([
        P("Montant total net", sBold),
        P("", sNormal), P("", sNormal),
        P(f" {fmt_eur(tot.get('montant_ttc', 0))}", sBold),
    ])

    n_lignes    = len(lg)
    n_rows      = len(table_data)
    idx_tva_start = 1 + n_lignes
    idx_total     = n_rows - 1

    t = Table(table_data, colWidths=col_w, repeatRows=1)

    style_cmds = [
        # En-tête
        ("BACKGROUND",    (0, 0), (-1, 0),  GRIS_HEADER),
        ("LINEBELOW",     (0, 0), (-1, 0),  0.5, NOIR),
        ("TOPPADDING",    (0, 0), (-1, -1), 4),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
        ("LEFTPADDING",   (0, 0), (-1, -1), 6),
        ("RIGHTPADDING",  (0, 0), (-1, -1), 6),
        ("VALIGN",        (0, 0), (-1, -1), "TOP"),
        # Bordures externes
        ("BOX",           (0, 0), (-1, -1), 0.5, NOIR),
        ("INNERGRID",     (0, 0), (-1, n_lignes), 0.3, colors.Color(0.8, 0.8, 0.8)),
        # Fusion colonnes TVA
        ("SPAN",          (0, idx_tva_start), (2, idx_tva_start)),
        # Total — fond gris léger
        ("BACKGROUND",    (0, idx_total), (-1, idx_total), GRIS_TOTAL),
        ("LINEABOVE",     (0, idx_total), (-1, idx_total), 0.5, NOIR),
        ("SPAN",          (0, idx_total), (2, idx_total)),
    ]

    # Si non exonéré, fusionner aussi la ligne HT et TVA
    if not exonere and n_rows > idx_tva_start + 2:
        for row_i in range(idx_tva_start, idx_total):
            style_cmds.append(("SPAN", (0, row_i), (2, row_i)))

    t.setStyle(TableStyle(style_cmds))
    story.append(t)

    # ── AFFAIRE SUIVIE PAR ────────────────────────────────────────────────────
    story.append(SP(8))
    story.append(HRFlowable(width=CW, thickness=0.5,
                             color=colors.Color(0.7, 0.7, 0.7), spaceAfter=4))

    contact_parts = ["Affaire suivie par : "]
    if em.get("nom"):
        contact_parts.append(em["nom"])
    if em.get("tel"):
        contact_parts.append(f"  Tél : {em['tel']}")
    if em.get("email"):
        contact_parts.append(f"  Email : {em['email']}")

    story.append(P("   ".join(contact_parts), sSmall))

    if em.get("siret"):
        story.append(P(f"SIRET : {em['siret']}", sSmall))

    # ── BUILD ──────────────────────────────────────────────────────────────────
    doc.build(story, onFirstPage=on_page, onLaterPages=on_page)
    return buf.getvalue()
