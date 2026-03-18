"""
generate_annexes.py — v3
Génère pixel-perfect les Annexe 1 et Annexe 1bis AAP INSPÉ Lille HdF
Robuste aux textes longs et caractères spéciaux.
"""
import io, re
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.enums import TA_CENTER, TA_LEFT
from reportlab.lib.utils import ImageReader
from reportlab.platypus import (
    Paragraph, Spacer, Table, TableStyle, PageBreak
)
from reportlab.platypus.doctemplate import BaseDocTemplate, PageTemplate, Frame

# ─── Couleurs exactes ────────────────────────────────────────────────────────
GRAY_HEADER    = colors.Color(0.945, 0.945, 0.945)
BLACK          = colors.black
WHITE          = colors.white
COL_JAUNE      = colors.Color(1.0, 0.949, 0.8)
COL_VERT       = colors.Color(0.886, 0.937, 0.855)
COL_BLEU_CLAIR = colors.Color(0.851, 0.882, 0.949)
COL_TOTAL_ROW  = colors.Color(1.0, 0.80, 0.0)
COL_BLEU_NOTE  = colors.Color(0.0, 0.0, 0.502)

# ─── Dimensions A4 ────────────────────────────────────────────────────────────
PAGE_W, PAGE_H = A4
ML, MR, MT, MB = 65.5, 46.3, 56.0, 48.0
CW = PAGE_W - ML - MR   # 483.5 pts

# ─── Styles ───────────────────────────────────────────────────────────────────
def _s(name, font='Helvetica', size=10, align=TA_LEFT, color=BLACK,
       bold=False, italic=False, space_before=0, space_after=4, leading=None):
    fn = font
    if bold and italic: fn += '-BoldOblique'
    elif bold:          fn += '-Bold'
    elif italic:        fn += '-Oblique'
    return ParagraphStyle(name, fontName=fn, fontSize=size, alignment=align,
                          textColor=color, spaceBefore=space_before,
                          spaceAfter=space_after, leading=leading or size * 1.25)

S = {
    'title':    _s('title',    size=14, bold=True,   align=TA_CENTER, space_after=3),
    'h1':       _s('h1',       size=12, bold=True,   space_before=12, space_after=5),
    'sec_hdr':  _s('sec_hdr',  size=10, bold=True,   space_after=0),
    'body':     _s('body',     size=10, space_after=4, leading=13),
    'body_it':  _s('body_it',  size=10, italic=True, space_after=4, leading=13),
    'body_sm':  _s('body_sm',  size=8,  space_after=2, leading=10),
    'lbl':      _s('lbl',      size=9,  leading=11),
    'lbl_bold': _s('lbl_bold', size=9,  bold=True, leading=11),
    'val':      _s('val',      size=9,  leading=11, color=colors.Color(0.15, 0.15, 0.15)),
    'note_blue':_s('note_blue',size=7,  italic=True, color=COL_BLEU_NOTE, leading=9),
    'tbl_hdr':  _s('tbl_hdr',  size=7,  bold=True, leading=9, align=TA_CENTER),
    'tbl_cell': _s('tbl_cell', size=7.5, leading=10),
    'tbl_total':_s('tbl_total',size=8,  bold=True, leading=10),
}

# ─── Page template ────────────────────────────────────────────────────────────
class AAPDocTemplate(BaseDocTemplate):
    def __init__(self, buf, logo_bytes=None, **kw):
        self._logo = logo_bytes
        super().__init__(buf, **kw)
        frame = Frame(ML, MB, CW, PAGE_H - MT - MB - 70,
                      id='main', leftPadding=0, rightPadding=0,
                      topPadding=0, bottomPadding=0)
        self.addPageTemplates([PageTemplate(id='main', frames=[frame],
                                            onPage=self._header_footer)])

    def _header_footer(self, cv, doc):
        cv.saveState()
        if self._logo:
            logo_w, logo_h = 238, 51
            x = (PAGE_W - logo_w) / 2
            y = PAGE_H - MT - 8
            cv.drawImage(ImageReader(io.BytesIO(self._logo)), x, y,
                         width=logo_w, height=logo_h,
                         preserveAspectRatio=True, mask='auto')
        cv.setFont('Helvetica', 9)
        cv.setFillColor(colors.Color(0.4, 0.4, 0.4))
        cv.drawRightString(PAGE_W - MR, 26, str(doc.page))
        cv.restoreState()

# ─── Helpers texte sécurisé ───────────────────────────────────────────────────
def clean(text: str, max_chars: int = 4000) -> str:
    """Nettoie un texte pour ReportLab : échappe XML, retire balises orphelines."""
    if not text:
        return ''
    # Tronquer
    text = str(text)
    if len(text) > max_chars:
        text = text[:max_chars] + '…'
    # Échapper les caractères XML spéciaux (sauf si déjà échappés)
    text = text.replace('&', '&amp;')
    text = text.replace('<', '&lt;').replace('>', '&gt;')
    return text

def para(text: str, style, max_chars: int = 4000) -> Paragraph:
    """Crée un Paragraph robuste."""
    try:
        return Paragraph(clean(text, max_chars), style)
    except Exception:
        try:
            safe = re.sub(r'[^\x20-\x7E\n]', ' ', str(text))[:max_chars]
            return Paragraph(safe, style)
        except Exception:
            return Paragraph('(texte non affichable)', style)

def sp(h=6):
    return Spacer(1, h)

def section_box(text: str) -> Table:
    return Table(
        [[para(text, S['sec_hdr'])]],
        colWidths=[CW],
        style=TableStyle([
            ('BACKGROUND',    (0, 0), (-1, -1), GRAY_HEADER),
            ('BOX',           (0, 0), (-1, -1), 0.5, BLACK),
            ('TOPPADDING',    (0, 0), (-1, -1), 5),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 5),
            ('LEFTPADDING',   (0, 0), (-1, -1), 7),
        ])
    )

def form_table(rows: list, col1_w: float = 168) -> Table:
    col2_w = CW - col1_w
    data = []
    for lbl_raw, val in rows:
        bold = lbl_raw.startswith('**')
        lbl_txt = lbl_raw[2:] if bold else lbl_raw
        p_lbl = para(f'<b>{clean(lbl_txt)}</b>' if bold else clean(lbl_txt), S['lbl'])
        p_val = para(str(val or ''), S['val'], max_chars=2000)
        data.append([p_lbl, p_val])
    ts = TableStyle([
        ('BOX',           (0, 0), (-1, -1), 0.5, BLACK),
        ('INNERGRID',     (0, 0), (-1, -1), 0.3, BLACK),
        ('VALIGN',        (0, 0), (-1, -1), 'TOP'),
        ('TOPPADDING',    (0, 0), (-1, -1), 4),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 4),
        ('LEFTPADDING',   (0, 0), (-1, -1), 5),
        ('RIGHTPADDING',  (0, 0), (-1, -1), 5),
        ('BACKGROUND',    (0, 0), (0, -1), colors.Color(0.975, 0.975, 0.975)),
    ])
    return Table(data, colWidths=[col1_w, col2_w], style=ts)

def text_box(text: str, min_height: float = 60) -> Table:
    return Table(
        [[para(str(text or ''), S['body'], max_chars=4000)]],
        colWidths=[CW],
        style=TableStyle([
            ('BOX',           (0, 0), (-1, -1), 0.5, BLACK),
            ('TOPPADDING',    (0, 0), (-1, -1), 6),
            ('BOTTOMPADDING', (0, 0), (-1, -1), min_height),
            ('LEFTPADDING',   (0, 0), (-1, -1), 7),
            ('RIGHTPADDING',  (0, 0), (-1, -1), 7),
        ])
    )

def fmt_eur(v) -> str:
    try:
        f = float(v or 0)
        if f == 0:
            return '0,00 €'
        return f'{f:,.2f}'.replace(',', ' ').replace('.', ',') + ' €'
    except:
        return str(v)

# ═══════════════════════════════════════════════════════════════════════════════
# ANNEXE 1
# ═══════════════════════════════════════════════════════════════════════════════
def build_annexe1(project: dict, logo_bytes: bytes = None) -> bytes:
    buf = io.BytesIO()
    doc = AAPDocTemplate(buf, logo_bytes=logo_bytes, pagesize=A4,
                         topMargin=MT + 70, bottomMargin=MB,
                         leftMargin=ML, rightMargin=MR)
    story = []
    coord = project.get('coordinateur', {})

    # Titre
    story += [
        sp(10),
        para('<b>Annexe 1</b>', S['title']),
        para('<b>AAP 2026 pour projet 2027 et 2028</b>', S['title']),
        sp(6),
        para('<b>DOCUMENT DE CANDIDATURE</b>', S['title']),
        sp(14),
    ]

    # Section 1
    story.append(para('<b>1.  Identification précise du projet et de son porteur/ses porteurs</b>', S['h1']))
    story.append(sp(4))

    # A) Identité
    story.append(section_box('A) Identité du projet'))
    story.append(sp(4))
    mots = project.get('mots_cles', [])
    mots_str = ', '.join(mots) if isinstance(mots, list) else str(mots or '')
    story.append(form_table([
        ('Titre du projet:',           project.get('titre', '')),
        ('Acronyme (éventuellement):', project.get('acronyme', '')),
        ('Mots clefs (maximum 5):',    mots_str),
    ]))
    story.append(sp(10))

    # B) Résumé
    story.append(section_box('B) Résumé (10 lignes maximum)'))
    story.append(sp(4))
    story.append(text_box(project.get('resume', ''), min_height=90))
    story.append(para('<i>En cas de sélection, ce résumé sera publié sur le site internet de l\'INSPÉ Lille HdF dans la rubrique Recherche</i>', S['body_sm']))
    story.append(PageBreak())

    # C) Coordinateur
    story.append(section_box('C) Identification du porteur (coordinateur et unité de recherche) du projet'))
    story.append(sp(5))
    story.append(para('<i>Le coordinateur du projet doit être membre d\'une unité de recherche des universités régionales du périmètre de l\'Académie de Lille (UArtois, ULCO, ULille, UPHF).</i>', S['body_it']))
    story.append(sp(6))
    story.append(form_table([
        ('**Coordinateur\n(Nom, Prénom):',                coord.get('nom', '')),
        ('Titre/Grade:',                                   coord.get('grade', '')),
        ('Courriel:',                                      coord.get('email', '')),
        ('Téléphone:',                                     coord.get('telephone', '')),
        ('Institution de rattachement (Nom\net adresse):', coord.get('institution', '')),
        ('**UR de rattachement\n(identifiant EA UMR, nom et\nadresse):',
         f"{coord.get('ur_id', '')} – {coord.get('ur_nom', '')}".strip(' –')),
        ("Directeur de l'UR (Nom, prénom\net courriel de contact):", coord.get('ur_directeur', '')),
        ("Gestionnaire de l'UR (si\nApplicable - Nom, prénom,\ncourriel):", coord.get('ur_gestionnaire', '')),
        ("Tutelle de gestion de l'UR pour ce\nprojet (Nom et adresse):", coord.get('ur_tutelle', '')),
        ('**Autres membres de l\'UR impliqués\ndans le projet (Nom, prénom, titre,\ncourriel):', coord.get('membres', '')),
    ]))
    story.append(sp(10))

    # D) Partenaires
    story.append(section_box("D) Identification des autres partenaires (autant de fiches que de partenaires)"))
    story.append(sp(4))
    partenaires = project.get('partenaires', []) or [{}]
    for i, p in enumerate(partenaires):
        story.append(para(f'<i>Partenaire n° {i + 1}</i>', S['body_it']))
        story.append(sp(3))
        story.append(form_table([
            ('**Identité (Nom et statut)\nUnité de recherche, école,\nétablissement, …:', p.get('nom', '')),
            ('Coordonnées:',                p.get('coordonnees', '')),
            ("Expertise(s) du partenaire:", p.get('expertise', '')),
            ('**Contact partenaire porteur\n(Nom, prénom):', p.get('contact', '')),
            ('Titre / grade / fonction:',   p.get('titre_contact', '')),
            ('Courriel:',                   p.get('email_contact', '')),
            ('**Autres membres du partenaire impliqués\ndans le projet:', p.get('membres_partenaire', '')),
        ]))
        story.append(sp(5))
        story.append(para("<i>Préciser l'état actuel du partenariat :</i>", S['body_it']))
        story.append(text_box(p.get('etat_partenariat', ''), min_height=28))
        story.append(sp(8))
    story.append(PageBreak())

    # E) Publications
    story.append(section_box('E) Références équipe projet : porteur et partenaire(s)'))
    story.append(sp(4))
    story.append(para("5 publications maximum du porteur et de l'équipe projet dans le domaine :", S['body']))
    pub = str(project.get('publications', '') or '')
    story.append(text_box(pub, min_height=max(10, 60 - len(pub) // 10)))
    story.append(sp(12))

    # Section 2 : Description
    story.append(para('<b>2. Description du projet</b> (3 pages maximum)', S['h1']))
    story.append(para('<i>Faisant apparaître :</i>', S['body_it']))
    for hint in [
        "- Contexte (problématique, hypothèse de travail, résultats antérieurs…)",
        "- Enjeux généraux du projet (originalité, intérêt du sujet…)",
        "- Objectifs précis visés par le projet",
        "- Cadre(s) théorique(s), méthodologie(s) et sources employées",
        "- Résultats attendus",
        "- Modalités pratiques de restitution, de valorisation et de diffusion des résultats (notamment auprès des publics de l'INSPÉ Lille HdF) - livrables",
        "- Brève revue de la littérature existante",
    ]:
        story.append(para(hint, S['body']))
    story.append(sp(5))

    desc = str(project.get('description', '') or '')
    for line in desc.split('\n'):
        stripped = line.strip()
        if stripped.startswith('## '):
            story.append(para(f'<b>{clean(stripped[3:])}</b>', S['body']))
        elif stripped.startswith('# '):
            story.append(para(f'<b>{clean(stripped[2:])}</b>', S['h1']))
        elif stripped.startswith('• ') or stripped.startswith('* '):
            story.append(para(f'• {clean(stripped[2:])}', S['body']))
        elif stripped:
            story.append(para(stripped, S['body']))
        else:
            story.append(sp(3))
    story.append(sp(10))

    # Section 3 : Calendrier
    story.append(para('<b>3. Grandes étapes et calendrier prévisionnel</b> des tâches, livrables et jalons de réalisation du projet (1 page maximum)', S['h1']))

    cal = project.get('calendrier', [])
    if isinstance(cal, str):
        lines = [l for l in cal.split('\n') if '|' in l and not all(c in '|- ' for c in l)]
        cal = []
        for l in lines[1:]:
            cells = [c.strip() for c in l.split('|') if c.strip()]
            if len(cells) >= 4:
                cal.append({'etape': cells[0], 'debut': cells[1], 'fin': cells[2],
                            'duree': cells[3], 'livrables': cells[4] if len(cells) > 4 else ''})

    cal_rows = [[
        para('<b>Grandes étapes</b>', S['lbl_bold']),
        para('<b>Début\nprévisionnel</b>', S['lbl_bold']),
        para('<b>Fin\nprévisionnelle</b>', S['lbl_bold']),
        para('<b>Durée\nestimée</b>', S['lbl_bold']),
        para('<b>Livrables / jalons</b>', S['lbl_bold']),
    ]]
    for r in (cal or []):
        cal_rows.append([
            para(str(r.get('etape', '')), S['lbl']),
            para(str(r.get('debut', '')), S['lbl']),
            para(str(r.get('fin', '')),   S['lbl']),
            para(str(r.get('duree', '')), S['lbl']),
            para(str(r.get('livrables', '')), S['lbl'], max_chars=500),
        ])
    while len(cal_rows) < 9:
        cal_rows.append([para('', S['lbl'])] * 5)

    story.append(Table(cal_rows, colWidths=[148, 62, 62, 56, 155.5], style=TableStyle([
        ('BOX',           (0, 0), (-1, -1), 0.5, BLACK),
        ('INNERGRID',     (0, 0), (-1, -1), 0.3, BLACK),
        ('BACKGROUND',    (0, 0), (-1, 0),  GRAY_HEADER),
        ('TOPPADDING',    (0, 0), (-1, -1), 3),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 3),
        ('LEFTPADDING',   (0, 0), (-1, -1), 5),
        ('VALIGN',        (0, 0), (-1, -1), 'TOP'),
        ('FONTSIZE',      (0, 0), (-1, -1), 9),
    ])))
    story.append(PageBreak())

    # Section 4 : Budget
    story.append(para('<b>4. Budget prévisionnel et demande financière</b>', S['h1']))
    story.append(para("Voir Annexe 1bis (tableau joint).", S['body']))
    story.append(sp(8))
    for fin_label, fin_key in [
        ("Ce projet bénéficie-t-il d'un autre financement déjà obtenu ?", 'financement_existant'),
        ("Ce projet fait-il l'objet d'une autre demande de financement en cours ?", 'financement_cours'),
        ("Ce projet va-t-il faire l'objet d'une demande de financement à venir ?", 'financement_avenir'),
    ]:
        story.append(para(fin_label, S['body']))
        fin = project.get(fin_key) or {}
        story.append(form_table([
            ('Type de financement',    fin.get('type', '')),
            ('Nom du financeur',       fin.get('financeur', '')),
            ('Dispositif de financement', fin.get('dispositif', '')),
            ('Montant',                fin.get('montant', '')),
            ('Dépenses éligibles',     fin.get('eligibles', '')),
        ]))
        story.append(sp(6))
    story.append(PageBreak())

    # Établissement
    story.append(section_box("C) Établissement qui assurera la gestion financière du projet"))
    story.append(sp(5))
    etab = project.get('etablissement') or {}
    rib  = project.get('rib') or {}
    story.append(form_table([
        ('Etablissement :',   etab.get('nom', '')),
        ('Agent Comptable :', etab.get('agent_comptable', '')),
        ('Adresse:',          etab.get('adresse', '')),
        ('Tél :',             etab.get('tel', '')),
        ('Mél :',             etab.get('mel', '')),
    ]))
    story.append(sp(6))
    story.append(form_table([
        ('Responsable du suivi financier', etab.get('resp_financier', '')),
        ('Adresse:',                       etab.get('adresse_resp', '')),
        ('Tél :',                          etab.get('tel_resp', '')),
        ('Mél :',                          etab.get('mel_resp', '')),
    ]))
    story.append(sp(5))
    tva = 'Oui' if project.get('assujetti_tva') else 'Non'
    story.append(para(f"L'établissement gestionnaire est-il assujetti à la TVA ?  <b>{tva}</b>", S['body']))
    story.append(sp(6))
    story.append(form_table([
        ('Identification bancaire', ''),
        ('Banque',                  rib.get('banque', '')),
        ('Titulaire du compte',     rib.get('titulaire', '')),
        ('Domiciliation',           rib.get('domiciliation', '')),
        ('N° de compte',            rib.get('compte', '')),
        ('Code banque',             rib.get('code_banque', '')),
        ('Clé RIB',                 rib.get('cle_rib', '')),
        ('Code Guichet',            rib.get('code_guichet', '')),
    ]))
    story.append(sp(12))

    # Section 5 : Appui
    story.append(para("<b>5. Demande complémentaire d'appui ou d'accompagnement par l'INSPÉ Lille HdF</b>", S['h1']))
    story.append(text_box(project.get('demande_appui', ''), min_height=60))
    story.append(sp(16))

    # Signatures
    story.append(Table(
        [[para(' Date et lieu\n\n\n Nom et Signature* du porteur de projet', S['lbl']),
          para(" Date et lieu\n\n\n Nom et Signature* du directeur d'unité de recherche", S['lbl'])]],
        colWidths=[CW / 2, CW / 2],
        style=TableStyle([
            ('BOX',           (0, 0), (-1, -1), 0.5, BLACK),
            ('INNERGRID',     (0, 0), (-1, -1), 0.5, BLACK),
            ('TOPPADDING',    (0, 0), (-1, -1), 6),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 52),
            ('LEFTPADDING',   (0, 0), (-1, -1), 6),
        ])
    ))
    story.append(sp(5))
    story.append(para("*En l'absence de ces 2 signatures le projet ne sera pas évalué", S['body_sm']))

    doc.build(story)
    return buf.getvalue()


# ═══════════════════════════════════════════════════════════════════════════════
# ANNEXE 1BIS
# ═══════════════════════════════════════════════════════════════════════════════
def build_annexe1bis(project: dict, budget_lines: list, logo_bytes: bytes = None) -> bytes:
    buf = io.BytesIO()
    doc = AAPDocTemplate(buf, logo_bytes=logo_bytes, pagesize=A4,
                         topMargin=MT + 70, bottomMargin=MB,
                         leftMargin=ML, rightMargin=MR)
    story = []

    story += [
        sp(10),
        para('<b>AAP 2026 pour projet 2027 et 2028 - Annexe 1 bis : tableau des dépenses</b>', S['title']),
        sp(8),
        para("<b>A) Budget du projet et demande financière à l'INSPÉ Lille HdF</b>", S['body']),
        sp(4),
        para(f"<b>Titre du projet:</b>  {clean(project.get('titre', ''))}", S['body']),
        sp(3),
        para(f"<b>Nom et prénom du porteur :</b>  {clean(project.get('porteur', ''))}", S['body']),
        sp(8),
    ]

    # Bloc notes
    story.append(Table(
        [[para('<i>Pour la colonne "détail" du tableau</i>', S['note_blue'])],
         [para("- nous vous invitons à vous rapprocher de la/du gestionnaire administratif et financier de votre unité de recherche pour connaître les taux horaire de vacations, les montants de gratification de stage… en vigueur lors de la constitution du dossier.", S['note_blue'])],
         [para("- pour l'organisation des journées d'étude : merci de préciser le format envisagé (journée ou demi-journée) et les besoins : accueil café, repas (nombre), prises en charge…", S['note_blue'])],
        ],
        colWidths=[CW],
        style=TableStyle([
            ('BOX',         (0, 0), (-1, -1), 0.3, COL_BLEU_NOTE),
            ('BACKGROUND',  (0, 0), (-1, -1), colors.Color(0.94, 0.96, 1.0)),
            ('TOPPADDING',  (0, 0), (-1, -1), 3),
            ('BOTTOMPADDING',(0,0), (-1, -1), 3),
            ('LEFTPADDING', (0, 0), (-1, -1), 6),
        ])
    ))
    story.append(sp(8))

    # Colonnes budget
    cw = [88, 103, 40, 36, 33, 36, 33, 40, 74.5]

    def h(txt): return para(f'<b>{txt}</b>', S['tbl_hdr'])
    def c(txt): return para(str(txt or ''), S['tbl_cell'])
    def cb(txt): return para(f'<b>{str(txt or "")}</b>', S['tbl_total'])

    row1 = [
        h('NATURE DE LA DEPENSE\n(missions, prestations, frais de\nréception, vacations, gratification\nde stage, organisation séminaire,\ncolloque, journées d\'étude,\nmatériel consommable utile au\nprojet de recherche, ...)'),
        h('DETAIL\n(type de prestation, nombre\nd\'entretiens et d\'heures\nà retranscrire…)'),
        h('TOTAL\nTTC\n(en\neuros)'),
        h("Aide demandée à l'INSPÉ Lille HdF TTC (en euros)"),
        '', '', '', '',
        h('Co-\nfinancement,\nle cas\néchéant,\nTTC\n(en euros)'),
    ]
    row2 = ['', '', '', h('Dont montant TTC en euros pour 2027'), '', h('Dont montant TTC en euros pour 2028'), '', h('Total'), '']
    row3 = ['', '', '', h('Fonct.'), h('RH'), h('Fonct.'), h('RH'), '', '']

    bdata = [row1, row2, row3]

    totals = {'total': 0, 'f27': 0, 'r27': 0, 'f28': 0, 'r28': 0, 'ins': 0, 'cof': 0}

    for line in (budget_lines or []):
        t   = float(line.get('total', 0) or 0)
        f27 = float(line.get('fonct2027', 0) or 0)
        r27 = float(line.get('rh2027', 0) or 0)
        f28 = float(line.get('fonct2028', 0) or 0)
        r28 = float(line.get('rh2028', 0) or 0)
        cof = float(line.get('cofinancement', 0) or 0)
        ins = f27 + r27 + f28 + r28
        totals['total'] += t; totals['f27'] += f27; totals['r27'] += r27
        totals['f28'] += f28; totals['r28'] += r28; totals['ins'] += ins; totals['cof'] += cof
        bdata.append([
            c(line.get('nature', '')),
            c(line.get('detail', '')),
            c(fmt_eur(t)),
            c(fmt_eur(f27) if f27 else ''),
            c(fmt_eur(r27) if r27 else ''),
            c(fmt_eur(f28) if f28 else ''),
            c(fmt_eur(r28) if r28 else ''),
            c(fmt_eur(ins)),
            c(fmt_eur(cof) if cof else ''),
        ])

    while len(bdata) < 23:
        bdata.append([c(''), c(''), c('0,00 €'), c(''), c(''), c(''), c(''), c('0,00 €'), c('')])

    bdata.append([
        cb('Total'),
        para("<i>Tableau présenté à l'équilibre</i>", S['note_blue']),
        cb(fmt_eur(totals['total'])),
        cb(fmt_eur(totals['f27'])),
        cb(fmt_eur(totals['r27'])),
        cb(fmt_eur(totals['f28'])),
        cb(fmt_eur(totals['r28'])),
        cb(fmt_eur(totals['ins'])),
        cb(fmt_eur(totals['cof'])),
    ])

    N = len(bdata)
    HDR = 3
    DATA_END = N - 2
    TOT_ROW = N - 1

    ts = TableStyle([
        ('BOX',        (0, 0), (-1, -1), 0.5, BLACK),
        ('INNERGRID',  (0, 0), (-1, -1), 0.25, BLACK),
        ('SPAN',       (3, 0), (7, 0)),
        ('SPAN',       (3, 1), (4, 1)),
        ('SPAN',       (5, 1), (6, 1)),
        ('BACKGROUND', (0, 0), (-1, HDR - 1), GRAY_HEADER),
        ('BACKGROUND', (3, HDR), (3, DATA_END), COL_JAUNE),
        ('BACKGROUND', (4, HDR), (4, DATA_END), COL_VERT),
        ('BACKGROUND', (5, HDR), (5, DATA_END), COL_JAUNE),
        ('BACKGROUND', (6, HDR), (6, DATA_END), COL_VERT),
        ('BACKGROUND', (7, HDR), (7, DATA_END), COL_BLEU_CLAIR),
        ('BACKGROUND', (0, TOT_ROW), (-1, TOT_ROW), COL_TOTAL_ROW),
        ('VALIGN',     (0, 0), (-1, -1), 'MIDDLE'),
        ('ALIGN',      (2, 0), (-1, -1), 'CENTER'),
        ('TOPPADDING', (0, 0), (-1, -1), 2),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 2),
        ('LEFTPADDING',   (0, 0), (-1, -1), 3),
        ('RIGHTPADDING',  (0, 0), (-1, -1), 3),
        ('FONTSIZE',      (0, 0), (-1, -1), 7),
    ])

    story.append(Table(bdata, colWidths=cw, style=ts, repeatRows=3))
    doc.build(story)
    return buf.getvalue()


def load_logo(path: str) -> bytes:
    with open(path, 'rb') as f:
        return f.read()
