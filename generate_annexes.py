"""
generate_annexes.py — v2
Génère pixel-perfect les Annexe 1 et Annexe 1bis AAP INSPÉ Lille HdF
Usage: import et appel de build_annexe1() et build_annexe1bis()
"""
import io, base64
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.units import mm
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.enums import TA_CENTER, TA_LEFT
from reportlab.lib.utils import ImageReader
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle,
    PageBreak, KeepTogether
)
from reportlab.platypus.doctemplate import BaseDocTemplate, PageTemplate, Frame

# ─── Couleurs exactes extraites des PDFs officiels ────────────────────────────
GRAY_HEADER    = colors.Color(0.945, 0.945, 0.945)
BLACK          = colors.black
WHITE          = colors.white
COL_JAUNE      = colors.Color(1.0, 0.949, 0.8)        # fonctionnement 2027/2028
COL_VERT       = colors.Color(0.886, 0.937, 0.855)    # RH 2027/2028
COL_BLEU_CLAIR = colors.Color(0.851, 0.882, 0.949)    # total INSPÉ
COL_TOTAL_ROW  = colors.Color(1.0, 0.80, 0.0)         # ligne total jaune
COL_BLEU_NOTE  = colors.Color(0.0, 0.0, 0.502)        # texte bleu notes

# ─── Dimensions A4 (pts) ─────────────────────────────────────────────────────
PAGE_W, PAGE_H = A4   # 595.32 x 841.92
ML   = 65.5
MR   = 46.3
MT   = 56.0
MB   = 48.0
CW   = PAGE_W - ML - MR   # 483.5 pts

# ─── Styles typographiques ────────────────────────────────────────────────────
def _s(name, font='Helvetica', size=10, align=TA_LEFT, color=BLACK,
       bold=False, italic=False, space_before=0, space_after=4, leading=None):
    fn = font
    if bold and italic: fn += '-BoldOblique'
    elif bold:          fn += '-Bold'
    elif italic:        fn += '-Oblique'
    return ParagraphStyle(name, fontName=fn, fontSize=size, alignment=align,
                          textColor=color, spaceBefore=space_before,
                          spaceAfter=space_after, leading=leading or size*1.25)

S = {
    'title':      _s('title',     size=14, bold=True,   align=TA_CENTER, space_after=3),
    'h1':         _s('h1',        size=12, bold=True,   space_before=12, space_after=5),
    'sec_hdr':    _s('sec_hdr',   size=10, bold=True,   space_after=0),
    'body':       _s('body',      size=10, space_after=4, leading=13),
    'body_it':    _s('body_it',   size=10, italic=True, space_after=4, leading=13),
    'body_sm':    _s('body_sm',   size=8,  space_after=2, leading=10),
    'lbl':        _s('lbl',       size=9,  leading=11),
    'lbl_bold':   _s('lbl_bold',  size=9,  bold=True, leading=11),
    'val':        _s('val',       size=9,  leading=11, color=colors.Color(0.15,0.15,0.15)),
    'note_blue':  _s('note_blue', size=7,  italic=True, color=COL_BLEU_NOTE, leading=9),
    'tbl_hdr':    _s('tbl_hdr',   size=7,  bold=True, leading=9, align=TA_CENTER),
    'tbl_cell':   _s('tbl_cell',  size=7.5, leading=10),
    'tbl_total':  _s('tbl_total', size=8,  bold=True, leading=10),
    'footer':     _s('footer',    size=9,  align=TA_LEFT, color=colors.Color(0.4,0.4,0.4)),
}

# ─── Page template avec logo INSPÉ ───────────────────────────────────────────

class AAPDocTemplate(BaseDocTemplate):
    """Template de page avec logo centré + numéro de page bas droit."""
    def __init__(self, buf, logo_bytes=None, **kw):
        self._logo = logo_bytes
        super().__init__(buf, **kw)
        frame = Frame(ML, MB, CW, PAGE_H - MT - MB - 70,
                      id='main', leftPadding=0, rightPadding=0,
                      topPadding=0, bottomPadding=0)
        self.addPageTemplates([PageTemplate(id='main', frames=[frame], onPage=self._header_footer)])

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

# ─── Helpers ──────────────────────────────────────────────────────────────────

def sp(h=6):
    return Spacer(1, h)

def section_box(html: str) -> Table:
    """Bande grise header de section."""
    return Table(
        [[Paragraph(html, S['sec_hdr'])]],
        colWidths=[CW],
        style=TableStyle([
            ('BACKGROUND',    (0,0), (-1,-1), GRAY_HEADER),
            ('BOX',           (0,0), (-1,-1), 0.5, BLACK),
            ('TOPPADDING',    (0,0), (-1,-1), 5),
            ('BOTTOMPADDING', (0,0), (-1,-1), 5),
            ('LEFTPADDING',   (0,0), (-1,-1), 7),
        ])
    )

def form_table(rows: list, col1_w: float = 168) -> Table:
    """
    Tableau formulaire 2 colonnes (label | valeur).
    rows = list of (label_html, value_str)
    Préfixer le label avec '**' pour le mettre en gras.
    """
    col2_w = CW - col1_w
    data = []
    bold_rows = []
    for i, (lbl_raw, val) in enumerate(rows):
        bold = lbl_raw.startswith('**')
        lbl_txt = lbl_raw[2:] if bold else lbl_raw
        if bold:
            bold_rows.append(i)
        p_lbl = Paragraph(f'<b>{lbl_txt}</b>' if bold else lbl_txt, S['lbl'])
        p_val = Paragraph(val or '', S['val'])
        data.append([p_lbl, p_val])

    ts = TableStyle([
        ('BOX',           (0,0), (-1,-1), 0.5, BLACK),
        ('INNERGRID',     (0,0), (-1,-1), 0.3, BLACK),
        ('VALIGN',        (0,0), (-1,-1), 'TOP'),
        ('TOPPADDING',    (0,0), (-1,-1), 4),
        ('BOTTOMPADDING', (0,0), (-1,-1), 4),
        ('LEFTPADDING',   (0,0), (-1,-1), 5),
        ('RIGHTPADDING',  (0,0), (-1,-1), 5),
        ('BACKGROUND',    (0,0), (0,-1), colors.Color(0.975, 0.975, 0.975)),
    ])
    return Table(data, colWidths=[col1_w, col2_w], style=ts)

def text_box(text: str, min_height: float = 60) -> Table:
    """Zone de texte libre avec bordure."""
    return Table(
        [[Paragraph(text or '', S['body'])]],
        colWidths=[CW],
        style=TableStyle([
            ('BOX',           (0,0), (-1,-1), 0.5, BLACK),
            ('TOPPADDING',    (0,0), (-1,-1), 6),
            ('BOTTOMPADDING', (0,0), (-1,-1), min_height),
            ('LEFTPADDING',   (0,0), (-1,-1), 7),
            ('RIGHTPADDING',  (0,0), (-1,-1), 7),
        ])
    )

def fmt_eur(v) -> str:
    """Formate un nombre en euros français."""
    try:
        f = float(v or 0)
        if f == 0:
            return '0,00 €'
        s = f'{f:,.2f}'.replace(',', ' ').replace('.', ',')
        return s + ' €'
    except:
        return str(v)

# ═══════════════════════════════════════════════════════════════════════════════
# ANNEXE 1 — Document de Candidature
# ═══════════════════════════════════════════════════════════════════════════════

def build_annexe1(project: dict, logo_bytes: bytes = None) -> bytes:
    """
    Génère l'Annexe 1 au format PDF officiel INSPÉ.
    
    project dict keys:
      titre, acronyme, mots_cles (list|str), resume,
      coordinateur: {nom, grade, email, telephone, institution,
                     ur_id, ur_nom, ur_directeur, ur_gestionnaire,
                     ur_tutelle, membres},
      partenaires: list of {nom, coordonnees, expertise, contact,
                            titre_contact, email_contact,
                            membres_partenaire, etat_partenariat},
      publications: str,
      description: str  (peut contenir ## pour les sous-titres),
      calendrier: list of {etape, debut, fin, duree, livrables},
      financement_existant, financement_cours, financement_avenir:
          {type, financeur, dispositif, montant, eligibles},
      etablissement: {nom, agent_comptable, adresse, tel, mel,
                      resp_financier, adresse_resp, tel_resp, mel_resp},
      rib: {banque, titulaire, domiciliation, compte,
            code_banque, cle_rib, code_guichet},
      assujetti_tva: bool,
      demande_appui: str,
    """
    buf = io.BytesIO()
    doc = AAPDocTemplate(
        buf, logo_bytes=logo_bytes,
        pagesize=A4, topMargin=MT + 70, bottomMargin=MB,
        leftMargin=ML, rightMargin=MR,
    )
    story = []
    coord = project.get('coordinateur', {})

    # ── Titre ──────────────────────────────────────────────────────────────────
    story += [
        sp(10),
        Paragraph('<b>Annexe 1</b>', S['title']),
        Paragraph('<b>AAP 2026 pour projet 2027 et 2028</b>', S['title']),
        sp(6),
        Paragraph('<b>DOCUMENT DE CANDIDATURE</b>', S['title']),
        sp(14),
    ]

    # ── 1. Identification ──────────────────────────────────────────────────────
    story.append(Paragraph('<b>1.  Identification précise du projet et de son porteur/ses porteurs</b>', S['h1']))
    story.append(sp(4))

    # A) Identité
    story.append(section_box('A) Identité du projet'))
    story.append(sp(4))
    mots = project.get('mots_cles', [])
    mots_str = ', '.join(mots) if isinstance(mots, list) else str(mots)
    story.append(form_table([
        ('Titre du projet:',           project.get('titre', '')),
        ('Acronyme (éventuellement):', project.get('acronyme', '')),
        ('Mots clefs (maximum 5):',    mots_str),
    ]))
    story.append(sp(10))

    # B) Résumé
    story.append(section_box('B) Résumé <i>(10 lignes maximum)</i>'))
    story.append(sp(4))
    story.append(text_box(project.get('resume', ''), min_height=90))
    story.append(Paragraph(
        '<i>En cas de sélection, ce résumé sera publié sur le site internet de '
        "l'INSPÉ Lille HdF dans la rubrique Recherche</i>", S['body_sm']))
    story.append(PageBreak())

    # C) Coordinateur
    story.append(section_box(
        'C) Identification du porteur <i>(coordinateur et unité de recherche)</i> <b>du projet</b>'))
    story.append(sp(5))
    story.append(Paragraph(
        "<i>Le coordinateur du projet doit être membre d'une unité de recherche des "
        "universités régionales du périmètre de l'Académie de Lille "
        '(UArtois, ULCO, ULille, UPHF).</i>', S['body_it']))
    story.append(sp(6))
    story.append(form_table([
        ('**Coordinateur\n(Nom, Prénom):',                  coord.get('nom', '')),
        ('Titre/Grade:',                                     coord.get('grade', '')),
        ('Courriel:',                                        coord.get('email', '')),
        ('Téléphone:',                                       coord.get('telephone', '')),
        ('Institution de rattachement (Nom\net adresse):',   coord.get('institution', '')),
        ('**UR de rattachement\n(identifiant EA UMR, nom et\nadresse):',
         f"{coord.get('ur_id','')} – {coord.get('ur_nom','')}".strip(' –')),
        ("Directeur de l'UR (Nom, prénom\net courriel de contact):",
         coord.get('ur_directeur', '')),
        ('Gestionnaire de l\'UR (si\nApplicable - Nom, prénom,\ncourriel):',
         coord.get('ur_gestionnaire', '')),
        ('Tutelle de gestion de l\'UR pour ce\nprojet (Nom et adresse):',
         coord.get('ur_tutelle', '')),
        ('**Autres membres de l\'UR impliqués\ndans le projet (Nom, prénom, titre,\ncourriel):',
         coord.get('membres', '')),
    ]))
    story.append(sp(10))

    # D) Partenaires
    story.append(section_box(
        'D) Identification des autres partenaires <i>(autant de fiches que de partenaires)</i>'))
    story.append(sp(4))
    for i, p in enumerate(project.get('partenaires', [{}])):
        story.append(Paragraph(f'<i>Partenaire n° {i+1}</i>', S['body_it']))
        story.append(sp(3))
        story.append(form_table([
            ('**Identité (Nom et statut)\nUnité de recherche, école,\nétablissement, délégation ou\n'
             'service académique, association,\nentreprise, … :',   p.get('nom', '')),
            ('Coordonnées:',                                          p.get('coordonnees', '')),
            ("Expertise(s) du partenaire:",                          p.get('expertise', '')),
            ('**Contact partenaire porteur\n(Nom, prénom):',         p.get('contact', '')),
            ('Titre / grade / fonction:',                            p.get('titre_contact', '')),
            ('Courriel:',                                            p.get('email_contact', '')),
            ('**Autres membres du partenaire impliqués\ndans le projet\n(Nom, prénom, titre, courriel):',
             p.get('membres_partenaire', '')),
        ]))
        story.append(sp(5))
        story.append(Paragraph(
            "<i>Préciser l'état actuel du partenariat (existant, en cours de construction, à construire…) :</i>",
            S['body_it']))
        story.append(text_box(p.get('etat_partenariat', ''), min_height=28))
        story.append(sp(8))
    story.append(PageBreak())

    # E) Publications + Description + Calendrier
    story.append(section_box('E) Références équipe projet : porteur et partenaire(s)'))
    story.append(sp(4))
    story.append(Paragraph(
        "5 publications maximum du porteur et de l'équipe projet dans le domaine "
        '(liste et liens le cas échéant) :', S['body']))
    pub = project.get('publications', '')
    story.append(text_box(pub, min_height=max(10, 80 - len(pub) // 6)))
    story.append(sp(12))

    # Section 2 : Description
    story.append(Paragraph('<b>2. Description du projet</b> (3 pages maximum)', S['h1']))
    story.append(Paragraph('<i>Faisant apparaître :</i>', S['body_it']))
    for hint in [
        '- Contexte (problématique, hypothèse de travail, résultats antérieurs…)',
        '- Enjeux généraux du projet (originalité, intérêt du sujet…)',
        '- Objectifs précis visés par le projet',
        "- Cadre(s) théorique(s), méthodologie(s) et sources employées",
        '- Résultats attendus',
        "- Modalités pratiques de restitution, de valorisation et de diffusion des résultats "
        "(notamment auprès des publics de l'INSPÉ Lille HdF) - livrables",
        '- Brève revue de la littérature existante',
    ]:
        story.append(Paragraph(hint, S['body']))
    story.append(sp(5))

    desc = project.get('description', '')
    for line in desc.split('\n'):
        stripped = line.strip()
        if stripped.startswith('## '):
            story.append(Paragraph(f'<b>{stripped[3:]}</b>', S['body']))
        elif stripped.startswith('# '):
            story.append(Paragraph(f'<b>{stripped[2:]}</b>', S['h1']))
        elif stripped:
            story.append(Paragraph(stripped, S['body']))
        else:
            story.append(sp(3))
    story.append(sp(10))

    # Section 3 : Calendrier
    story.append(Paragraph(
        '<b>3. Grandes étapes et calendrier prévisionnel</b> des tâches, livrables '
        'et jalons de réalisation du projet (1 page maximum)', S['h1']))

    cal = project.get('calendrier', [])
    # Parse texte si nécessaire
    if isinstance(cal, str):
        lines = [l for l in cal.split('\n') if '|' in l
                 and not all(c in '|- ' for c in l)]
        cal = []
        for l in lines[1:]:  # skip header row
            cells = [c.strip() for c in l.split('|') if c.strip()]
            if len(cells) >= 4:
                cal.append({'etape': cells[0], 'debut': cells[1],
                            'fin': cells[2], 'duree': cells[3],
                            'livrables': cells[4] if len(cells) > 4 else ''})

    cal_hdr = [
        Paragraph('<b>Grandes étapes</b>', S['lbl_bold']),
        Paragraph('<b>Début\nprévisionnel</b>', S['lbl_bold']),
        Paragraph('<b>Fin\nprévisionnelle</b>', S['lbl_bold']),
        Paragraph('<b>Durée\nestimée</b>', S['lbl_bold']),
        Paragraph('', S['lbl_bold']),  # placeholder for livrables header
    ]
    cal_rows = [cal_hdr]
    for r in cal:
        cal_rows.append([
            Paragraph(r.get('etape', ''),    S['lbl']),
            Paragraph(r.get('debut', ''),    S['lbl']),
            Paragraph(r.get('fin', ''),      S['lbl']),
            Paragraph(r.get('duree', ''),    S['lbl']),
            Paragraph(r.get('livrables', ''), S['lbl']),
        ])
    while len(cal_rows) < 9:
        cal_rows.append([Paragraph('', S['lbl'])] * 5)

    cal_cw = [148, 62, 62, 56, 155.5]
    story.append(Table(cal_rows, colWidths=cal_cw, style=TableStyle([
        ('BOX',           (0,0), (-1,-1), 0.5, BLACK),
        ('INNERGRID',     (0,0), (-1,-1), 0.3, BLACK),
        ('BACKGROUND',    (0,0), (-1,0),  GRAY_HEADER),
        ('TOPPADDING',    (0,0), (-1,-1), 3),
        ('BOTTOMPADDING', (0,0), (-1,-1), 3),
        ('LEFTPADDING',   (0,0), (-1,-1), 5),
        ('VALIGN',        (0,0), (-1,-1), 'TOP'),
        ('FONTSIZE',      (0,0), (-1,-1), 9),
    ])))
    story.append(PageBreak())

    # Section 4 : Budget (renvoi + co-financements)
    story.append(Paragraph('<b>4. Budget prévisionnel et demande financière</b>', S['h1']))
    story.append(Paragraph(
        "L'aide financière accordée aux projets sélectionnés sera versée à l'unité de recherche "
        "porteuse du projet. Voir Annexe 1bis jointe.", S['body']))
    story.append(sp(8))

    for fin_label, fin_key in [
        ("Ce projet bénéficie-t-il d'un autre financement déjà obtenu ?", 'financement_existant'),
        ("Ce projet fait-il l'objet d'une autre demande de financement en cours ?", 'financement_cours'),
        ("Ce projet va-t-il faire l'objet d'une demande de financement à venir ?", 'financement_avenir'),
    ]:
        story.append(Paragraph(fin_label, S['body']))
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

    # Établissement gestionnaire
    story.append(section_box("C) Établissement qui assurera la gestion financière du projet"))
    story.append(sp(5))
    etab = project.get('etablissement') or {}
    rib  = project.get('rib') or {}
    story.append(form_table([
        ('Etablissement :',    etab.get('nom', '')),
        ('Agent Comptable :',  etab.get('agent_comptable', '')),
        ('Adresse:',           etab.get('adresse', '')),
        ('Tél :',              etab.get('tel', '')),
        ('Mél :',              etab.get('mel', '')),
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
    story.append(Paragraph(
        f"L'établissement gestionnaire est-il assujetti à la TVA ?  <b>{tva}</b>", S['body']))
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
    story.append(Paragraph(
        "<b>5. Demande complémentaire d'appui ou d'accompagnement par l'INSPÉ Lille HdF</b>",
        S['h1']))
    story.append(text_box(project.get('demande_appui', ''), min_height=60))
    story.append(sp(16))

    # Signatures
    story.append(Table(
        [[Paragraph(' Date et lieu\n\n\n Nom et Signature* du porteur de projet', S['lbl']),
          Paragraph(" Date et lieu\n\n\n Nom et Signature* du directeur d'unité de recherche", S['lbl'])]],
        colWidths=[CW/2, CW/2],
        style=TableStyle([
            ('BOX',           (0,0), (-1,-1), 0.5, BLACK),
            ('INNERGRID',     (0,0), (-1,-1), 0.5, BLACK),
            ('TOPPADDING',    (0,0), (-1,-1), 6),
            ('BOTTOMPADDING', (0,0), (-1,-1), 52),
            ('LEFTPADDING',   (0,0), (-1,-1), 6),
        ])
    ))
    story.append(sp(5))
    story.append(Paragraph(
        "*En l'absence de ces 2 signatures le projet ne sera pas évalué", S['body_sm']))

    doc.build(story)
    return buf.getvalue()


# ═══════════════════════════════════════════════════════════════════════════════
# ANNEXE 1BIS — Tableau des dépenses
# ═══════════════════════════════════════════════════════════════════════════════

def build_annexe1bis(project: dict, budget_lines: list, logo_bytes: bytes = None) -> bytes:
    """
    Génère l'Annexe 1bis (tableau des dépenses) au format officiel.

    project keys: titre, porteur, montant_inspe
    budget_lines: list of dicts with:
      nature, detail, total, fonct2027, rh2027, fonct2028, rh2028, cofinancement
    """
    buf = io.BytesIO()
    doc = AAPDocTemplate(
        buf, logo_bytes=logo_bytes,
        pagesize=A4, topMargin=MT + 70, bottomMargin=MB,
        leftMargin=ML, rightMargin=MR,
    )
    story = []

    story += [
        sp(10),
        Paragraph(
            '<b>AAP 2026 pour projet 2027 et 2028 - Annexe 1 bis : tableau des dépenses</b>',
            S['title']),
        sp(8),
        Paragraph('<b>A) Budget du projet et demande financière à l\'INSPÉ Lille HdF</b>', S['body']),
        sp(4),
        Paragraph(f'<b>Titre du projet:</b>  {project.get("titre", "")}', S['body']),
        sp(3),
        Paragraph(f'<b>Nom et prénom du porteur :</b>  {project.get("porteur", "")}', S['body']),
        sp(8),
    ]

    # Bloc notes bleutées
    notes = [
        ('<i>Pour la colonne "détail" du tableau</i>', True),
        ('- nous vous invitons à vous rapprocher de la/du gestionnaire administratif et financier '
         'de votre unité de recherche pour connaître les taux horaire de vacations, les montants '
         'de gratification de stage… en vigueur lors de la constitution du dossier et de mentionner '
         "l'ensemble des informations (exemple : X heures d'entretiens réalisées par un intervenant "
         'au taux horaire de X soit X € au total)', False),
        ("- pour l'organisation des journées d'étude : merci de préciser le format envisagé "
         '(journée ou demi-journée) et les besoins : accueil café, repas (nombre), prises en charge '
         '(hôtel, transport en fonction de la provenance des intervenants)…', False),
    ]
    notes_data = [[Paragraph(t, S['note_blue'])] for t, _ in notes]
    story.append(Table(notes_data, colWidths=[CW], style=TableStyle([
        ('BOX',           (0,0), (-1,-1), 0.3, COL_BLEU_NOTE),
        ('BACKGROUND',    (0,0), (-1,-1), colors.Color(0.94, 0.96, 1.0)),
        ('TOPPADDING',    (0,0), (-1,-1), 3),
        ('BOTTOMPADDING', (0,0), (-1,-1), 3),
        ('LEFTPADDING',   (0,0), (-1,-1), 6),
    ])))
    story.append(sp(8))

    # ── Tableau budgétaire ────────────────────────────────────────────────────
    # Largeurs colonnes (total = CW = 483.5)
    # [Nature | Détail | Total | F27 | RH27 | F28 | RH28 | TotINSPÉ | Cofin]
    cw = [88, 103, 40, 36, 33, 36, 33, 40, 74.5]

    def h(txt):
        return Paragraph(f'<b>{txt}</b>', S['tbl_hdr'])
    def c(txt):
        return Paragraph(txt or '', S['tbl_cell'])
    def cb(txt):
        return Paragraph(f'<b>{txt}</b>', S['tbl_total'])

    # Ligne header 1 (avec spans)
    row1 = [
        h('NATURE DE LA DEPENSE\n(missions, prestations, frais de réception,\nvacations, gratification de stage,\norganisation de séminaire, colloque,\njournées d\'étude, matériel consommable\nutile au projet de recherche, ...)'),
        h('DETAIL\n(type de prestation, nombre\nd\'entretiens et d\'heures\nà retranscrire…)'),
        h('TOTAL\nTTC\n(en\neuros)'),
        h('Aide demandée à l\'INSPÉ Lille HdF TTC (en euros)'),
        '', '', '', '',
        h('Co-\nfinancement,\nle cas\néchéant,\nTTC\n(en euros)'),
    ]
    # Ligne header 2
    row2 = ['', '',
        '',
        h('Dont montant TTC en euros pour 2027'),
        '',
        h('Dont montant TTC en euros pour 2028'),
        '',
        h('Total'),
        '',
    ]
    # Ligne header 3
    row3 = ['', '', '',
        h('Fonct.'), h('RH'),
        h('Fonct.'), h('RH'),
        '', '',
    ]

    bdata = [row1, row2, row3]

    totals = {'total': 0, 'f27': 0, 'r27': 0, 'f28': 0, 'r28': 0, 'ins': 0, 'cof': 0}

    for line in budget_lines:
        t   = float(line.get('total', 0) or 0)
        f27 = float(line.get('fonct2027', 0) or 0)
        r27 = float(line.get('rh2027', 0) or 0)
        f28 = float(line.get('fonct2028', 0) or 0)
        r28 = float(line.get('rh2028', 0) or 0)
        cof = float(line.get('cofinancement', 0) or 0)
        ins = f27 + r27 + f28 + r28
        totals['total'] += t; totals['f27'] += f27; totals['r27'] += r27
        totals['f28'] += f28; totals['r28'] += r28; totals['ins'] += ins
        totals['cof'] += cof
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

    # Compléter jusqu'à 20 lignes de données
    while len(bdata) < 23:
        bdata.append([c(''), c(''), c('0,00 €'), c(''), c(''), c(''), c(''), c('0,00 €'), c('')])

    # Ligne total
    bdata.append([
        cb('Total'),
        Paragraph('<i>Tableau présenté à l\'équilibre</i>', S['note_blue']),
        cb(fmt_eur(totals['total'])),
        cb(fmt_eur(totals['f27'])),
        cb(fmt_eur(totals['r27'])),
        cb(fmt_eur(totals['f28'])),
        cb(fmt_eur(totals['r28'])),
        cb(fmt_eur(totals['ins'])),
        cb(fmt_eur(totals['cof'])),
    ])

    N = len(bdata)
    HDR = 3   # 3 rows headers
    DATA_START = HDR
    DATA_END   = N - 2   # last data row (before total)
    TOT_ROW    = N - 1

    ts = TableStyle([
        ('BOX',           (0,0), (-1,-1), 0.5, BLACK),
        ('INNERGRID',     (0,0), (-1,-1), 0.25, BLACK),
        # Header spans row 0
        ('SPAN',          (3,0), (7,0)),   # "Aide demandée INSPÉ" spans cols 3-7
        # Header spans row 1
        ('SPAN',          (3,1), (4,1)),   # "2027"
        ('SPAN',          (5,1), (6,1)),   # "2028"
        # Header background
        ('BACKGROUND',    (0,0), (-1,HDR-1), GRAY_HEADER),
        # Column colors (data rows only)
        ('BACKGROUND',    (3,DATA_START), (3,DATA_END), COL_JAUNE),
        ('BACKGROUND',    (4,DATA_START), (4,DATA_END), COL_VERT),
        ('BACKGROUND',    (5,DATA_START), (5,DATA_END), COL_JAUNE),
        ('BACKGROUND',    (6,DATA_START), (6,DATA_END), COL_VERT),
        ('BACKGROUND',    (7,DATA_START), (7,DATA_END), COL_BLEU_CLAIR),
        # Total row
        ('BACKGROUND',    (0,TOT_ROW), (-1,TOT_ROW), COL_TOTAL_ROW),
        # Layout
        ('VALIGN',        (0,0), (-1,-1), 'MIDDLE'),
        ('ALIGN', (2,0), (-1,-1), 'CENTER'),
        ('TOPPADDING',    (0,0), (-1,-1), 2),
        ('BOTTOMPADDING', (0,0), (-1,-1), 2),
        ('LEFTPADDING',   (0,0), (-1,-1), 3),
        ('RIGHTPADDING',  (0,0), (-1,-1), 3),
        ('FONTSIZE',      (0,0), (-1,-1), 7),
    ])

    story.append(Table(bdata, colWidths=cw, style=ts, repeatRows=3))
    doc.build(story)
    return buf.getvalue()


# ─── Utility : encode logo from file path ─────────────────────────────────────

def load_logo(path: str) -> bytes:
    with open(path, 'rb') as f:
        return f.read()


# ─── Test ─────────────────────────────────────────────────────────────────────
if __name__ == '__main__':
    logo = load_logo('/tmp/logo_Image1.png')

    sample_project = {
        'titre': "IAG et mémoires de master MEEF : usages, représentations et pratiques d'encadrement à l'INSPE Lille HdF",
        'acronyme': 'MémorIA-MASTER',
        'mots_cles': ["IA générative", "mémoire de master", "MEEF", "formation enseignants", "encadrement"],
        'resume': "Ce projet vise à analyser les usages réels et les représentations de l'intelligence artificielle générative (IAG) chez les étudiants de master MEEF lors de la rédaction de leur mémoire de recherche, ainsi que les pratiques d'encadrement des directeurs/directrices de mémoire face à ces nouveaux outils. S'inscrivant dans l'Axe 2 de l'AAP INSPÉ (Apprendre et enseigner dans et hors l'école), cette recherche adopte une méthodologie mixte articulant questionnaire à grande échelle et entretiens semi-directifs. Les résultats permettront de produire un guide de ressources pour les formateurs et d'alimenter les réflexions sur l'éthique de l'IAG dans la formation universitaire.",
        'coordinateur': {
            'nom': 'Matuszak, Céline',
            'grade': "MCF en Sciences de l'Information et de la Communication",
            'email': 'celine.matuszak@univ-lille.fr',
            'telephone': 'À compléter',
            'institution': "Université de Lille – INSPÉ Lille HdF, 2 rue du Recteur Céleste, 59000 Lille",
            'ur_id': 'EA 4073',
            'ur_nom': 'GERiiCO',
            'ur_directeur': 'À compléter – direction.geriico@univ-lille.fr',
            'ur_gestionnaire': 'À compléter',
            'ur_tutelle': 'Université de Lille',
            'membres': 'Kergosien, Eric (MCF, SIC) – eric.kergosien@univ-lille.fr\nde la Broise, Patrice (PU, SIC) – patrice.delabroise@univ-lille.fr',
        },
        'partenaires': [],
        'publications': "1. Matuszak, C. (2023). Écriture académique et outils numériques. Revue des Sciences de l'Éducation, 49(2).\n2. Kergosien, E., & Matuszak, C. (2022). Annotation sémantique et corpus académiques. Document Numérique, 25(1).\n3. de la Broise, P. (2021). Communication institutionnelle et formation. Communication & Organisation, 60.\n4. Kergosien, E. (2023). TAL et analyse de discours pédagogiques. ISKO International.\n5. de la Broise, P., & Matuszak, C. (2024). L'IA générative en contexte de formation. Spirale, 73.",
        'description': """## 1. Contexte et problématique
L'essor des outils d'IA générative (IAG) comme ChatGPT transforme profondément les pratiques d'écriture académique. Dans le contexte de la formation des enseignants via les masters MEEF, les mémoires de recherche constituent un exercice central, à la fois formateur et évaluatif. Or les chartes et pratiques d'encadrement face à l'IAG restent très hétérogènes à l'INSPÉ Lille HdF.

## 2. Hypothèses de travail
Nous faisons l'hypothèse que les usages de l'IAG par les étudiants MEEF sont hétérogènes et peu encadrés, et que les directeurs/directrices de mémoire n'ont pas encore élaboré de posture professionnelle stabilisée face à ces outils.

## 3. Objectifs précis
- Cartographier les usages déclarés de l'IAG par les étudiants MEEF
- Analyser leurs représentations de l'intégrité académique liées à ces usages
- Identifier les pratiques d'encadrement émergentes des directeurs de mémoire
- Produire des ressources opérationnelles pour les formateurs INSPÉ

## 4. Cadres théoriques et méthodologie
Approche mixte : questionnaire en ligne (n≈300 étudiants master MEEF INSPÉ Lille HdF) + 20 entretiens semi-directifs (étudiants et encadrants). Cadres : écriture académique (Rinck, 2011), intégrité académique (Lancaster & Cotarlan, 2021), IA et formation (Ouyang & Jiao, 2021).

## 5. Résultats attendus et livrables
Rapport de recherche, guide pratique pour formateurs, article dans revue spécialisée, séminaire de restitution ouvert à la communauté INSPÉ Lille HdF, bilan scientifique et financier.

## 6. Modalités de valorisation
Séminaire de restitution organisé à l'INSPÉ Lille HdF (2028), ouvert aux formateurs, enseignants-chercheurs et acteurs éducatifs. Guide mis en ligne sur le site INSPÉ. Communication en colloque (AREF, JOCAIR).

## 7. Revue de littérature
Rinck (2011) ; Lancaster & Cotarlan (2021) ; Ouyang & Jiao (2021) ; Amiel & Reeves (2008) ; Mollick & Mollick (2023).""",
        'calendrier': [
            {'etape': 'Phase préparatoire', 'debut': 'Janv. 2027', 'fin': 'Mars 2027', 'duree': '3 mois', 'livrables': 'Protocole validé, grilles entretien'},
            {'etape': 'Enquête quantitative (questionnaire)', 'debut': 'Avr. 2027', 'fin': 'Juin 2027', 'duree': '3 mois', 'livrables': 'Base de données questionnaire'},
            {'etape': 'Analyse quantitative', 'debut': 'Juil. 2027', 'fin': 'Sept. 2027', 'duree': '3 mois', 'livrables': 'Rapport intermédiaire'},
            {'etape': 'Entretiens semi-directifs (n=20)', 'debut': 'Oct. 2027', 'fin': 'Déc. 2027', 'duree': '3 mois', 'livrables': 'Corpus retranscrit'},
            {'etape': 'Analyse qualitative et croisée', 'debut': 'Janv. 2028', 'fin': 'Juin 2028', 'duree': '6 mois', 'livrables': 'Article soumis, analyse complète'},
            {'etape': 'Séminaire de restitution INSPÉ', 'debut': 'Sept. 2028', 'fin': 'Oct. 2028', 'duree': '1 mois', 'livrables': "Journée d'étude, captation vidéo"},
            {'etape': 'Guide formateurs + bilan', 'debut': 'Nov. 2028', 'fin': 'Déc. 2028', 'duree': '2 mois', 'livrables': 'Guide en ligne, bilan scientifique'},
        ],
        'financement_existant': {}, 'financement_cours': {}, 'financement_avenir': {},
        'etablissement': {
            'nom': 'Université de Lille',
            'agent_comptable': 'À compléter',
            'adresse': 'À compléter',
            'tel': 'À compléter', 'mel': 'À compléter',
            'resp_financier': 'À compléter',
        },
        'rib': {},
        'assujetti_tva': False,
        'demande_appui': "Mise à disposition d'une salle et d'un amphithéâtre à l'INSPÉ Lille HdF pour le séminaire de restitution (automne 2028). Captation vidéo du séminaire. Mise en relation avec la délégation académique au numérique et les référents éducation prioritaire.",
    }

    budget_lines = [
        {'nature': 'Vacations', 'detail': "Appui collecte 2027 : préparation questionnaire, relances, anonymisation, retranscriptions courtes", 'total': 1600, 'rh2027': 1600},
        {'nature': 'Prestations de recherche', 'detail': 'Traitement données qualitatives / transcription partielle externalisée', 'total': 1200, 'fonct2027': 1200},
        {'nature': 'Missions', 'detail': "Déplacements sites INSPÉ, réunions coordination et collecte complémentaire", 'total': 400, 'fonct2027': 400},
        {'nature': 'Matériel consommable / temporaire', 'detail': "Licence informatique temporaire et impressions dédiées à la phase d'enquête", 'total': 300, 'fonct2027': 300},
        {'nature': 'Vacations', 'detail': "Appui analyse finale 2028 : codage, mise en forme corpus, synthèses, aide rédaction guide", 'total': 1100, 'rh2028': 1100},
        {'nature': "Organisation séminaire / journée d'étude", 'detail': "Accueil café, repas intervenants, supports et logistique légère de restitution", 'total': 1200, 'fonct2028': 1200},
        {'nature': 'Missions', 'detail': "Déplacements et nuitée d'intervenants extérieurs pour le séminaire", 'total': 700, 'fonct2028': 700},
        {'nature': 'Prestation de valorisation', 'detail': "Mise en page graphique du guide final et du kit de ressources", 'total': 1200, 'fonct2028': 1200},
        {'nature': 'Matériel consommable / diffusion', 'detail': "Impressions finales limitées et supports de diffusion", 'total': 300, 'fonct2028': 300},
    ]

    pdf1 = build_annexe1(sample_project, logo)
    with open('/tmp/final_annexe1.pdf', 'wb') as f:
        f.write(pdf1)
    print(f'✓ Annexe 1  → {len(pdf1)//1024} KB, {__import__("pdfplumber").open("/tmp/final_annexe1.pdf").pages.__len__()} pages')

    pdf2 = build_annexe1bis(
        {'titre': sample_project['titre'], 'porteur': 'Matuszak, Céline', 'montant_inspe': 8000},
        budget_lines, logo
    )
    with open('/tmp/final_annexe1bis.pdf', 'wb') as f:
        f.write(pdf2)
    print(f'✓ Annexe 1bis → {len(pdf2)//1024} KB')
    print('DONE ✓')
