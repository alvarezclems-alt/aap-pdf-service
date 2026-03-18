# AAP PDF Service — Guide complet

## Ce que fait ce service

Il génère les **Annexe 1** (document de candidature) et **Annexe 1bis** (tableau des dépenses) au format PDF officiel INSPÉ Lille HdF, pixel-perfect :
- Logo INSPÉ + Université de Lille centré en haut
- Sections en bandes grises `RGB(241,241,241)`
- Tableaux formulaires à 2 colonnes avec bordures noires
- Tableau budget avec colonnes colorées exactes (jaune/vert/bleu) + ligne total jaune
- Numérotation en bas à droite
- Format A4, marges officielles

---

## Architecture

```
aap-pdf-service/
  generate_annexes.py   ← moteur Python (ReportLab) — peut s'utiliser seul
  api.py                ← serveur Flask exposant 3 endpoints REST
  logo_inspe.png        ← logo extrait des PDFs officiels
  requirements.txt      ← dépendances Python
  pdfClient.ts          ← client TypeScript pour l'app Lovable
```

---

## Option A — Utilisation standalone (script Python)

```python
from generate_annexes import build_annexe1, build_annexe1bis, load_logo

logo = load_logo('logo_inspe.png')

project = {
    'titre': 'Mon projet de recherche',
    'acronyme': 'MONPROJ',
    'mots_cles': ['éducation', 'numérique', 'formation'],
    'resume': 'Résumé du projet en 10 lignes...',
    'coordinateur': {
        'nom': 'Dupont, Marie',
        'grade': 'MCF',
        'email': 'marie.dupont@univ-lille.fr',
        'telephone': '03 20 00 00 00',
        'institution': 'Université de Lille – INSPÉ Lille HdF',
        'ur_id': 'EA 0000',
        'ur_nom': 'Mon Laboratoire',
        'ur_directeur': 'Martin, Jean – j.martin@univ-lille.fr',
        'ur_gestionnaire': 'À compléter',
        'ur_tutelle': 'Université de Lille',
        'membres': 'Nom1, Prenom1 (MCF) – email1@univ-lille.fr',
    },
    'partenaires': [],
    'publications': '1. ...\n2. ...',
    'description': '## 1. Contexte\n...\n## 2. Méthodes\n...',
    'calendrier': [
        {'etape': 'Phase 1', 'debut': 'Janv. 2027', 'fin': 'Juin 2027',
         'duree': '6 mois', 'livrables': 'Rapport intermédiaire'},
    ],
    'demande_appui': 'Mise à disposition d\'une salle...',
}

budget_lines = [
    {'nature': 'Vacations', 'detail': 'Retranscriptions', 'total': 1600,
     'rh2027': 1600, 'fonct2027': 0, 'fonct2028': 0, 'rh2028': 0},
    # ... autres lignes
]

# Générer les PDFs
pdf1 = build_annexe1(project, logo)
pdf2 = build_annexe1bis({'titre': project['titre'], 'porteur': 'Dupont, Marie'}, budget_lines, logo)

with open('Annexe1.pdf', 'wb') as f: f.write(pdf1)
with open('Annexe1bis.pdf', 'wb') as f: f.write(pdf2)
```

---

## Option B — Service Flask (intégration Lovable)

### 1. Installation

```bash
cd aap-pdf-service
pip install -r requirements.txt
```

### 2. Démarrage

```bash
python api.py
# ou avec un port personnalisé :
PORT=8080 python api.py
```

### 3. Ajouter dans Lovable

Dans **Lovable Settings → Secrets** :
```
VITE_PDF_API_URL = http://localhost:5050
```

### 4. Utiliser le client TypeScript

Copier `pdfClient.ts` dans `src/lib/` de votre projet Lovable, puis :

```typescript
import { downloadAll, downloadAnnexe1, downloadAnnexe1bis } from './lib/pdfClient'

// Télécharger les 2 PDFs dans un ZIP :
await downloadAll(projectForPDF, budgetLines)

// Ou séparément :
await downloadAnnexe1(projectForPDF)
await downloadAnnexe1bis(projectForPDF, budgetLines)
```

### 5. Mapper les données du store vers ProjectForPDF

```typescript
import type { Project, AapProfile } from '../types'
import type { ProjectForPDF, BudgetLine } from './pdfClient'

export function toProjectForPDF(project: Project, aap: AapProfile): ProjectForPDF {
  return {
    titre:      project.titre,
    acronyme:   project.acronyme,
    mots_cles:  project.motsCles,
    resume:     project.sections?.resume ?? '',
    coordinateur: {
      nom:           project.coordinateur.nom,
      grade:         project.coordinateur.grade,
      email:         project.coordinateur.email,
      telephone:     project.coordinateur.telephone,
      institution:   project.coordinateur.institution,
      ur_id:         project.unite.identifiant,
      ur_nom:        project.unite.nom,
      ur_directeur:  project.unite.directeur,
      ur_gestionnaire: project.unite.gestionnaire,
      ur_tutelle:    project.unite.tutelle,
      membres:       project.membres.map(m => `${m.nom} (${m.grade}) – ${m.email}`).join('\n'),
    },
    partenaires:  [],
    publications: project.sections?.publications ?? '',
    description:  project.sections?.description ?? '',
    calendrier:   parseCalendrier(project.sections?.calendrier ?? ''),
    demande_appui: project.sections?.appui ?? '',
    montant_inspe: project.montantDemande,
  }
}

export function toBudgetLines(budgetText: string): BudgetLine[] {
  // Parse le texte tableau généré par l'IA
  return budgetText
    .split('\n')
    .filter(l => l.includes('|') && !l.match(/^[-|]+$/))
    .slice(1)
    .map(l => {
      const cols = l.split('|').map(c => c.trim()).filter(Boolean)
      return {
        nature:      cols[0] ?? '',
        detail:      cols[1] ?? '',
        total:       parseFloat((cols[2] ?? '0').replace(/[^\d.,]/g, '').replace(',', '.')) || 0,
        fonct2027:   parseFloat((cols[3] ?? '0').replace(/[^\d.,]/g, '').replace(',', '.')) || 0,
        rh2027:      parseFloat((cols[4] ?? '0').replace(/[^\d.,]/g, '').replace(',', '.')) || 0,
        fonct2028:   parseFloat((cols[5] ?? '0').replace(/[^\d.,]/g, '').replace(',', '.')) || 0,
        rh2028:      parseFloat((cols[6] ?? '0').replace(/[^\d.,]/g, '').replace(',', '.')) || 0,
        cofinancement: 0,
      }
    })
    .filter(l => l.nature && l.total > 0)
}

function parseCalendrier(text: string) {
  return text
    .split('\n')
    .filter(l => l.includes('|') && !l.match(/^[-|]+$/))
    .slice(1)
    .map(l => {
      const cols = l.split('|').map(c => c.trim()).filter(Boolean)
      return { etape: cols[0] ?? '', debut: cols[1] ?? '', fin: cols[2] ?? '',
               duree: cols[3] ?? '', livrables: cols[4] ?? '' }
    })
    .filter(r => r.etape)
}
```

---

## Endpoints REST

| Méthode | URL | Body | Réponse |
|---------|-----|------|---------|
| GET | `/health` | — | `{"status": "ok"}` |
| POST | `/generate-annexe1` | `{project: {...}}` | PDF binaire |
| POST | `/generate-annexe1bis` | `{project: {...}, budget_lines: [...]}` | PDF binaire |
| POST | `/generate-all` | `{project: {...}, budget_lines: [...]}` | ZIP (2 PDFs) |

---

## Déploiement production

Pour déployer sur un serveur (Render, Railway, Fly.io) :

```bash
# Dockerfile minimal
FROM python:3.12-slim
WORKDIR /app
COPY requirements.txt .
RUN pip install -r requirements.txt gunicorn
COPY . .
CMD ["gunicorn", "--bind", "0.0.0.0:5050", "api:app"]
```

Variables d'environnement à définir dans Lovable :
```
VITE_PDF_API_URL = https://votre-service.railway.app
```
