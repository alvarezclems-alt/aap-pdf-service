/**
 * pdfClient.ts
 * Client TypeScript pour appeler le service de génération PDF (api.py)
 * depuis l'app Lovable / React.
 *
 * Usage :
 *   import { downloadAnnexe1, downloadAnnexe1bis, downloadAll } from './pdfClient'
 *   await downloadAll(project, budgetLines)
 */

const PDF_API_URL = import.meta.env.VITE_PDF_API_URL ?? 'http://localhost:5050'

// ─── Types ────────────────────────────────────────────────────────────────────

export interface Coordinateur {
  nom: string
  grade: string
  email: string
  telephone: string
  institution: string
  ur_id: string
  ur_nom: string
  ur_directeur: string
  ur_gestionnaire: string
  ur_tutelle: string
  membres: string
}

export interface Partenaire {
  nom: string
  coordonnees: string
  expertise: string
  contact: string
  titre_contact: string
  email_contact: string
  membres_partenaire: string
  etat_partenariat: string
}

export interface CalendrierRow {
  etape: string
  debut: string
  fin: string
  duree: string
  livrables: string
}

export interface Financement {
  type: string
  financeur: string
  dispositif: string
  montant: string
  eligibles: string
}

export interface Etablissement {
  nom: string
  agent_comptable: string
  adresse: string
  tel: string
  mel: string
  resp_financier: string
  adresse_resp?: string
  tel_resp?: string
  mel_resp?: string
}

export interface RIB {
  banque: string
  titulaire: string
  domiciliation: string
  compte: string
  code_banque: string
  cle_rib: string
  code_guichet: string
}

export interface ProjectForPDF {
  titre: string
  acronyme: string
  mots_cles: string[]
  resume: string
  coordinateur: Coordinateur
  partenaires: Partenaire[]
  publications: string
  description: string
  calendrier: CalendrierRow[]
  financement_existant?: Financement
  financement_cours?: Financement
  financement_avenir?: Financement
  etablissement?: Etablissement
  rib?: RIB
  assujetti_tva?: boolean
  demande_appui: string
  montant_inspe?: number
}

export interface BudgetLine {
  nature: string
  detail: string
  total: number
  fonct2027: number
  rh2027: number
  fonct2028: number
  rh2028: number
  cofinancement?: number
}

// ─── Helpers ──────────────────────────────────────────────────────────────────

async function postAndDownload(
  endpoint: string,
  body: object,
  filename: string
): Promise<void> {
  const res = await fetch(`${PDF_API_URL}${endpoint}`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify(body),
  })

  if (!res.ok) {
    const err = await res.json().catch(() => ({ error: 'Unknown error' }))
    throw new Error(err.error ?? `HTTP ${res.status}`)
  }

  const blob = await res.blob()
  const url  = URL.createObjectURL(blob)
  const a    = document.createElement('a')
  a.href     = url
  a.download = filename
  a.click()
  URL.revokeObjectURL(url)
}

// ─── Public API ───────────────────────────────────────────────────────────────

/** Télécharge uniquement l'Annexe 1 (dossier de candidature). */
export async function downloadAnnexe1(project: ProjectForPDF): Promise<void> {
  const safe = (project.acronyme ?? 'projet').replace(/[^a-zA-Z0-9]/g, '_')
  await postAndDownload('/generate-annexe1', { project }, `Annexe1_${safe}.pdf`)
}

/** Télécharge uniquement l'Annexe 1bis (tableau des dépenses). */
export async function downloadAnnexe1bis(
  project: ProjectForPDF,
  budgetLines: BudgetLine[]
): Promise<void> {
  const safe = (project.titre ?? 'budget').slice(0, 30).replace(/[^a-zA-Z0-9]/g, '_')
  await postAndDownload(
    '/generate-annexe1bis',
    { project: { titre: project.titre, porteur: project.coordinateur?.nom, montant_inspe: project.montant_inspe ?? 8000 }, budget_lines: budgetLines },
    `Annexe1bis_${safe}.pdf`
  )
}

/** Télécharge un ZIP contenant les deux annexes. */
export async function downloadAll(
  project: ProjectForPDF,
  budgetLines: BudgetLine[]
): Promise<void> {
  const safe = (project.acronyme ?? 'projet').replace(/[^a-zA-Z0-9]/g, '_')
  await postAndDownload(
    '/generate-all',
    { project, budget_lines: budgetLines },
    `Dossier_AAP_${safe}.zip`
  )
}

/** Vérifie que le service PDF est accessible. */
export async function checkPdfService(): Promise<boolean> {
  try {
    const res = await fetch(`${PDF_API_URL}/health`, { signal: AbortSignal.timeout(3000) })
    return res.ok
  } catch {
    return false
  }
}
