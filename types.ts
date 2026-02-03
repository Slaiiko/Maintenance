
export enum MaintenanceStatus {
  DONE = 'DONE',
  PLANNED = 'PLANNED', // "A planifier"
  NA = 'NA',
  UNKNOWN = 'UNKNOWN'
}

export interface RawRow {
  A: string; // Realiser / En cours / NA / Annulé
  B: string; // Devis
  C: string; // Affaire
  D: string; // Site
  E: string; // Adresse
  F: string; // Interlocuteur
  G: string; // Descriptif
  H: string; // Marque/Série
  I: string; // Bornes
  J: string | number; // Date Affaire
  K: string | number; // Date Fin Chantier
  L: string | number; // Maintenance 1 (Year 2)
  M: string | number; // Maintenance 2 (Year 3)
  N: string | number; // Date Facture
  O: string | number; // Date Contrat Maintenance
  P: string; // Facture
  Q: string; // Affaire
  R: string; // Site avec Maintenance (Oui/Non)
}

export interface ProcessedSite {
  id: string;
  statusConstruction: string; // Col A
  hasMaintenanceContract: boolean; // Col R
  siteName: string; // Col D
  address: string; // Col E
  terminalsRaw: string; // Col I
  interlocutor: string; // Col F
  brand: string; // Col H
  
  // Dates
  dateContrat: Date | null; // Col O (or N fallback)
  dateFinChantier: Date | null; // Col K parsed
  dateMaintenance1Raw: string; // Col L
  dateMaintenance2Raw: string; // Col M
  
  // Computed Statuses
  contractAgeYears: number;
  contractStatus: 'En cours' | 'Reconduction' | 'Expiré' | 'Terminé' | 'Sans Objet';
  
  // Maintenance Logic
  maintenance1Status: MaintenanceStatus;
  maintenance2Status: MaintenanceStatus;
  
  // Specific Logics
  maintenanceNeededCount: number; // 0 or 1 (Detail View Logic)
  isLate: boolean;
  
  // Raw Data for details
  raw: RawRow;
}

export interface SiteListFilters {
  status?: string;
  maintenance?: string; // 'needed', 'late'
  search?: string;
  region?: string;
}

export type ViewMode = 'dashboard' | 'list' | 'planning';
