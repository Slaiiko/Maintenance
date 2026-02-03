import * as XLSX from 'xlsx';
import { RawRow, ProcessedSite, MaintenanceStatus } from './types';
import { differenceInYears, addYears, isBefore, isValid, parse } from 'date-fns';

// Helper to parse Excel dates which might be serial numbers or strings
const parseExcelDate = (val: string | number | undefined): Date | null => {
  if (!val) return null;
  
  // If number (Excel serial date)
  if (typeof val === 'number') {
    // Excel date to JS date conversion (approximate for 1900 system)
    return new Date(Math.round((val - 25569) * 86400 * 1000));
  }

  // If string, try standard formats
  if (typeof val === 'string') {
    const trimmed = val.trim();
    if (!trimmed) return null;
    
    // Try DD/MM/YYYY
    const dmy = trimmed.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})/);
    if (dmy) {
      return new Date(parseInt(dmy[3]), parseInt(dmy[2]) - 1, parseInt(dmy[1]));
    }
  }
  return null;
};

// Helper to determine if a maintenance cell means "Done", "Planned" or "NA"
const getMaintenanceStatus = (val: string | number | undefined): MaintenanceStatus => {
  if (!val) return MaintenanceStatus.UNKNOWN; // Usually implies not done or info missing
  
  const str = String(val).toLowerCase().trim();
  
  if (str.includes('a planifier') || str.includes('à planifier')) {
    return MaintenanceStatus.PLANNED;
  }
  
  if (str === 'na' || str === 'n/a' || str === 'so' || str === 's.o') {
    return MaintenanceStatus.NA;
  }
  
  // If it has a date or "RDM OK" or "RDM INEXISTANT", it's considered "Done/Handled" in terms of "Not waiting to be planned"
  if (str.includes('rdm') || parseExcelDate(val) !== null || str.length > 5) {
    return MaintenanceStatus.DONE;
  }
  
  return MaintenanceStatus.UNKNOWN;
};

export const processExcelData = (data: any[]): ProcessedSite[] => {
  const processed: ProcessedSite[] = [];
  const now = new Date();

  data.forEach((row: any, index) => {
    const r = row as RawRow;

    // Filter A: Only "Réalisé" or "En cours"
    // Normalize string: remove accents, lowercase, trim
    const rawA = String(r.A || '').trim();
    const cleanA = rawA.normalize("NFD").replace(/[\u0300-\u036f]/g, "").toLowerCase();

    // Check strict exclusion first (Annulé, NA, or Header row like "Statut")
    if (cleanA.includes('annule') || cleanA === 'na' || cleanA === 'n/a' || cleanA === 'so') return;

    // Robust check for "realise" (matches réalisé, Realise, Réaliser...) and "en cours"
    const isRealise = cleanA.includes('realis');
    const isEnCours = cleanA.includes('en cours');

    if (!isRealise && !isEnCours) return;

    // 1. Determine Contract Start Date
    // Use O. If O is "Pas de contrat" or empty, use N (only if R=Oui).
    const hasMaint = (r.R || '').toLowerCase().trim() === 'oui';
    let startDate = parseExcelDate(r.O);
    
    // Logic: If R=Oui and O=Pas de contrat (or empty/invalid), try N
    const oVal = String(r.O || '').toLowerCase();
    if (hasMaint && (!startDate || oVal.includes('pas de contrat'))) {
      const nDate = parseExcelDate(r.N);
      if (nDate) startDate = nDate;
    }

    // Parse Date Fin Chantier (K)
    const dateFinChantier = parseExcelDate(r.K);

    // 2. Statuses for Maintenance Columns L & M
    const m1Status = getMaintenanceStatus(r.L);
    const m2Status = getMaintenanceStatus(r.M);

    // 3. Contract Age & Status
    let contractStatus: ProcessedSite['contractStatus'] = 'Sans Objet';
    let age = 0;

    if (hasMaint && startDate) {
      age = differenceInYears(now, startDate);
      
      if (age >= 5) {
        contractStatus = 'Expiré';
      } else if (age >= 3) {
        contractStatus = 'Reconduction';
      } else {
        // Less than 3 years. Check if "Terminé" (Both L & M done/handled)
        // Rule: "si c'est inférieur a 3 ans et que la colonne M et L ont les valeurs inscrite... alors Terminé"
        const lDone = m1Status === MaintenanceStatus.DONE || m1Status === MaintenanceStatus.NA;
        const mDone = m2Status === MaintenanceStatus.DONE || m2Status === MaintenanceStatus.NA;
        
        if (lDone && mDone) {
          contractStatus = 'Terminé';
        } else {
          contractStatus = 'En cours';
        }
      }
    } else if (hasMaint && !startDate) {
        // R=Oui but no date found
        contractStatus = 'En cours'; // Default fallback or flag as error? Prompt implies "En cours" logic roughly
    }

    // 4. Calculate "Maintenance Needed" (Detail View Logic - Max 1 per line)
    // Rule: if L=Planned -> 1. Else if L=Done/NA and M=Planned -> 1. Else 0.
    let maintNeeded = 0;
    if (m1Status === MaintenanceStatus.PLANNED) {
      maintNeeded = 1;
    } else if (m2Status === MaintenanceStatus.PLANNED) {
      maintNeeded = 1;
    }

    // 5. Calculate "Late" (Retard)
    // Rule: L=Planned & StartDate > 1yr ago OR M=Planned & StartDate > 2yr ago
    let isLate = false;
    if (startDate) {
      const oneYearAgo = addYears(startDate, 1);
      const twoYearsAgo = addYears(startDate, 2);
      
      if (m1Status === MaintenanceStatus.PLANNED && isBefore(oneYearAgo, now)) isLate = true;
      if (m2Status === MaintenanceStatus.PLANNED && isBefore(twoYearsAgo, now)) isLate = true;
    }

    processed.push({
      id: `${index}-${r.C || 'no-id'}`,
      statusConstruction: rawA,
      hasMaintenanceContract: hasMaint,
      siteName: r.D || 'Site Inconnu',
      address: r.E || '',
      terminalsRaw: r.I || '',
      interlocutor: r.F || '',
      brand: r.H || '',
      dateContrat: startDate,
      dateFinChantier: dateFinChantier,
      dateMaintenance1Raw: String(r.L || ''),
      dateMaintenance2Raw: String(r.M || ''),
      contractAgeYears: age,
      contractStatus,
      maintenance1Status: m1Status,
      maintenance2Status: m2Status,
      maintenanceNeededCount: maintNeeded,
      isLate,
      raw: r
    });
  });

  return processed;
};

// Analyse intelligente du nombre de bornes
export const getTerminalCount = (raw: string): number => {
  const r = raw.toLowerCase().trim();
  if (!r) return 0;

  // 1. Chercher "2x", "2 x"
  const multMatch = r.match(/(\d+)\s*x/);
  if (multMatch) return parseInt(multMatch[1]);

  // 2. Chercher "double" ou "2 pdl"
  if (r.includes('double') || r.includes('2 pdl') || r.includes('2pdl')) return 2;

  // 3. Chercher un chiffre explicite au début (ex: "2 bornes")
  const startDigit = r.match(/^(\d+)\s/);
  if (startDigit) return parseInt(startDigit[1]);

  // Par défaut 1
  return 1;
};

// Analyse intelligente du type de bornes
export const parseTerminalType = (raw: string): string => {
  // Normalisation : minuscules, sans accents
  const r = raw.toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "");

  if (r.includes('60') || r.includes('dc')) {
    return 'Borne DC 60kW+';
  }
  
  if (r.includes('22')) {
    if (r.includes('2x') || r.includes('double')) return 'Borne AC 2x22kW';
    return 'Borne AC 22kW';
  }
  
  if (r.includes('7') || r.includes('3.7') || r.includes('7.4')) {
    if (r.includes('2x') || r.includes('double')) return 'Borne AC 2x7.4kW';
    return 'Borne AC 7.4kW';
  }

  if (r.includes('simple')) return 'Borne Simple (Inconnu)';
  if (r.includes('double')) return 'Borne Double (Inconnu)';

  return 'Autre / Non spécifié';
};

export const extractRegion = (address: string) => {
  const match = address.match(/\b\d{5}\b/);
  if (match) {
    const dept = match[0].substring(0, 2);
    return dept; // Just return department code
  }
  return 'Inconnu';
};

// Coordonnées approximatives (X%, Y%) sur une carte de France pour chaque département
export const DEPT_COORDS: Record<string, { x: number, y: number }> = {
  '01': { x: 73, y: 55 }, '02': { x: 62, y: 18 }, '03': { x: 58, y: 50 }, '04': { x: 85, y: 78 }, '05': { x: 86, y: 70 },
  '06': { x: 92, y: 80 }, '07': { x: 70, y: 68 }, '08': { x: 68, y: 15 }, '09': { x: 45, y: 90 }, '10': { x: 68, y: 30 },
  '11': { x: 55, y: 88 }, '12': { x: 55, y: 75 }, '13': { x: 78, y: 85 }, '14': { x: 35, y: 20 }, '15': { x: 55, y: 65 },
  '16': { x: 38, y: 58 }, '17': { x: 30, y: 55 }, '18': { x: 55, y: 45 }, '19': { x: 50, y: 60 }, '21': { x: 70, y: 42 },
  '22': { x: 18, y: 30 }, '23': { x: 50, y: 55 }, '24': { x: 40, y: 65 }, '25': { x: 85, y: 45 }, '26': { x: 72, y: 70 },
  '27': { x: 45, y: 22 }, '28': { x: 48, y: 30 }, '29': { x: 10, y: 32 }, '30': { x: 70, y: 80 }, '31': { x: 42, y: 88 },
  '32': { x: 40, y: 82 }, '33': { x: 30, y: 68 }, '34': { x: 60, y: 85 }, '35': { x: 28, y: 32 }, '36': { x: 48, y: 50 },
  '37': { x: 40, y: 42 }, '38': { x: 78, y: 62 }, '39': { x: 78, y: 48 }, '40': { x: 30, y: 80 }, '41': { x: 48, y: 40 },
  '42': { x: 68, y: 58 }, '43': { x: 65, y: 65 }, '44': { x: 25, y: 45 }, '45': { x: 52, y: 35 }, '46': { x: 50, y: 72 },
  '47': { x: 38, y: 75 }, '48': { x: 62, y: 75 }, '49': { x: 32, y: 42 }, '50': { x: 28, y: 22 }, '51': { x: 65, y: 25 },
  '52': { x: 75, y: 32 }, '53': { x: 32, y: 35 }, '54': { x: 82, y: 25 }, '55': { x: 75, y: 22 }, '56': { x: 18, y: 40 },
  '57': { x: 85, y: 20 }, '58': { x: 60, y: 45 }, '59': { x: 58, y: 8 }, '60': { x: 52, y: 18 }, '61': { x: 38, y: 28 },
  '62': { x: 52, y: 10 }, '63': { x: 60, y: 60 }, '64': { x: 28, y: 88 }, '65': { x: 38, y: 90 }, '66': { x: 58, y: 92 },
  '67': { x: 92, y: 22 }, '68': { x: 90, y: 35 }, '69': { x: 70, y: 58 }, '70': { x: 80, y: 40 }, '71': { x: 70, y: 50 },
  '72': { x: 40, y: 35 }, '73': { x: 85, y: 62 }, '74': { x: 85, y: 55 }, '75': { x: 54, y: 26 }, '76': { x: 45, y: 15 },
  '77': { x: 58, y: 28 }, '78': { x: 50, y: 26 }, '79': { x: 32, y: 52 }, '80': { x: 55, y: 15 }, '81': { x: 52, y: 82 },
  '82': { x: 45, y: 78 }, '83': { x: 82, y: 85 }, '84': { x: 75, y: 80 }, '85': { x: 25, y: 52 }, '86': { x: 40, y: 50 },
  '87': { x: 45, y: 58 }, '88': { x: 82, y: 32 }, '89': { x: 62, y: 35 }, '90': { x: 88, y: 42 }, '91': { x: 52, y: 28 },
  '92': { x: 53, y: 26 }, '93': { x: 55, y: 25 }, '94': { x: 55, y: 27 }, '95': { x: 52, y: 24 }
};