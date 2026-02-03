const fileInput = document.getElementById('file-input');
const summaryCards = document.getElementById('summary-cards');
const sitesTable = document.getElementById('sites-table');
const maintenanceFilter = document.getElementById('maintenance-filter');
const contractFilter = document.getElementById('contract-filter');
const todoFilter = document.getElementById('todo-filter');
const detailPanel = document.getElementById('site-detail-panel');
const detailSiteTitle = document.getElementById('detail-site-title');
const detailSiteAddress = document.getElementById('detail-site-address');
const detailMaintenanceBadge = document.getElementById('detail-maintenance-badge');
const detailInterlocuteur = document.getElementById('detail-interlocuteur');
const detailAddressLink = document.getElementById('detail-address-link');
const detailBornes = document.getElementById('detail-bornes');
const detailMarque = document.getElementById('detail-marque');
const detailContractStatus = document.getElementById('detail-contract-status');
const detailTodo = document.getElementById('detail-todo');
const detailDelay = document.getElementById('detail-delay');
const detailDevis = document.getElementById('detail-devis');
const detailAffaire = document.getElementById('detail-affaire');
const detailDateAffaire = document.getElementById('detail-date-affaire');
const detailDateFin = document.getElementById('detail-date-fin');
const detailDescriptif = document.getElementById('detail-descriptif');
const detailFacture = document.getElementById('detail-facture');
const detailAffaireRef = document.getElementById('detail-affaire-ref');

const charts = {
  region: null,
  type: null,
  year: null,
};

const tabs = document.querySelectorAll('.tab');
const panels = document.querySelectorAll('.tab-panel');
const detailTabs = document.querySelectorAll('.detail-tab');
const detailPanels = document.querySelectorAll('.detail-panel-content');

tabs.forEach((tab) => {
  tab.addEventListener('click', () => {
    tabs.forEach((item) => item.classList.remove('active'));
    panels.forEach((panel) => panel.classList.remove('active'));
    tab.classList.add('active');
    document.getElementById(tab.dataset.tab).classList.add('active');
  });
});

detailTabs.forEach((tab) => {
  tab.addEventListener('click', () => {
    detailTabs.forEach((item) => item.classList.remove('active'));
    detailPanels.forEach((panel) => panel.classList.remove('active'));
    tab.classList.add('active');
    document.getElementById(tab.dataset.detailTab).classList.add('active');
  });
});

const COLUMN_MAP = {
  statusChantier: 'A',
  devis: 'B',
  affaire: 'C',
  site: 'D',
  adresse: 'E',
  interlocuteur: 'F',
  descriptif: 'G',
  marque: 'H',
  bornes: 'I',
  dateAffaire: 'J',
  dateFin: 'K',
  maintenanceL: 'L',
  maintenanceM: 'M',
  dateFacture: 'N',
  dateContrat: 'O',
  facture: 'P',
  affaireRef: 'Q',
  siteMaintenance: 'R',
};

const REGION_COLUMN = 'T';

let allRows = [];

function normalizeString(value) {
  return String(value ?? '').trim();
}

function getColumnValue(row, letter) {
  return normalizeString(row[letter]);
}

function isIgnoredStatus(value) {
  const normalized = normalizeString(value).toLowerCase();
  return normalized === 'na' || normalized === 'annulé' || normalized === 'annule';
}

function isMaintenanceOui(value) {
  return normalizeString(value).toLowerCase() === 'oui';
}

function isMaintenanceNon(value) {
  return normalizeString(value).toLowerCase() === 'non';
}

function isAPlanifier(value) {
  return normalizeString(value).toLowerCase() === 'a planifier';
}

function parseDate(value) {
  if (!value) return null;
  if (value instanceof Date) return value;
  const parsed = new Date(value);
  if (Number.isNaN(parsed.getTime())) return null;
  return parsed;
}

function getStartDate(row) {
  const hasContract = normalizeString(getColumnValue(row, COLUMN_MAP.dateContrat)).toLowerCase() !== 'pas de contrat';
  const contractDate = parseDate(getColumnValue(row, COLUMN_MAP.dateContrat));
  const factureDate = parseDate(getColumnValue(row, COLUMN_MAP.dateFacture));
  if (hasContract && contractDate) return contractDate;
  return factureDate;
}

function getContractStatus(row) {
  if (!isMaintenanceOui(getColumnValue(row, COLUMN_MAP.siteMaintenance))) return 'Pas de maintenance';
  const startDate = getStartDate(row);
  if (!startDate) return 'À renseigner';
  const today = new Date();
  const diffYears = (today - startDate) / (1000 * 60 * 60 * 24 * 365.25);

  if (diffYears > 5) return 'Dépassé';
  if (diffYears >= 3 && diffYears <= 5) return 'Reconduction';

  const maintenanceL = getColumnValue(row, COLUMN_MAP.maintenanceL);
  const maintenanceM = getColumnValue(row, COLUMN_MAP.maintenanceM);
  const maintenanceDoneL = !isAPlanifier(maintenanceL) && maintenanceL !== '' && maintenanceL.toLowerCase() !== 'na';
  const maintenanceDoneM = !isAPlanifier(maintenanceM) && maintenanceM !== '' && maintenanceM.toLowerCase() !== 'na';

  if (maintenanceDoneL && maintenanceDoneM) return 'Terminé';
  return 'En cours';
}

function getContractKey(status) {
  const normalized = status.toLowerCase();
  if (normalized === 'en cours') return 'en_cours';
  if (normalized === 'reconduction') return 'reconduction';
  if (normalized === 'terminé') return 'termine';
  if (normalized === 'dépassé') return 'depasse';
  return 'autre';
}

function getMaintenanceTodoPerLine(row) {
  const maintenanceL = getColumnValue(row, COLUMN_MAP.maintenanceL);
  const maintenanceM = getColumnValue(row, COLUMN_MAP.maintenanceM);

  if (isAPlanifier(maintenanceL)) return 1;
  if (!isAPlanifier(maintenanceL) && maintenanceL !== '') {
    if (isAPlanifier(maintenanceM)) return 1;
    return 0;
  }
  return 0;
}

function getMaintenanceDoneGeneral(row) {
  const maintenanceL = getColumnValue(row, COLUMN_MAP.maintenanceL);
  const maintenanceM = getColumnValue(row, COLUMN_MAP.maintenanceM);
  const doneL = maintenanceL !== '' && !isAPlanifier(maintenanceL) && maintenanceL.toLowerCase() !== 'na';
  const doneM = maintenanceM !== '' && !isAPlanifier(maintenanceM) && maintenanceM.toLowerCase() !== 'na';
  return { doneL, doneM };
}

function isMaintenanceDelay(row) {
  const startDate = getStartDate(row);
  if (!startDate) return false;
  const diffYears = (new Date() - startDate) / (1000 * 60 * 60 * 24 * 365.25);
  const maintenanceL = getColumnValue(row, COLUMN_MAP.maintenanceL);
  const maintenanceM = getColumnValue(row, COLUMN_MAP.maintenanceM);

  if (isAPlanifier(maintenanceL) && diffYears > 1) return true;
  if (isAPlanifier(maintenanceM) && diffYears > 2) return true;
  return false;
}

function parseBornesCount(value) {
  const raw = normalizeString(value).toLowerCase();
  if (!raw) return 0;

  const compact = raw.replace(/\s+/g, '');
  const quantityMatches = [...compact.matchAll(/(\d+)\s*[x×]/g)];
  if (quantityMatches.length) {
    return quantityMatches.reduce((sum, match) => sum + Number(match[1]), 0);
  }

  if (compact.includes('double')) return 2;
  if (compact.includes('simple')) return 1;

  if (compact.includes('kw') || compact.match(/\d+(?:[.,]\d+)?/)) return 1;

  return 1;
}

function normalizeTypeLabel(value) {
  const raw = normalizeString(value).toLowerCase();
  if (!raw) return 'Non renseigné';

  const compact = raw.replace(/\s+/g, '');
  const xMatch = compact.match(/(\d+)\s*[x×]\s*(\d+(?:[.,]\d+)?)/);
  if (xMatch) return `${xMatch[1]}x${xMatch[2].replace(',', '.')}`;

  const kwMatch = compact.match(/(\d+(?:[.,]\d+)?)\s*kw/);
  if (kwMatch) return `${kwMatch[1].replace(',', '.')} kW`;

  if (compact.includes('double')) return 'Double';
  if (compact.includes('simple')) return 'Simple';

  const powerMatch = compact.match(/(\d+(?:[.,]\d+)?)/);
  if (powerMatch) return `${powerMatch[1].replace(',', '.')} kW`;

  return raw;
}

function rowToDisplay(row) {
  const addressValue = getColumnValue(row, COLUMN_MAP.adresse);
  const addressLink = addressValue
    ? `https://www.google.com/maps/search/?api=1&query=${encodeURIComponent(addressValue)}`
    : '';
  return {
    site: getColumnValue(row, COLUMN_MAP.site),
    addressValue,
    addressLink,
    maintenance: isMaintenanceOui(getColumnValue(row, COLUMN_MAP.siteMaintenance)) ? 'Oui' : 'Non',
    contractStatus: getContractStatus(row),
    todo: getMaintenanceTodoPerLine(row),
    delay: isMaintenanceDelay(row) ? 'Oui' : 'Non',
    bornes: getColumnValue(row, COLUMN_MAP.bornes),
    interlocuteur: getColumnValue(row, COLUMN_MAP.interlocuteur),
    marque: getColumnValue(row, COLUMN_MAP.marque),
    devis: getColumnValue(row, COLUMN_MAP.devis),
    affaire: getColumnValue(row, COLUMN_MAP.affaire),
    dateAffaire: getColumnValue(row, COLUMN_MAP.dateAffaire),
    dateFin: getColumnValue(row, COLUMN_MAP.dateFin),
    descriptif: getColumnValue(row, COLUMN_MAP.descriptif),
    facture: getColumnValue(row, COLUMN_MAP.facture),
    affaireRef: getColumnValue(row, COLUMN_MAP.affaireRef),
  };
}

function renderEmptyState() {
  summaryCards.innerHTML = '';
  const template = document.getElementById('empty-state');
  summaryCards.appendChild(template.content.cloneNode(true));
  sitesTable.innerHTML = '';

  Object.values(charts).forEach((chart) => {
    if (chart) chart.destroy();
  });
}

function renderSummary(data) {
  summaryCards.innerHTML = '';
  const cards = [
    { label: 'Sites avec maintenance (R = Oui)', value: data.maintenanceOui },
    { label: 'Sites sans maintenance (R = Non)', value: data.maintenanceNon },
    { label: 'Nombre de sites (doublons inclus)', value: data.totalSites },
    { label: 'Maintenances à faire (vue détail)', value: data.maintenanceTodoPerLine },
    { label: 'Maintenances à faire (total L + M)', value: data.maintenanceTodoGeneral },
    { label: 'Maintenances réalisées (L+M)', value: data.maintenanceDoneGeneral },
    { label: 'Reconductions à réaliser', value: data.reconduction },
    { label: 'Contrats expirés (> 5 ans)', value: data.expired },
    { label: 'Retards de maintenance', value: data.delays },
    { label: 'Bornes posées (approx.)', value: data.totalBornes },
  ];

  cards.forEach((card) => {
    const div = document.createElement('div');
    div.className = 'card';
    div.innerHTML = `<p>${card.label}</p><strong>${card.value}</strong>`;
    summaryCards.appendChild(div);
  });
}

function renderCharts(data) {
  const regionCtx = document.getElementById('region-chart');
  const typeCtx = document.getElementById('type-chart');
  const yearCtx = document.getElementById('year-chart');

  if (charts.region) charts.region.destroy();
  if (charts.type) charts.type.destroy();
  if (charts.year) charts.year.destroy();

  charts.region = new Chart(regionCtx, {
    type: 'pie',
    data: {
      labels: Object.keys(data.regionCounts),
      datasets: [
        {
          data: Object.values(data.regionCounts),
          backgroundColor: ['#4464ad', '#8b5cf6', '#10b981', '#f59e0b', '#ef4444', '#14b8a6'],
        },
      ],
    },
  });

  charts.type = new Chart(typeCtx, {
    type: 'pie',
    data: {
      labels: Object.keys(data.typeCounts),
      datasets: [
        {
          data: Object.values(data.typeCounts),
          backgroundColor: ['#1d4ed8', '#f97316', '#22c55e', '#eab308', '#ec4899', '#0ea5e9'],
        },
      ],
    },
  });

  charts.year = new Chart(yearCtx, {
    type: 'line',
    data: {
      labels: Object.keys(data.yearCounts),
      datasets: [
        {
          label: 'Bornes posées',
          data: Object.values(data.yearCounts),
          borderColor: '#1f2937',
          backgroundColor: '#93c5fd',
          tension: 0.2,
        },
      ],
    },
    options: {
      scales: {
        y: {
          beginAtZero: true,
        },
      },
    },
  });
}

function renderTables(rows) {
  const filteredRows = rows.filter((row) => {
    const maintenanceValue = isMaintenanceOui(getColumnValue(row, COLUMN_MAP.siteMaintenance));
    const contractStatus = getContractKey(getContractStatus(row));
    const todoValue = getMaintenanceTodoPerLine(row) > 0;

    if (maintenanceFilter.value === 'oui' && !maintenanceValue) return false;
    if (maintenanceFilter.value === 'non' && maintenanceValue) return false;

    if (contractFilter.value !== 'all' && contractStatus !== contractFilter.value) return false;

    if (todoFilter.value === 'todo' && !todoValue) return false;
    if (todoFilter.value === 'done' && todoValue) return false;

    return true;
  });

  sitesTable.innerHTML = '';
  filteredRows.forEach((row) => {
    const display = rowToDisplay(row);

    const siteRow = document.createElement('tr');
    siteRow.addEventListener('click', () => {
      setSiteDetail(display);
      siteRow.scrollIntoView({ behavior: 'smooth', block: 'center' });
    });
    siteRow.innerHTML = `
      <td>${display.site}</td>
      <td>${display.addressLink ? `<a href="${display.addressLink}" target="_blank">${display.addressValue}</a>` : ''}</td>
      <td>${display.maintenance}</td>
      <td>${display.contractStatus}</td>
      <td>${display.todo}</td>
      <td>${display.delay}</td>
      <td>${display.bornes}</td>
    `;
    sitesTable.appendChild(siteRow);
  });
}

function setSiteDetail(display) {
  detailSiteTitle.textContent = display.site || 'Site non renseigné';
  detailSiteAddress.textContent = display.addressValue || 'Adresse non renseignée';
  detailMaintenanceBadge.textContent = display.maintenance === 'Oui' ? 'Maintenance' : 'Sans maintenance';
  detailMaintenanceBadge.style.background = display.maintenance === 'Oui' ? '#e0f2fe' : '#fee2e2';
  detailMaintenanceBadge.style.color = display.maintenance === 'Oui' ? '#0369a1' : '#b91c1c';

  detailInterlocuteur.textContent = display.interlocuteur || '-';
  detailBornes.textContent = display.bornes || '-';
  detailMarque.textContent = display.marque || '-';
  detailContractStatus.textContent = display.contractStatus || '-';
  detailTodo.textContent = String(display.todo ?? '-');
  detailDelay.textContent = display.delay || '-';
  detailDevis.textContent = display.devis || '-';
  detailAffaire.textContent = display.affaire || '-';
  detailDateAffaire.textContent = display.dateAffaire || '-';
  detailDateFin.textContent = display.dateFin || '-';
  detailDescriptif.textContent = display.descriptif || '-';
  detailFacture.textContent = display.facture || '-';
  detailAffaireRef.textContent = display.affaireRef || '-';

  if (display.addressLink) {
    detailAddressLink.innerHTML = `<a href="${display.addressLink}" target="_blank">${display.addressValue}</a>`;
  } else {
    detailAddressLink.textContent = '-';
  }

  detailPanel.scrollIntoView({ behavior: 'smooth', block: 'start' });
}

function computeData(rows) {
  let maintenanceOui = 0;
  let maintenanceNon = 0;
  let maintenanceTodoPerLine = 0;
  let maintenanceTodoGeneral = 0;
  let maintenanceDoneGeneral = 0;
  let reconduction = 0;
  let expired = 0;
  let delays = 0;
  let totalBornes = 0;

  const regionCounts = {};
  const typeCounts = {};
  const yearCounts = {};

  rows.forEach((row) => {
    const maintenanceFlag = getColumnValue(row, COLUMN_MAP.siteMaintenance);
    if (isMaintenanceOui(maintenanceFlag)) maintenanceOui += 1;
    if (isMaintenanceNon(maintenanceFlag)) maintenanceNon += 1;

    maintenanceTodoPerLine += getMaintenanceTodoPerLine(row);
    if (isAPlanifier(getColumnValue(row, COLUMN_MAP.maintenanceL))) maintenanceTodoGeneral += 1;
    if (isAPlanifier(getColumnValue(row, COLUMN_MAP.maintenanceM))) maintenanceTodoGeneral += 1;
    const { doneL, doneM } = getMaintenanceDoneGeneral(row);
    maintenanceDoneGeneral += (doneL ? 1 : 0) + (doneM ? 1 : 0);

    const contractStatus = getContractStatus(row);
    if (contractStatus === 'Reconduction') reconduction += 1;
    if (contractStatus === 'Dépassé') expired += 1;

    if (isMaintenanceDelay(row)) delays += 1;

    const bornesValue = getColumnValue(row, COLUMN_MAP.bornes);
    totalBornes += parseBornesCount(bornesValue);

    const region = normalizeString(getColumnValue(row, REGION_COLUMN)) || 'Non renseigné';
    regionCounts[region] = (regionCounts[region] || 0) + parseBornesCount(bornesValue);

    const typeLabel = normalizeTypeLabel(getColumnValue(row, COLUMN_MAP.bornes));
    typeCounts[typeLabel] = (typeCounts[typeLabel] || 0) + 1;

    const yearDate = parseDate(getColumnValue(row, COLUMN_MAP.dateAffaire)) ||
      parseDate(getColumnValue(row, COLUMN_MAP.dateFin)) ||
      getStartDate(row);
    if (yearDate) {
      const year = yearDate.getFullYear();
      yearCounts[year] = (yearCounts[year] || 0) + parseBornesCount(bornesValue);
    }
  });

  const sortedYearCounts = Object.keys(yearCounts)
    .sort()
    .reduce((acc, key) => {
      acc[key] = yearCounts[key];
      return acc;
    }, {});

  return {
    maintenanceOui,
    maintenanceNon,
    totalSites: rows.length,
    maintenanceTodoPerLine,
    maintenanceTodoGeneral,
    maintenanceDoneGeneral,
    reconduction,
    expired,
    delays,
    totalBornes: Math.round(totalBornes),
    regionCounts,
    typeCounts,
    yearCounts: sortedYearCounts,
  };
}

function updateView(rows) {
  if (!rows.length) {
    renderEmptyState();
    return;
  }

  const data = computeData(rows);
  renderSummary(data);
  renderCharts(data);
  renderTables(rows);
  if (rows.length) {
    setSiteDetail(rowToDisplay(rows[0]));
  }
}

function parseSheet(sheet) {
  const json = XLSX.utils.sheet_to_json(sheet, { header: 'A', defval: '' });
  return json;
}

function handleFiles(files) {
  const filePromises = [...files].map((file) => file.arrayBuffer().then((buffer) => {
    const workbook = XLSX.read(buffer, { type: 'array' });
    const rows = [];
    workbook.SheetNames.forEach((name) => {
      rows.push(...parseSheet(workbook.Sheets[name]));
    });
    return rows;
  }));

  Promise.all(filePromises)
    .then((results) => {
      allRows = results.flat().filter((row) => !isIgnoredStatus(getColumnValue(row, COLUMN_MAP.statusChantier)));
      updateView(allRows);
    })
    .catch((error) => {
      console.error('Erreur de lecture des fichiers', error);
    });
}

fileInput.addEventListener('change', (event) => {
  if (!event.target.files.length) return;
  handleFiles(event.target.files);
});

[maintenanceFilter, contractFilter, todoFilter].forEach((input) => {
  input.addEventListener('change', () => updateView(allRows));
});

renderEmptyState();
