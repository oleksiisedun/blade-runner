const tableRows = 93;
const tableColumns = 13;
const separator = ', ';
const separatorRegex = /,\s?/;
const ranks = [
  [/\(О\)$/, 'оф.'],
  [/\(ОМ\)$/, 'оф.м.'],
  [/\(С\)$/, 'с-нт.'],
  [/\(СМ\)$/, 'с-нт.м.'],
  [/[^\)]$/, 'солд.'],
  [/\(М\)$/, 'солд.м.']
];
const nonUnitSheets = ['Відомість', 'Додаток', 'Handbook'];

const getActiveSheet = () => SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

/** @param {Array} arr */
const flat = arr => arr.reduce((acc, r) => [...acc, ...Array.isArray(r) ? flat(r) : [r]], []);

/** @param {Array} personnel */
const getPersonnel = (...personnel) => flat(personnel)
  .filter(p => p && typeof p === 'string' && !/«.+»/.test(p))
  .reduce((acc, r) => [...acc, ...r.split(separatorRegex)], []);

/** @param {Array} personnel */
const getForeignPersonnel = personnel => personnel.filter(p => /[\u0400-\u04FF\s\d]+ ([\u0400-\u04FF]{2,}|\d+)$/.test(p));

/** @param {string} rank */
const isOfficer = rank => /(лейтенант|капітан|майор|полковник)/i.test(rank);

/** @param {string} rank */
const isSergeant = rank => /сержант/i.test(rank);

/** @param {string} call */
const isMobilized = call => /^мобіл/i.test(call);

/** @param {Array} personnel */
const countOfficers = (...personnel) => countByRank(ranks[0][0], personnel);

/** @param {Array} personnel */
const countMobilizedOfficers = (...personnel) => countByRank(ranks[1][0], personnel);

/** @param {Array} personnel */
const countSergeants = (...personnel) => countByRank(ranks[2][0], personnel);

/** @param {Array} personnel */
const countMobilizedSergeants = (...personnel) => countByRank(ranks[3][0], personnel);

/** @param {Array} personnel */
const countSoldiers = (...personnel) => countByRank(ranks[4][0], personnel);

/** @param {Array} personnel */
const countMobilizedSoldiers = (...personnel) => countByRank(ranks[5][0], personnel);

/** @param {Array} personnel */
const getDuplicatedPersonnel = personnel => personnel.filter((p, i, arr) => arr.lastIndexOf(p) !== i);

const getAllTableValues = (sheet = getActiveSheet()) => sheet.getRange(1, 1, tableRows, tableColumns).getValues();

function getAllAvailableUnitPersonnel(sheet = getActiveSheet()) {
  const id = `allAvailableUnitPersonnel${sheet.getName()}`;
  const cache = CacheService.getScriptCache();
  const cached = cache.get(id);
  
  if (cached != null) return JSON.parse(cached);
  
  const allAvailableUnitPersonnel = getPersonnel(sheet.getRange('S2:T').getValues());
  cache.put(id, JSON.stringify(allAvailableUnitPersonnel), 180);

  return allAvailableUnitPersonnel;
}

/** 
 * @param {RegExp} rank
 * @param {Array} personnel 
 */
function countByRank(rank, personnel) {
  const allPersonnel = getPersonnel(personnel);
  const foreignPersonnel = getForeignPersonnel(allPersonnel);

  return allPersonnel.filter(p => !foreignPersonnel.includes(p)).filter(p => rank.test(p)).length;
}

/** @param {string} fullName */
function getFullNameParts(fullName) {
  const parts = fullName.trim().split(/\s+/);
  let surname, name, middleName;

  if (parts.length > 3) {
    if (parts.at(-1) === 'Огли') parts.pop();
    
    middleName = parts.pop();
    name = parts.pop();
    surname = parts.join(' ');
  } else {
    [surname, name = '', middleName = ''] = parts;
  }

  return { surname, name, middleName };
}

/** @param {Array} foreignPersonnel */
const getForeignPersonnelByUnit = foreignPersonnel => foreignPersonnel.map(([unit, fullName]) => {
  if (!unit) return null;

  const { surname, name, middleName } = getFullNameParts(fullName);
  const commonSurname = foreignPersonnel.filter(([, fullName]) => fullName.split(' ')[0] === surname).length > 1;

  return `${surname}${commonSurname ? ` ${name[0]} ${middleName[0]}` : ''} ${unit}`;
});

/** @param {Array} personnel */
const getPersonnelByRank = personnel => personnel.filter(([rank]) => rank).map(([rank, fullName, call]) => {
  if (!rank) return null;

  const { surname, name } = getFullNameParts(fullName);
  const commonSurname = personnel.filter(([, fullName]) => fullName.split(' ')[0] === surname).length > 1;
  const modifiers = (isOfficer(rank) ? 'О' : isSergeant(rank) ? 'С' : '') + (isMobilized(call) ? 'М' : '');
  
  return surname + (commonSurname ? ` ${name[0]}` : '') + (modifiers ? ` (${modifiers})` : '');
});

/** @param {String|Array} personnel */
function countPersonnelDetailed(...personnel) {
  const allPersonnel = getPersonnel(personnel);
  const personnelDetailed = ranks.reduce((acc, r) => {
    const personnelWithRank = allPersonnel.filter(p => r[0].test(p));
    if (personnelWithRank.length) acc.push(`${personnelWithRank.length} ${r[1]}`);
    return acc;
  }, []);

  return allPersonnel.length ? `${allPersonnel.length} (${personnelDetailed.join(separator)})` : '';
}

/** @param {Array} personnel */
function countPersonnelOnPositionsDetailed(...personnel) {
  const allPersonnel = getPersonnel(personnel);
  const foreignPersonnel = getForeignPersonnel(allPersonnel);
  const nativePersonnel = allPersonnel.filter(p => !foreignPersonnel.includes(p));

  return `${countPersonnelDetailed(nativePersonnel)}${foreignPersonnel.length ? ` + ${foreignPersonnel.length}` : ''}`;
}

/** @param {Array} personnel */
function countPlacedPersonnelDetailed(...personnel) {
  const allPersonnel = getPersonnel(personnel);
  const foreignPersonnel = getForeignPersonnel(allPersonnel);

  return countPersonnelDetailed(allPersonnel.filter(p => !foreignPersonnel.includes(p)));
}

/** 
 * @param {Array} allAvailablePersonnel
 * @param {Array} personnel 
 */
function countPersonnelSummary(allAvailablePersonnel, ...personnel) {
  const [управління, вільназміна,, позиції] = personnel;

  return [
    [
      countOfficers(allAvailablePersonnel),
      countMobilizedOfficers(allAvailablePersonnel),
      countSergeants(allAvailablePersonnel),
      countMobilizedSergeants(allAvailablePersonnel),
      countSoldiers(allAvailablePersonnel),
      countMobilizedSoldiers(allAvailablePersonnel)
    ],
    [
      countOfficers(personnel),
      countMobilizedOfficers(personnel),
      countSergeants(personnel),
      countMobilizedSergeants(personnel),
      countSoldiers(personnel),
      countMobilizedSoldiers(personnel)
    ],
    [countPersonnelDetailed(управління, вільназміна)],
    [countPlacedPersonnelDetailed(personnel)],
    [countPersonnelOnPositionsDetailed(позиції)]
  ]
}

/** @param {Array} data */
const sortByDates = data => data.toSorted((a, b) => {
  const [, dateA] = a;
  const [, dateB] = b;
  const [dayA, monthA] = dateA.split('.');
  const [dayB, monthB] = dateB.split('.');

  return monthA - monthB || dayA - dayB;
});

function showMissedPersonnel() {
  const sheet = getActiveSheet();

  if (sheet.getName() === 'Відомість') return;
  
  const allAvailablePersonnel = getAllAvailableUnitPersonnel(sheet);
  const allTableValues = getAllTableValues(sheet);
  const allAddedPersonnel = getPersonnel(allTableValues).filter(v => allAvailablePersonnel.includes(v));
  const missedPersonnel = allAvailablePersonnel.filter(p => !allAddedPersonnel.includes(p));

  SpreadsheetApp.getUi().alert(missedPersonnel.length ? missedPersonnel : '💪');
}

function showDuplicatedPersonnel() {
  const sheet = getActiveSheet();

  if (sheet.getName() === 'Відомість') return;

  const allAvailablePersonnel = getAllAvailableUnitPersonnel(sheet);
  const allTableValues = getAllTableValues(sheet);
  const allAddedPersonnel = getPersonnel(allTableValues).filter(v => allAvailablePersonnel.includes(v));
  const duplicatedPersonnel = getDuplicatedPersonnel(allAddedPersonnel)

  SpreadsheetApp.getUi().alert(duplicatedPersonnel.length ? duplicatedPersonnel : '💪');
}

function showRedundantPersonnel() {
  const sheet = getActiveSheet();
  const cellA1Notation = SpreadsheetApp.getUi().prompt('Введіть адресу клітинки, наприклад C7').getResponseText();

  if (sheet.getName() === 'Відомість' || !cellA1Notation) return;

  const allAvailablePersonnel = getAllAvailableUnitPersonnel(sheet);
  const cellPersonnel = sheet.getRange(cellA1Notation).getValue().split(separatorRegex);
  const redundantPersonnel = cellPersonnel.filter(p => !allAvailablePersonnel.includes(p));

  SpreadsheetApp.getUi().alert(redundantPersonnel.length ? redundantPersonnel : '💪');
}

/** 
 * @param {Array} units 
 * @param {Array} names
*/
function getPersonnelPresenceStatus(units, names) {
  const sheets = {};
  const lineNumbers = {};
  const missedPersonnel = {};
  const missedPersonnelRaw = {};

  return units.map(([u], i) => {
    if (!u) return [null];

    const unit = u.replace(/\s+\(.+\)/, '');
    const name = names[i][0];

    if (!sheets[unit]) sheets[unit] = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(unit);
    if (!lineNumbers[unit]) lineNumbers[unit] = sheets[unit].getRange('A:A').getValues().findIndex(([v]) => `${v}`.includes('Відсутні')) + 2;
    if (!missedPersonnelRaw[unit]) missedPersonnelRaw[unit] = sheets[unit].getRange(`A${lineNumbers[unit]}:L${tableRows}`).getValues();
    if (!missedPersonnel[unit]) missedPersonnel[unit] = flat(missedPersonnelRaw[unit]).filter(p => p);
    if (missedPersonnel[unit].includes(name)) {
      let index;

      for (const row of missedPersonnelRaw[unit]) {
        index = row.indexOf(name);
        if (index != -1) break;
      }

      return ['🔴', missedPersonnelRaw[unit][0][index].replace(/\s+\(.+\)/, '')];
    }

    return ['🟢', null];
  });
}

/**
 * @param {SpreadsheetApp.Sheet} sheet
 * @param {Array} ranges
 */
function refreshRanges(sheet, ...ranges) {
  const storage = {};

  ranges.forEach(range => {
    if (range.includes(':')) { 
      storage[range] = { formulas: sheet.getRange(range).getFormulas(), values: sheet.getRange(range).getValues() };
    } else {
      storage[range] = { formula: sheet.getRange(range).getFormula() };
    }
  });

  ranges.forEach(range => sheet.getRange(range).clearContent());
  SpreadsheetApp.flush();

  ranges.forEach(range => {
    if (range.includes(':')) {
      const { values, formulas } = storage[range];

      for (let i = 0; i < values.length; i += 1) {
        for (let j = 0; j < values[i].length; j += 1) {
          if (formulas[i][j] === '') continue;
          values[i][j] = formulas[i][j];
        }
      }

      sheet.getRange(range).setValues(values);
    } else {
      const { formula } = storage[range];

      sheet.getRange(range).setFormula(formula);
    }
  });
  SpreadsheetApp.flush();
}

function refreshUnitSheet(sheet = getActiveSheet()) {
  if (nonUnitSheets.includes(sheet.getName())) return;
  
  const lock = LockService.getDocumentLock();
  
  lock.waitLock(10 * 1000);
  refreshRanges(sheet, 'U2', 'X3');
  lock.releaseLock();
}

function refreshHandbookSheet() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Handbook');
  const lock = LockService.getDocumentLock();
  
  lock.waitLock(10 * 1000);
  refreshRanges(sheet, 'A1', 'H1');
  lock.releaseLock();
}

function refreshAllUnitSheets() {
  SpreadsheetApp.getActiveSpreadsheet().getSheets().forEach(sheet => {
    if (['Відомість', 'Додаток', 'Handbook'].includes(sheet.getName())) return;
    refreshUnitSheet(sheet);
  });
}

function exportPDF() {
  const id = SpreadsheetApp.getActiveSpreadsheet().getId();
  const exportFolderId = '1wMEfjgV0vd6CV_dwBXp3GArJRfqBmryC';
  const options = 'format=pdf&portrait=false&size=a4&scale=4&left_margin=0.5&right_margin=0.5&top_margin=0.5&bottom_margin=0.5';
  const url = `https://docs.google.com/spreadsheets/d/${id}/export?${options}`;
  const date = new Date(Date.now()).toLocaleString('uk-UA', { day: 'numeric', month: 'numeric', year: 'numeric' });
  const time = new Date(Date.now()).toLocaleTimeString('uk-UA', { hour: '2-digit', minute: '2-digit' });

  const blob = getFileAsBlob(url);
  blob.setName(`СТРОЙОВКА ${date} ${time}.pdf`);

  const exportFolder = DriveApp.getFolderById(exportFolderId);
  exportFolder.createFile(blob);
}

function getFileAsBlob(url) {
  const response = UrlFetchApp.fetch(url, {
    muteHttpExceptions: true,
    headers: { Authorization: `Bearer ${ScriptApp.getOAuthToken()}` },
  });

  return response.getBlob();
}

function refreshAndExportPDF() {
  refreshAllUnitSheets();
  refreshHandbookSheet();
  Utilities.sleep(30 * 1000);
  exportPDF();
}

function getPositionsMissedRanges() {
  const allTableValues = getAllTableValues();
  const rows = [
    allTableValues.findIndex(([v]) => `${v}`.startsWith('Позиції')) + 3,
    allTableValues.findIndex(([v]) => `${v}`.startsWith('Відсутні'))
  ];
  const columns = Array.from({ length: tableColumns - 1 }, (_, i) => i).filter(n => !(n % 2)).map(n => String.fromCharCode(n + 'A'.charCodeAt()));
  
  return columns.map(column => [`${column}${rows[0]}:${column}${rows[1]}`, `${column}${rows[1] + 3}:${column}${tableRows}`]);
}

function getAddedPersonnelUnits(personnel, order) {
  const names = getPersonnel(personnel);
  const unitRegExp = /(\d\s)?([\u0400-\u04FF]{2,}|\d+)$/;
  const namesUnits = names.map(p => (p.match(unitRegExp) ?? [])[0]);
  const allUnits = [...new Set(namesUnits)];
  const units = order.map(([p]) => (p.match(unitRegExp) ?? [])[0]);
  const emptyUnits = allUnits.filter(u => !units.includes(u));

  return units
    .map(u => u ? u : (emptyUnits.pop() ?? null))
    .map(unit => unit ? [`${unit} (${namesUnits.filter(u => u === unit).length})`] : null);
}

function getSecondedUnitCount(unit, personnel) {
  const names = getPersonnel(personnel);

  if (!names[0]) return null;

  return `${unit} (${names.length})`;
}

function onEdit(e) {
  const sheet = e.source.getActiveSheet();
  const value = e.range.getValue();
  const [row, column] = [e.range.getRow() - 1, e.range.getColumn() - 1];
  
  if (nonUnitSheets.includes(sheet.getName()) || typeof value !== 'string' || row > tableRows || column > tableColumns) return;

  const valueArr = value.split(separatorRegex);
  const allAvailablePersonnel = getAllAvailableUnitPersonnel(sheet);
  const allTableValues = getAllTableValues(sheet);

  if (valueArr.some(v => allAvailablePersonnel.includes(v))) {
    for (let i = 0; i < tableRows; i += 1) {
      for (let j = 0; j < tableColumns; j += 1) {
        if (i === row && j === column || typeof allTableValues[i][j] !== 'string') continue;
        
        const currentValueArr = allTableValues[i][j].split(separatorRegex);

        if (currentValueArr.some(v => valueArr.includes(v))) {
          allTableValues[i][j] = currentValueArr.filter(v => !valueArr.includes(v)).join(separator);
          sheet.getRange(i + 1, j + 1).setValue(allTableValues[i][j]);
        }
      }
    }
  }

  const allAddedPersonnel = getPersonnel(allTableValues).filter(v => allAvailablePersonnel.includes(v));
  const missedPersonnel = allAvailablePersonnel.filter(p => !allAddedPersonnel.includes(p));
  const duplicatedPersonnel = getDuplicatedPersonnel(allAddedPersonnel);

  if (missedPersonnel.length) SpreadsheetApp.getActiveSpreadsheet().toast(missedPersonnel, 'Відсутній в таблиці О/С', 10);
  if (duplicatedPersonnel.length) SpreadsheetApp.getActiveSpreadsheet().toast(duplicatedPersonnel, 'Продубльований О/С', 10);
}

function onOpen() {
  const menu = SpreadsheetApp.getUi().createMenu("Алла 💖");
  menu.addItem("Відсутній в таблиці О/С", "showMissedPersonnel");
  menu.addItem("Продубльований О/С", "showDuplicatedPersonnel");
  menu.addItem("Знайти зайвий О/С в комірці", "showRedundantPersonnel");
  menu.addSeparator();
  menu.addItem("Оновити поточний лист підрозділу", "refreshUnitSheet");
  menu.addItem("Оновити всі листи підрозділів", "refreshAllUnitSheets");
  menu.addItem("Оновити загальний список О/С", "refreshHandbookSheet");
  menu.addSeparator();
  menu.addItem("Експортувати PDF", "exportPDF");
  menu.addToUi();
}
