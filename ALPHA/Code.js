const MAX_COLUMN_HEIGHT = 52;
const emptyRow = [null, null, null, null, null, null, null, null];

const formatPosition = position => [null, null, null, position, null, null, null, null];

const formatPersonnel = ([number, callsign, date, name, radio, weapon1, weapon2, unit]) => [
  number, 
  unit, 
  callsign.toUpperCase(), 
  date, 
  name, 
  `${radio}`.replace(/[,\s]+/g, ' ').toUpperCase(),
  weapon1,
  weapon2
];

const formatPositionWithPersonnel = ([position, personnel]) => [formatPosition(position), ...personnel.map(p => formatPersonnel(p))];

/** @param {Array} data */
function getPositionsWithPersonnel(data) {
  const records = data.filter(([p]) => p);
  const positions = [...new Set(records.map(([p]) => p))];
  const positionsWithPersonnel = positions.map(p => [p, records.filter(r => r[0] === p).map((r, i) => [i + 1, ...r.slice(1)])]);
  
  return positionsWithPersonnel
    .reduce((acc, p) => {
      const position = formatPositionWithPersonnel(p);
      if (acc.at(-1).length + position.length > MAX_COLUMN_HEIGHT) acc.push([]);
      acc.at(-1).push(...position);
      return acc;
    }, [[]])
    .map(c => {
      while(c.length < MAX_COLUMN_HEIGHT) c.push(emptyRow);
      return c;
    })
    .reduce((acc, c, i) => i ? acc.map((e, i) => [...e, ...(c[i] ?? [null])]) : [...c], []);
}

/** 
 * @param {Array} data
 * @param {Array} order 
 */
function getPositionsWithPersonnelS3(data, order) {
  const getPositionShortName = name => (name.match(/(?<=[«"]).+?(?=[»"])/) ?? [])[0] ?? name;
  const records = data.filter(([p]) => p);
  const positions = [...new Set(records.map(([p]) => p))];
  const positionsOrder = order.filter(([p]) => p).map(([p]) => getPositionShortName(p));
  const positionsWithPersonnel = positions.map(p => [p, records.filter(r => r[0] === p).map((r, i) => [i + 1, ...r.slice(1)])]);

  return [
    ...positionsWithPersonnel
      .filter(([p]) => positionsOrder.includes(getPositionShortName(p)))
      .sort((p1, p2) => positionsOrder.indexOf(getPositionShortName(p1[0])) - positionsOrder.indexOf(getPositionShortName(p2[0])))
      .reduce((acc, p) => [...acc, ...formatPositionWithPersonnel(p)], []),
    ...positionsWithPersonnel
      .filter(([p]) => !positionsOrder.includes(getPositionShortName(p)))
      .reduce((acc, p) => [...acc, ...formatPositionWithPersonnel(p)], [])
  ];
}

function exportPDF() {
  const id = SpreadsheetApp.getActiveSpreadsheet().getId();
  const url = `https://docs.google.com/spreadsheets/d/${id}/export?format=pdf&gid=707226597&size=a4&left_margin=0.5&right_margin=0.5&top_margin=0.5&bottom_margin=0.5`;
  const blob = getFileAsBlob(url);
  const exportFolder = DriveApp.getFolderById('1FyMdYBmQ4_I0bzqoYl6XswLQj0tpPbh4');
  const date = new Date(Date.now()).toLocaleString('uk-UA', { day: 'numeric', month: 'numeric', year: 'numeric' });
  const time = new Date(Date.now()).toLocaleTimeString('uk-UA', { hour: '2-digit', minute: '2-digit' });
  
  blob.setName(`ALPHA S-3 ${date} ${time}`);
  exportFolder.createFile(blob);
}

function getFileAsBlob(url) {
  const response = UrlFetchApp.fetch(url, {
    muteHttpExceptions: true,
    headers: {
      Authorization: `Bearer ${ScriptApp.getOAuthToken()}`,
    },
  });

  return response.getBlob();
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
  SpreadsheetApp.flush();

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

function refreshAlpha() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('3 БОП');
  const lock = LockService.getDocumentLock();
  
  lock.waitLock(10 * 1000);
  refreshRanges(sheet, 'B2');
  lock.releaseLock();
}

function onOpen() {
  const menu = SpreadsheetApp.getUi().createMenu("Макс S-3 💪");
  menu.addItem("Експортувати PDF", "exportPDF");
  menu.addToUi();
}
