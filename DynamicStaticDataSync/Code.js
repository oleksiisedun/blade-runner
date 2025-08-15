/** @type {PropertiesService.Properties} */
const scriptProperties = PropertiesService.getScriptProperties();

/** @param {String} key */
const getPropertiesStore = key => JSON.parse(scriptProperties.getProperty(key) ?? "{}");

/** 
 * @param {String} key 
 * @param {Object} store
 */
const setPropertiesStore = (key, store) => scriptProperties.setProperty(key, JSON.stringify(store));

/** 
 * @param {String} storeKey 
 * @param {String} propertyKey 
 */
const getStoreProperty = (storeKey, propertyKey) => getPropertiesStore(storeKey)[propertyKey];

/** 
 * @param {String} storeKey 
 * @param {Object} property
 */
const setStoreProperty = (storeKey, property) => setPropertiesStore(storeKey, { ...getPropertiesStore(storeKey), ...property });

/** @param {Array} data */
const getId = (...data) => data.join(' ').replace(/\s+/g, '-').replace(/["'«»,\.]/g, '').toLowerCase();

/** @param {String} columnLetter */
function getColumnNumber(columnLetter) {
  let columnNumber = 0;

  for (let i = 0; i < columnLetter.length; i += 1) {
    columnNumber *= 26;
    columnNumber += columnLetter.charCodeAt(i) - 'A'.charCodeAt() + 1;
  }

  return columnNumber;
}

/**
 * @param {Object} options
 * @param {String} options.sheetName
 * @param {String[]} options.idColumns
 * @param {String[]} options.targetColumns
 */
function alignColumns(options) {
  const { sheetName, idColumns, targetColumns } = options;
  const headerRowsCount = 1;
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  const dataRange = sheet.getDataRange();
  const dataRangeValues = dataRange.getValues().slice(headerRowsCount);
  const store = getPropertiesStore(sheetName);
  const ids = dataRangeValues.map(row => getId(...idColumns.map(column => row[getColumnNumber(column) - 1])));
  
  const lock = LockService.getDocumentLock();
  lock.waitLock(10 * 1000);

  targetColumns.forEach(column => {
    const columnValues = ids.map(id => {
      const columnId = `${id}-${column}`;

      return [columnId in store ? store[columnId] : null]; 
    });

    sheet.getRange(`${column}${1 + headerRowsCount}:${column}${dataRangeValues.length + headerRowsCount}`).setValues(columnValues);
  });

  lock.releaseLock();
}

/**
 * @param {Object} options
 * @param {SpreadsheetApp.Sheet} options.sheet
 * @param {String[]} options.idColumns
 * @param {String} options.column
 * @param {String|Number} options.row
 * @param {String} options.value
 */
function saveColumnValue(options) {
  const { sheet, idColumns, column, row, value } = options;
  const id = getId(...idColumns.map(column => sheet.getRange(`${column}${row}`).getValue()));

  setStoreProperty(sheet.getName(), { [`${id}-${column}`]: value });
}

/**
 * @typedef {Object} EditEventData
 * @property {SpreadsheetApp.Sheet} sheet
 * @property {String|Number} value
 * @property {String} column
 * @property {String|Number} row
 * 
 * @param {Event} e
 * @returns {EditEventData}
 */
function getEditEventData(e) {
  const [column, row] = e.range.getA1Notation().match(/[A-Z]+|\d+/g);
  
  return { 
    sheet: e.source.getActiveSheet(),
    value: e.value,
    column,
    row
  };
}

/**
 * @param {Event} e
 */
function onEdit(e) {
  const { sheet, value, column, row } = getEditEventData(e);
  const idColumns = ['E', 'F', 'C', 'G', 'H', 'K'];
  const targetSheetName = 'Sheet4';

  if (sheet.getName() === targetSheetName) {
    saveColumnValue({ sheet, idColumns, column, row, value });
  } else {
    alignColumns({ sheetName: targetSheetName, idColumns, targetColumns: ['N', 'O', 'P', 'Q', 'R', 'S'] });
  }
}

/** @param {String} date */
function parseDate(date) {
  const [dd, mm, yyyy] = date.split('.');

  return Date.parse(`${mm}/${dd}/${yyyy}`);
}

/** 
 * @param {Array} data 
 * @param {Number} dateColumn
 */
const sortDataByDate = (data, dateColumn) => data.filter(row => row[dateColumn - 1]).toSorted((a, b) => parseDate(a[dateColumn - 1]) - parseDate(b[dateColumn - 1]));
