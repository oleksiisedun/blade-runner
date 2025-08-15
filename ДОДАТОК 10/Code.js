/** @param {Array} values */
const addNumbers = (values, i = 0) => values.map(([v]) => v ? ++i : null);

/** @param {String} range */
const getValues = (range, sheet = SpreadsheetApp.getActiveSheet()) => sheet.getRange(range).getValues().map(([v]) => v).filter(v => v);

function showDuplicatedNames() {
  const names = getValues('E7:E');
  const duplicatedNames = names.map(name => name.trim().toUpperCase()).filter((name, i, names) => names.lastIndexOf(name) !== i);

  SpreadsheetApp.getUi().alert(duplicatedNames.length ? duplicatedNames.join('\n') : '💪');
}

function showMissedNames() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const activeSheetName = sheet.getName();
  const allNames = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Handbook').getRange('A2:B').getValues().filter(([unit]) => {
    if (!unit) return false;
    if (activeSheetName === '3 БОП' || unit.includes(activeSheetName) || activeSheetName.includes(unit)) return true;
    return false;
  }).map(([, name]) => name);
  const names = getValues('E10:E', sheet).map(n => n.trim().toLowerCase());
  const missedNames = allNames.filter(n => !names.includes(n.toLowerCase()));

  SpreadsheetApp.getUi().alert(missedNames.length ? missedNames.join('\n') : '💪');
}

function showForeignNames() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const activeSheetName = sheet.getName();
  const allNames = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Handbook').getRange('A2:B').getValues().filter(([unit]) => {
    if (!unit) return false;
    if (activeSheetName === '3 БОП' || unit.includes(activeSheetName) || activeSheetName.includes(unit)) return true;
    return false;
  }).map(([, name]) => name.toUpperCase());
  const names = getValues('E7:E', sheet).map(n => n.trim().toUpperCase());
  const foreignNames = names.filter(name => !allNames.includes(name));

  SpreadsheetApp.getUi().alert(foreignNames.length ? foreignNames.join('\n') : '💪');
} 

function getFileAsBlob(url) {
  const response = UrlFetchApp.fetch(url, {
    muteHttpExceptions: true,
    headers: { Authorization: `Bearer ${ScriptApp.getOAuthToken()}` }
  });

  return response.getBlob();
}

function exportXLSX() {
  const sheetsToRemove = ['УПР, Штаб', '7 РОП', '8 РОП', '9 РОП', 'МБ', 'РВП', 'ШР', 'РМТЗ', 'ВЗ', 'МП', 'Handbook'];
  const ranges = ['A10:K933', 'N10:EW933', 'EY10:FT933'];
  const donorIds = [
    '1IhnAlOeeUTbAMUpI7rzW7_IG4pSiBQWSk_VIR5k3LqM',
    '1R35WvC0_1kMqAyQNI80EgzSXzouVFIQ1d87jSKepERE',
    '1sgN5jQ4gfvdwiVfgV2ojEOrRFobVLAnsqWkgeoj2_vw',
    '1iQuu9aB7TWW06fgtTWn7ZLZtnMQ-RpGRc0ObGxLuht4',
    '1Zr4F0MJzqKRLev60yiiOe3j_5JJrqmzkCSHINiLEY50',
    '1cd7lkQl0_lBIIF7EhFdgi4fcARAZu755cfmPTGicu3A',
    '1Tqhx-q27atsjGNTE8fzSxY_5v59JnIK8VVbFNWj7HD8',
    '1G7h2JQ76govdwmnyB9nzvtLvLyW5jp5M2--tEYmPbeQ',
    '1Wue6-VhTPFd0Dmiwyj06cBm_kaJ4DTMIg18bbiZcDFk',
    '1Uql7OFzfLfRPmqnGjmCgaNnNiDhEDGh068UhLoORNCY',
    '1h3KlxibtHMPf_Hjab4n6_jU3XWAFfep_yBpy5Wnh4Tw',
    '1AJPFrFvbOJ5Q8lW37ZY5hSSfQEzLSeA29H50qNUIVko',
    '17YAnBNboYkxHdcd0xaXxfMb3MpnolcHFQj50fRA42Ak',
    '11iq0MfJLunapcdgdOdVMBBY1Z0mwvuKPMEFP8sgDCp8',
    '1iRL7A3YHrS2KB6rgPVAOBO7pI6PmOcv3WsydWn4B6V0',
    '1XVjdtHHHJjFwKJ4xOOyy4iHHP0aUySw_snwMAZw7oJI',
    '1wyeHhUycf4QVVDcyeK-MRdk6-6ZH8XEkpKEKlIq-MZI',
    '1uvTu3tSvElztW6pkLl4p00p9YRUi-nDHVNNEe6eiaN8'
  ];
  const exportFolder = DriveApp.getFolderById('1cPi7wQA6IcdrLih4pRgvVV15zjoanKVI');
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet().copy('ДОДАТОК 10 TEMP');
  const id = spreadsheet.getId();
  const url = `https://docs.google.com/spreadsheets/d/${id}/export?format=xlsx`;
  const date = new Date(Date.now()).toLocaleString('uk-UA', { day: 'numeric', month: 'numeric', year: 'numeric' });
  const time = new Date(Date.now()).toLocaleTimeString('uk-UA', { hour: '2-digit', minute: '2-digit' });

  donorIds.forEach(donorId => addImportrangePermission(id, donorId));

  spreadsheet.getSheets().forEach(sheet => {
    if (!sheetsToRemove.includes(sheet.getName())) {
      sheet.setFrozenRows(0);
      sheet.setFrozenColumns(0);

      ranges.forEach(range => {
        const dataRange = sheet.getRange(range);
        dataRange.copyTo(dataRange, { contentsOnly: true });
      });

      SpreadsheetApp.flush();
    }
  });

  spreadsheet.getSheets().forEach(sheet => {
    if (sheetsToRemove.includes(sheet.getName())) {
      spreadsheet.deleteSheet(sheet);
    }
  });

  const blob = getFileAsBlob(url);

  blob.setName(`ДОДАТОК 10 ${date} ${time}.xlsx`);
  exportFolder.createFile(blob);
  DriveApp.getFileById(id).setTrashed(true);
}

function addImportrangePermission(id, donorId) {
  const url = `https://docs.google.com/spreadsheets/d/${id}/externaldata/addimportrangepermissions?donorDocId=${donorId}`;
  const params = {
    method: 'post',
    headers: { Authorization: `Bearer ${ScriptApp.getOAuthToken()}` },
    muteHttpExceptions: true
  };
  
  UrlFetchApp.fetch(url, params);
  SpreadsheetApp.flush();
}

function showPositionsCount() {
  const positions = getValues('FS10:FS');
  const positionsCount = positions.reduce((acc, p) => ({ ...acc, [p]: (acc[p] || 0) + 1 }), {});
  const positionsTable = Object.keys(positionsCount).map(key => `${key}: ${positionsCount[key]}`).join('\n');
  SpreadsheetApp.getUi().alert(positionsTable);
}

function onOpen() {
  const menu = SpreadsheetApp.getUi().createMenu("Алла 💗");
  
  menu.addItem("Розрахунок позицій", "showPositionsCount");
  menu.addSeparator();
  menu.addItem("Пошук дублікатів імен в/с", "showDuplicatedNames");
  menu.addItem("Пошук пропущених імен в/с", "showMissedNames");
  menu.addItem("Пошук зайвих імен в/с", "showForeignNames");
  menu.addSeparator();
  menu.addItem("Експортувати XLSX", "exportXLSX");
  menu.addToUi();
}
