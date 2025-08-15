/** @param {Array} values */
const addNumbers = (values, i = 0) => values.map(([v]) => v ? ++i : null);

/** @param {Array} units */
const getUnits = units => units.filter(([u]) => u).map(([u]) => u);

/** @param {String} rank */
function getRankType(rank) {
  if (/(лейтенант|капітан|майор|полковник)/i.test(rank)) return 'Офіцер';
  if (/сержант/i.test(rank)) return 'Сержант';
  if (/солдат/i.test(rank)) return 'Солдат';
  
  return null;
}

/** 
 * @param {Array} personnel 
 * @param {Array} nativeUnits
 * @param {Array} foreignUnits 
 */
function getForeignPersonnelCountByRank(personnel, nativeUnits, foreignUnits) {
  const foreignPersonnel = personnel.filter(([unit,,,, status]) => !getUnits(nativeUnits).includes(unit) && status === 'В НАЯВНОСТІ');
  const foreignPersonnelRanks = foreignPersonnel.map(([,, rank]) => rank);
  const foreignPersonnelUnitsTo = getUnits(nativeUnits).map(unit => [unit, foreignPersonnel.filter(([, u]) => u === unit).length]);
  const foreignPersonnelUnitsFrom = getUnits(foreignUnits).map(unit => [unit, foreignPersonnel.filter(([u]) => u === unit).length]).filter(([, n]) => n > 0);
  
  return [
    [
      foreignPersonnelRanks.filter(r => getRankType(r) === 'Офіцер').length,
      foreignPersonnelRanks.filter(r => getRankType(r) === 'Сержант').length,
      foreignPersonnelRanks.filter(r => getRankType(r) === 'Солдат').length
    ],
    [null],
    ...foreignPersonnelUnitsTo,
    [null],
    ...foreignPersonnelUnitsFrom
  ];
};

function getPersonnelDataFolderLinks() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const cache = CacheService.getScriptCache();
  const root = DriveApp.getFolderById('1hrSHH2AMv_zvlQprztKCGOpyXwvUqyuq');
  const prompt = SpreadsheetApp.getUi().prompt("unitColumn nameColumn linkColumn");
  const [unitColumn, nameColumn, linkColumn] = prompt.getResponseText().split(' ');
  const rangeStart = 2;
  const cacheExpirationTime = 24 * 60 * 60;
  const units = sheet.getRange(`${unitColumn}${rangeStart}:${unitColumn}`).getValues();
  const names = sheet.getRange(`${nameColumn}${rangeStart}:${nameColumn}`).getValues();
  const emptyText = [SpreadsheetApp.newRichTextValue().setText('').build()];
  const noDataText = [SpreadsheetApp.newRichTextValue().setText('Відсутні').build()];
  const getLink = url => [SpreadsheetApp.newRichTextValue().setText('Відкрити').setLinkUrl(url).build()];
  const unitFolders = {};

  const links = units.map(([unit], i) => {
    const [name] = names[i];

    if (!unit || !name) return emptyText;
    
    const cachedUrl = cache.get(`${unit} ${name}`);

    if (cachedUrl) return getLink(cachedUrl);

    if (!unitFolders[unit] && root.getFoldersByName(unit).hasNext()) unitFolders[unit] = root.getFoldersByName(unit).next();
    
    if (!unitFolders[unit]) return noDataText;

    const personFolderSearch = unitFolders[unit].getFoldersByName(name);
    
    if (!personFolderSearch.hasNext()) return noDataText;
      
    const personFolder = personFolderSearch.next();

    if (!personFolder.getFiles().hasNext()) return noDataText;

    const personFolderUrl = personFolder.getUrl();
    cache.put(`${unit} ${name}`, personFolderUrl, cacheExpirationTime);
    
    return getLink(personFolderUrl);
  });

  sheet.getRange(`${linkColumn}${rangeStart}:${linkColumn}${rangeStart + links.length - 1}`).setRichTextValues(links);
}

function parseDate(date) {
  if (date) {
    const [dd, mm, yyyy] = date.split('.');

    return Date.parse(`${mm}/${dd}/${yyyy}`);
  }
  return Date.now();
}

/** @param {String} birthDate */
function calcAge(birthDate, date) {
  const year = 365.25 * 24 * 60 * 60 * 1000;
  const dateRegex = /\d{2}\.\d{2}\.\d{4}/;

  if (typeof birthDate !== 'string' || !dateRegex.test(birthDate) || date === '') return null;

  return Math.floor((parseDate(date) - parseDate(birthDate)) / year);
}

/** @param {Array} birthDates */
const getAgeData = birthDates => birthDates.map(([d]) => [calcAge(d)]);

const getAgeOnDate = (birthDates, dates) => birthDates.map(([birthDate], i) => [calcAge(birthDate, dates[i][0])]);

/**
 * {String} id
 * {RegExp} re
 */
function getDataFromDocById(id, re) {
  const doc = DocumentApp.openById(id);
  const body = doc.getBody();
  const data = body.getText().match(re);

  return data ? data[0] : null;
}

const getBirthDateFromDoc = id => getDataFromDocById(id, /(?<=дата народження.+)\d{1,2}.\d{1,2}.\d{2,4}/ims);

const getConscriptionDateFromDoc = id => getDataFromDocById(id, /(?<=коли призваний.+)\d{1,2}.\d{1,2}.\d{2,4}/ims);

function convertFileToGoogleDocs(root, unit, name, fileName) {
  try {
    if (!root.getFoldersByName(unit).hasNext()) return;

    const unitFolder = root.getFoldersByName(unit).next();
    const personFolderSearch = unitFolder.getFoldersByName(name);

    if (!personFolderSearch.hasNext()) return;

    const personFolder = personFolderSearch.next();
    const fileSearch = personFolder.searchFiles(`title contains "${fileName}"`);
    const existingFileSearch = personFolder.getFilesByName(`${name} ${fileName.toUpperCase()}`);

    if (!fileSearch.hasNext() || existingFileSearch.hasNext()) return;

    const questionary = fileSearch.next();
    const file = Drive.Files.create(
      {
        name: `${name} ${fileName.toUpperCase()}`,
        mimeType: MimeType.GOOGLE_DOCS
      },
      questionary.getBlob()
    );

    DriveApp.getFileById(file.id).moveTo(personFolder);

    return true;
  } catch {
    console.log(`Error converting ${fileName} for ${name}`);
  }
}

function getDataFromDocByName(root, unit, personName, fileName, extractFunc) {
  try {
    if (!root.getFoldersByName(unit).hasNext()) return null;
    
    const unitFolder = root.getFoldersByName(unit).next();
    const personFolderSearch = unitFolder.getFoldersByName(personName);

    if (!personFolderSearch.hasNext()) return null;

    const personFolder = personFolderSearch.next();
    const convertedFileSearch = personFolder.getFilesByName(`${personName} ${fileName.toUpperCase()}`);

    if (!convertedFileSearch.hasNext()) return null;
    
    return extractFunc(convertedFileSearch.next().getId());
  } catch {
    console.log(`Error getting data for ${personName}`);
    return null;
  }
}

function getBirthDates() {
  const root = DriveApp.getFolderById('1hrSHH2AMv_zvlQprztKCGOpyXwvUqyuq');
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const prompt = SpreadsheetApp.getUi().prompt("unitColumn nameColumn dateColumn rangeStart rangeEnd");
  const [unitColumn, nameColumn, datesColumn, rangeStart, rangeEnd] = prompt.getResponseText().split(' ');
  const dates = [];

  for (let i = +rangeStart; i <= +rangeEnd; i += 1) {
    const unit = sheet.getRange(`${unitColumn}${i}`).getValue();
    const name = sheet.getRange(`${nameColumn}${i}`).getValue();

    if (!name) {
      dates.push([null]);
      continue;
    }

    const date = getDataFromDocByName(root, unit, name, 'анкета', getBirthDateFromDoc); 

    dates.push([date]);

    if (date) SpreadsheetApp.getActiveSpreadsheet().toast(name, 'Отримана дата народження', 3);
  }

  sheet.getRange(`${datesColumn}${rangeStart}:${datesColumn}${rangeEnd}`).setValues(dates);
}

function getConscriptionDates() {
  const root = DriveApp.getFolderById('1hrSHH2AMv_zvlQprztKCGOpyXwvUqyuq');
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const prompt = SpreadsheetApp.getUi().prompt("unitColumn nameColumn dateColumn rangeStart rangeEnd");
  const [unitColumn, nameColumn, datesColumn, rangeStart, rangeEnd] = prompt.getResponseText().split(' ');
  const dates = [];

  for (let i = +rangeStart; i <= +rangeEnd; i += 1) {
    const unit = sheet.getRange(`${unitColumn}${i}`).getValue();
    const name = sheet.getRange(`${nameColumn}${i}`).getValue();

    if (!name) {
      dates.push([null]);
      continue;
    }

    const date = getDataFromDocByName(root, unit, name, 'анкета', getConscriptionDateFromDoc); 

    dates.push([date]);

    if (date) SpreadsheetApp.getActiveSpreadsheet().toast(name, 'Отримана дата призову', 3);
  }

  sheet.getRange(`${datesColumn}${rangeStart}:${datesColumn}${rangeEnd}`).setValues(dates);
}

function convertToGoogleDocs() {
  const root = DriveApp.getFolderById('1hrSHH2AMv_zvlQprztKCGOpyXwvUqyuq');
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const prompt = SpreadsheetApp.getUi().prompt("unitColumn nameColumn fileName rangeStart rangeEnd");
  const [unitColumn, nameColumn, fileName, rangeStart, rangeEnd] = prompt.getResponseText().split(' ');

  for (let i = +rangeStart; i <= +rangeEnd; i += 1) {
    const unit = sheet.getRange(`${unitColumn}${i}`).getValue();
    const name = sheet.getRange(`${nameColumn}${i}`).getValue();

    if (!name) continue;

    if (convertFileToGoogleDocs(root, unit, name, fileName)) SpreadsheetApp.getActiveSpreadsheet().toast(name, 'Успішна конвертація', 3);
  }
}

function exportPDF() {
  const id = SpreadsheetApp.getActiveSpreadsheet().getId();
  const sheetId = '1189103823';
  const exportFolderId = '1H2kQbBvMUkZGlsd5Ue76ON36arQsRUa4';
  const options = `gid=${sheetId}&format=pdf&portrait=false&size=a4&scale=2&left_margin=0.5&right_margin=0.5&top_margin=0.5&bottom_margin=0.5&printnotes=false`;
  const url = `https://docs.google.com/spreadsheets/d/${id}/export?${options}`;
  const date = new Date(Date.now()).toLocaleString('uk-UA', { day: 'numeric', month: 'numeric', year: 'numeric' });
  const time = new Date(Date.now()).toLocaleTimeString('uk-UA', { hour: '2-digit', minute: '2-digit' });
  
  const blob = getFileAsBlob(url);
  blob.setName(`ШТАТКА ${date} ${time}.pdf`);

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

const getAllBirthDates = () => SpreadsheetApp
  .getActiveSpreadsheet()
  .getSheetByName('Штатка')
  .getRange('D3:K')
  .getValues()
  .filter(([,,,,r]) => r)
  .map(([rank1,,, position, rank2, name,, birthDate]) => ({
    'Звання по штату': rank1,
    'Посада': position,
    'Військове звання': rank2,
    'ПІБ': name,
    'Дата народження': birthDate
  }));

function getDaysUntilBirthday(birthDate) {
  const today = new Date();
  const currentYear = today.getFullYear();
  const currentMonth = today.getMonth();
  const currentDate = today.getDate();

  const parts = birthDate.split('.');
  const birthDay = +parts[0];
  const birthMonth = parts[1] - 1;

  let nextBirthdayYear = currentYear;

  if (birthMonth < currentMonth || (birthMonth === currentMonth && birthDay <= currentDate)) nextBirthdayYear += 1;

  const nextBirthday = new Date(nextBirthdayYear, birthMonth, birthDay);
  const timeDifference = nextBirthday.getTime() - today.getTime();

  return Math.ceil(timeDifference / (1000 * 60 * 60 * 24));
}

function sergeantsBirthdayNotification() {
  const sergeants = getAllBirthDates().filter(p => getRankType(p['Звання по штату']) == 'Сержант' && getDaysUntilBirthday(p['Дата народження']) == 1);

  if (!sergeants.length) return;

  MailApp.sendEmail(
    'oleksiisedun@gmail.com,allatripolskaa9@gmail.com,deniro198666@gmail.com',
    'Прийдешні дні народження сержантів 🎉',
    sergeants.map(s => `${s['Військове звання']} ${s['ПІБ']} ${s['Дата народження']}`).join('\n')
  );
}

function officersBirthdayNotification() {
  const officers = getAllBirthDates().filter(p => getRankType(p['Звання по штату']) == 'Офіцер' && getDaysUntilBirthday(p['Дата народження']) == 1);

  if (!officers.length) return;

  MailApp.sendEmail(
    'oleksiisedun@gmail.com,prochuhani@gmail.com,allatripolskaa9@gmail.com',
    'Прийдешні дні народження офіцерів 🎉',
    officers.map(s => `${s['Військове звання']} ${s['ПІБ']} ${s['Дата народження']}`).join('\n')
  );
}

function hasStateAward(data) {
  const stateAwards = [
    'Герой України', 
    'орден Золота Зірка', 
    'Богдана Хмельницького', 
    'Орден за мужність', 
    'Орден княгині Ольги', 
    'Данила Галицького',
    'За військову службу Україні',
    'За бездоганну службу',
    'Захиснику Вітчизни',
    'Іменна вогнепальна зброя',
    'Хрест бойових заслуг',
    'За оборону України'
  ];

  const isStateAward = award => stateAwards.some(stateAward => award.toLowerCase().includes(stateAward.toLowerCase()));

  return data.map(awards => awards.some(award => isStateAward(award)) ? 'Має державну нагороду' : null);
}

function onOpen() {
  const menu = SpreadsheetApp.getUi().createMenu("Алла 💓");
  menu.addItem("Оновити установчі дані О/С", "getPersonnelDataFolderLinks");
  menu.addSeparator();
  menu.addItem("Отримати дати народження з анкет", "getBirthDates");
  menu.addItem("Отримати дати призову з анкет", "getConscriptionDates");
  menu.addSeparator();
  menu.addItem("Конвертувати установчі дані О/С", "convertToGoogleDocs");
  menu.addSeparator();
  menu.addItem("Експортувати PDF", "exportPDF");
  menu.addToUi();
}
