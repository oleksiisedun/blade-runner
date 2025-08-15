const units = ['УПР', 'Штаб', '7 РОП', '8 РОП', '9 РОП', 'МБ', 'РВП', 'ШР', 'РМТЗ', 'ВЗ', 'МП'];
const exportFolderId = '1n3-QRktxzZkTEgv69JiAgsIZvbGBN4zx';

const getIdFromLink = link => (link.match(/[\w\d-]{20,}/) ?? [''])[0];

function parseDate(date) {
  const [dd, mm, yyyy] = date.split('.');

  return Date.parse(`${mm}/${dd}/${yyyy}`);
}

/** @param {String} startDate */
function calcServiceTime(startDate) {
  const year = 365.25 * 24 * 60 * 60 * 1000;
  const dateRegex = /\d{2}\.\d{2}\.\d{4}/;

  if (typeof startDate !== 'string' || !dateRegex.test(startDate)) return null;

  return Math.round((Date.now() - parseDate(startDate)) / year);
}

function getPersonInfo(sheet) {
  const values = sheet.getRange(3, 2, 32, 13).getValues();
  const getCloseRelative = (...values) => ({
    'Ступінь родинних зв’язків': values[0],
    'ПІБ': values[1],
    'Місце проживання/адреса': values[2],
    'Телефон': values[3],
    'Дата народження': values[4],
    'Місце роботи': values[5]
  });
  const getServiceRecord = (...values) => ({
    'З якого по який час': values[0],
    'Найменування посади': values[1]
  });

  return {
    'ПІБ': values[0][1],
    'Попередні ПІБ': values[1][1],
    'Дата народження': values[2][1],
    'Місце народження': values[3][1],
    'ІПН': values[4][1],
    'Номер телефону': values[5][1],
    'Сімейний стан': values[6][1],
    'Наявність житла': values[7][1],
    'Адреса фактичного проживання': values[8][1],
    'Адреса реєстрації': values[9][1],
    'Підрозділ': values[10][1],
    'Військове звання': values[11][1],
    'Серія, № в/квитка (посвідчення)': values[12][1],
    'Номер жетону': values[13][1],
    'Ким призваний': values[14][1],
    'Дата призову': values[15][1],
    'Контракт укладено до': values[16][1],
    'УБД серія, номер': values[17][1],
    'УБД дата видачі': values[18][1],
    'Зброя 1 тип, номер': values[19][1],
    'Зброя 2 тип, номер': values[20][1],
    'Військова освіта': values[21][1],
    'Цивільна освіта': values[22][1],
    'Володіння мовами': values[23][1],
    'Науковий ступінь, вчене звання': values[24][1],
    'Нагороди, почесні звання': values[25][1],
    'Стягнення': values[26][1],
    'В яких регіонах бажає служити': values[27][1],
    'Ким працював раніше, стаж': values[28][1],
    'Посвідчення водія (категорії)': values[29][1],
    'Фото': values[0][11],
    'Фото Паспорт': values[1][11],
    'Фото ІПН': values[2][11],
    'Фото ВК': values[3][11],
    'Фото УБД': values[4][11],
    'Національність': values[7][11],
    'Громадянство': values[8][11],
    'Наявність судимостей': values[9][11],
    'СЗЧ, дата': values[10][11],
    'Словесний портрет': values[11][11],
    'Особливі прикмети': values[12][11],
    'Примітки': values[13][11],
    'Близькі родичі': [
      getCloseRelative(values[1][3], values[1][4], values[1][5], values[1][6], values[1][7], values[1][8]),
      getCloseRelative(values[2][3], values[2][4], values[2][5], values[2][6], values[2][7], values[2][8]),
      getCloseRelative(values[3][3], values[3][4], values[3][5], values[3][6], values[3][7], values[3][8]),
      getCloseRelative(values[4][3], values[4][4], values[4][5], values[4][6], values[4][7], values[4][8]),
      getCloseRelative(values[5][3], values[5][4], values[5][5], values[5][6], values[5][7], values[5][8]),
      getCloseRelative(values[6][3], values[6][4], values[6][5], values[6][6], values[6][7], values[6][8]),
      getCloseRelative(values[7][3], values[7][4], values[7][5], values[7][6], values[7][7], values[7][8]),
      getCloseRelative(values[8][3], values[8][4], values[8][5], values[8][6], values[8][7], values[8][8])
    ],
    'Проходження служби': [
      getServiceRecord(values[12][3], values[12][4]),
      getServiceRecord(values[13][3], values[13][4]),
      getServiceRecord(values[14][3], values[14][4]),
      getServiceRecord(values[15][3], values[15][4]),
      getServiceRecord(values[16][3], values[16][4]),
      getServiceRecord(values[17][3], values[17][4]),
      getServiceRecord(values[18][3], values[18][4]),
      getServiceRecord(values[19][3], values[19][4]),
      getServiceRecord(values[20][3], values[20][4]),
      getServiceRecord(values[21][3], values[21][4]),
      getServiceRecord(values[22][3], values[22][4]),
      getServiceRecord(values[23][3], values[23][4]),
      getServiceRecord(values[24][3], values[24][4]),
      getServiceRecord(values[25][3], values[25][4]),
      getServiceRecord(values[26][3], values[26][4])
    ]
  };
}

function getAllPersonsInfo() {
  const allPersonsInfo = [];

  for (const sheet of SpreadsheetApp.getActiveSpreadsheet().getSheets()) {
    const sheetName = sheet.getName();
    
    if (sheetName == 'ШАБЛОН' || units.includes(sheetName)) continue;

    allPersonsInfo.push(getPersonInfo(sheet));
  }
  
  return allPersonsInfo;
}

function getAllPersonsInfoTable() {
  const allPersonsInfo = getAllPersonsInfo();
  const mainTableColumns = Object.keys(allPersonsInfo[0]);
  const table = [mainTableColumns];
  const mergeData = data => data.map(r => Object.values(r)).reduce((acc, r) => r[1] ? `${acc}${r.join(' ').trim()}\n` : acc, '');

  allPersonsInfo.forEach(personInfo => {
    const personRow = mainTableColumns.slice(0, -2).map(column => personInfo[column]);

    personRow.push(mergeData(personInfo['Близькі родичі']));
    personRow.push(mergeData(personInfo['Проходження служби']));

    table.push(personRow)
  });

  return table;
}

function reloadTable() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets().find(sheet => units.includes(sheet.getName()));
  const range = sheet.getRange('A1');
  const formula = range.getFormula();

  range.clearContent();
  SpreadsheetApp.flush();
  range.setFormula(formula);
}

function resizeImage(image, targetHeight) {
  const height = image.getHeight();
  const width = image.getWidth();

  image.setHeight(targetHeight).setWidth(targetHeight / (height / width));
}

function exportWantedCard() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const sheetName = sheet.getName();

  if (sheetName == 'ШАБЛОН' || units.includes(sheetName)) return;

  const wantedCardData = [
    'Військове звання',
    'ПІБ',
    'Дата народження',
    'Місце народження',
    'Національність',
    'Ким призваний',
    'Дата призову',
    'Адреса фактичного проживання',
    'Наявність судимостей',
    'Словесний портрет',
    'Особливі прикмети'
  ];
  const personInfo = getPersonInfo(sheet);
  const templateId = '19HanFqGmfFFHvwKXZQPH-RA3EpgeX1cIBt1sM3z59oA';
  const photoId = (personInfo['Фото'].match(/[\w\d-]{20,}/) ?? [''])[0];
  const relative = Object.values(personInfo['Близькі родичі'].find(data => data['Телефон'])).slice(0, 4).join(', ');
  const photo = DriveApp.getFileById(photoId).getBlob();
  const exportFolder = DriveApp.getFolderById(exportFolderId);
  const wantedCardId = DriveApp.getFileById(templateId).makeCopy(`РК ${sheetName}`, exportFolder).getId();
  const wantedCard = DocumentApp.openById(wantedCardId);
  const body = wantedCard.getBody();

  wantedCardData.forEach(data => body.replaceText(`{${data}}`, personInfo[data]));
  body.replaceText('{Родич}', relative);
  body.getParagraphs()[1].insertInlineImage(0, photo);
  resizeImage(body.getImages()[0], 400);
}

function exportF1() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const sheetName = sheet.getName();

  if (sheetName == 'ШАБЛОН' || units.includes(sheetName)) return;

  const f1Data = [
    'Військове звання',
    'ПІБ',
    'Дата народження',
    'Дата призову',
    'Номер жетону',
    'Місце народження',
    'Контракт укладено до',
    'Військова освіта',
    'Цивільна освіта',
    'Володіння мовами',
    'Науковий ступінь, вчене звання',
    'Нагороди, почесні звання',
    'Стягнення',
    'Сімейний стан',
    'Наявність житла',
    'В яких регіонах бажає служити',
    'Ким працював раніше, стаж',
    'Сімейний стан'
  ];
  const personInfo = getPersonInfo(sheet);
  const templateId = '1yW1ieQdmctmoSY-AtPK_H9tJq5jk_u11QoSY-0tCLRI';
  const photoId = getIdFromLink(personInfo['Фото']);
  const photo = DriveApp.getFileById(photoId).getBlob();
  const exportFolder = DriveApp.getFolderById(exportFolderId);
  const f1Id = DriveApp.getFileById(templateId).makeCopy(`Ф-1 ${sheetName}`, exportFolder).getId();
  const f1 = DocumentApp.openById(f1Id);
  const body = f1.getBody();

  const currentPosition = personInfo['Проходження служби'].findLast(data => data['Найменування посади'])['Найменування посади'];
  const currentPositionDate = personInfo['Проходження служби'].findLast(data => data['З якого по який час'])['З якого по який час'].match(/^\d{2}\.\d{2}\.\d{4}/)[0];
  const totalServiceTime = calcServiceTime(personInfo['Дата призову']);
  const childrenPhones = personInfo['Близькі родичі']
    .filter(data => data['Ступінь родинних зв’язків'].startsWith('дитина') && data['Телефон'])
    .map(data => `${data['ПІБ']} ${data['Телефон']}`)
    .join('; ');
  const otherPhones = personInfo['Близькі родичі']
    .slice(-3)
    .filter(data => !/дружина|чоловік/i.test(data['Ступінь родинних зв’язків']) && data['Телефон'])
    .map(data => `${data['ПІБ']} ${data['Телефон']}`)
    .join('; ');

  f1Data.forEach(data => body.replaceText(`{${data}}`, personInfo[data]));
  body.replaceText('{Актуальна посада}', currentPosition);
  body.replaceText('{Дата призначення на посаду}', currentPositionDate);
  body.replaceText('{Загальний стаж служби}', totalServiceTime);
  
  personInfo['Проходження служби'].forEach((data, i) => {
    if (i > 9) return;

    body.replaceText(`{Період ${i}}`, data['З якого по який час']);
    body.replaceText(`{Посада ${i}}`, data['Найменування посади']);
  });

  body.replaceText('{ПІБ матері}', personInfo['Близькі родичі'][0]['ПІБ']);
  body.replaceText('{Адреса фактичного проживання матері}', personInfo['Близькі родичі'][0]['Місце проживання/адреса']);
  body.replaceText('{Номер телефона матері}', personInfo['Близькі родичі'][0]['Телефон']);
  body.replaceText('{Дата народження матері}', personInfo['Близькі родичі'][0]['Дата народження']);
  body.replaceText('{Місце роботи матері}', personInfo['Близькі родичі'][0]['Місце роботи']);

  body.replaceText('{ПІБ батька}', personInfo['Близькі родичі'][1]['ПІБ']);
  body.replaceText('{Адреса фактичного проживання батька}', personInfo['Близькі родичі'][1]['Місце проживання/адреса']);
  body.replaceText('{Номер телефона батька}', personInfo['Близькі родичі'][1]['Телефон']);
  body.replaceText('{Дата народження батька}', personInfo['Близькі родичі'][1]['Дата народження']);
  body.replaceText('{Місце роботи батька}', personInfo['Близькі родичі'][1]['Місце роботи']);

  body.replaceText('{ПІБ дитини 1}', personInfo['Близькі родичі'][2]['ПІБ']);
  body.replaceText('{Дата народження дитини 1}', personInfo['Близькі родичі'][2]['Дата народження']);
  body.replaceText('{ПІБ дитини 2}', personInfo['Близькі родичі'][3]['ПІБ']);
  body.replaceText('{Дата народження дитини 2}', personInfo['Близькі родичі'][3]['Дата народження']);
  body.replaceText('{ПІБ дитини 3}', personInfo['Близькі родичі'][4]['ПІБ']);
  body.replaceText('{Дата народження дитини 3}', personInfo['Близькі родичі'][4]['Дата народження']);

  const partner = personInfo['Близькі родичі'].find(data => /дружина|чоловік/i.test(data['Ступінь родинних зв’язків']));
  body.replaceText('{ПІБ дружини/чоловіка}', partner ? partner['ПІБ'] : '');
  body.replaceText('{Номер телефона дружини/чоловіка}', partner ? partner['Телефон'] : '');
  body.replaceText('{Номери телефонів дітей}', childrenPhones);
  body.replaceText('{Номери телефонів інших родичів}', otherPhones);

  body.getParagraphs()[2].insertInlineImage(0, photo);
  resizeImage(body.getImages()[0], 400);
}

function exportQuestionary() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const sheetName = sheet.getName();

  if (sheetName == 'ШАБЛОН' || units.includes(sheetName)) return;

  const questionaryData = [
    'Військове звання',
    'ПІБ',
    'Дата народження',
    'Дата призову',
    'Номер жетону',
    'Підрозділ',
    'Ким призваний',
    'Зброя 1 тип, номер',
    'Зброя 2 тип, номер',
    'Адреса фактичного проживання',
    'Номер телефону'
  ];
  const personInfo = getPersonInfo(sheet);
  const templateId = '1C1JGPFI78kiFAmIkfcvIevq_Rq4Kna1FC-x_LkiT4Ps';
  const photoId = getIdFromLink(personInfo['Фото']);
  const photo = DriveApp.getFileById(photoId).getBlob();
  const exportFolder = DriveApp.getFolderById(exportFolderId);
  const questionaryId = DriveApp.getFileById(templateId).makeCopy(`Анкета ${sheetName}`, exportFolder).getId();
  const questionary = DocumentApp.openById(questionaryId);
  const body = questionary.getBody();

  const relatives = personInfo['Близькі родичі'].filter(data => data['Телефон']);

  questionaryData.forEach(data => body.replaceText(`{${data}}`, personInfo[data]));

  for (let i = 0; i < 5; i += 1) {
    body.replaceText(`{Ступінь родинних зв’язків ${i}}`, relatives[i] ? relatives[i]['Ступінь родинних зв’язків'] : '');
    body.replaceText(`{ПІБ ${i}}`, relatives[i] ? relatives[i]['ПІБ'] : '');
    body.replaceText(`{Місце проживання/адреса ${i}}`, relatives[i] ? relatives[i]['Місце проживання/адреса'] : '');
    body.replaceText(`{Телефон ${i}}`, relatives[i] ? relatives[i]['Телефон'] : '');
  }

  body.getTables()[0].getCell(0, 0).insertImage(0, photo);
  resizeImage(body.getImages()[0], 300);
}

function onOpen() {
  const exportMenu = SpreadsheetApp.getUi().createMenu("Експорт");
  exportMenu.addItem('Анкета', 'exportQuestionary');
  exportMenu.addItem('Довідка (Ф-1)', 'exportF1');
  exportMenu.addItem('Розшукова картка', 'exportWantedCard');
  
  const menu = SpreadsheetApp.getUi().createMenu("Алла 💖");
  menu.addItem("Оновити таблицю", "reloadTable");
  menu.addSeparator();
  menu.addSubMenu(exportMenu);
  
  menu.addToUi();
}
