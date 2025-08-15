/*********************************************************************************************/
const muted = true;
const repeatCount = 2;
const sheetName = 'Поранені';
const text = 'АЛЛА, ГАРНОГО ДНЯ! ❤️❤️❤️';
/*********************************************************************************************/

function runTextLineInSheetTab(options) {
  const { text, sheet, frameSize = 4, repeatCount = 1 } = options;
  const frames = [];
  const textLine = Array.from({ length: repeatCount }, () => text).join(' ');

  for (let i = 0; i <= textLine.length - frameSize; i += 1) {
    frames.push(textLine.slice(i, frameSize + i));
  }

  frames.forEach(f => sheet.setName(f));
}

function sheetTabNotification() {
  if (muted) return;

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);

  runTextLineInSheetTab({ text, frameSize: sheetName.length, sheet, repeatCount });
  Utilities.sleep(500);
  sheet.setName(sheetName);
}
