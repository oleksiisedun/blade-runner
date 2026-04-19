// =============================================================================
// This script is maintained in a Git repository and is normally edited in an
// external editor, co-authored with Claude Code.
// Repository: https://github.com/oleksiisedun/blade-runner
//
// ⚠ WARNING: Any changes made directly in the Apps Script web editor may be
// overwritten the next time the code is pushed from the repository.
// =============================================================================

const PROGRESS_KEY = PROGRESS_KEY;

function showWebView() {
  const template = HtmlService.createTemplateFromFile('index');

  const html = template.evaluate().setWidth(500).setHeight(500);

  SpreadsheetApp.getUi().showModalDialog(html, 'BladeRunner');
}

function onOpen() {
  const menu = SpreadsheetApp.getUi().createMenu("Oleksii 🛸");
  menu.addItem('Відкрити WebView', 'showWebView');
  menu.addToUi();
}

function heavyTask() {
  if (!scriptProperties.getProperty(PROGRESS_KEY)) {
    setStore(PROGRESS_KEY, { count: 0 });
  }

  for (;;) {
    const progress = getStore(PROGRESS_KEY);

    Utilities.sleep(1000 * 60);

    if (!scriptProperties.getProperty(PROGRESS_KEY)) {
      return;
    }

    if (progress.count === 7) {
      scriptProperties.deleteProperty(PROGRESS_KEY);
      break;
    }

    setStore(PROGRESS_KEY, { ...progress, count: progress.count + 1 });
  }
}

function cancelTask() {
  scriptProperties.deleteProperty(PROGRESS_KEY);
}
