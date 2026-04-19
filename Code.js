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
  if (!scriptProperties.getProperty('PROGRESS')) {
    console.log('Progress counter init');
    setPropertiesStore('PROGRESS', { count: 0 });
  }

  for (;;) {
    const currentCount = getStoreProperty('PROGRESS', 'count');
    console.log(`Current count ${currentCount}`);

    Utilities.sleep(1000 * 60);

    if (currentCount == 7) {
      console.log('We are finished')
      break;
    }
    setStoreProperty('PROGRESS', { count: currentCount + 1 });
  }
}
