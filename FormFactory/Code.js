function populateForm() {
  const formId = SpreadsheetApp.getUi().prompt("Form ID").getResponseText();
  const form = FormApp.openById(formId);

  while (form.getItems().length > 0) {
    form.deleteItem(form.getItems().pop());
  }

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const formData = sheet.getDataRange().getValues().slice(1);
  const formItems = formData.map(([question, helpText, type, required, options]) => {
    return { question, helpText, type, required: required == 'Yes', options: options.split(';').map(o => o.trim()) };
  });

  formItems.forEach(({ question, helpText, type, required, options }) => {
    let formItem;
    switch (type) {
      case "paragraph":
        formItem = form.addParagraphTextItem();
        break;
      case "checkbox":
        formItem = form.addCheckboxItem().setChoiceValues(options);
        break;
      case "radio":
        formItem = form.addMultipleChoiceItem().setChoiceValues(options);
        break;
      case "dropdown":
        formItem = form.addListItem().setChoiceValues(options);
        break;
      case "date":
        formItem = form.addDateItem();
        break;
      default:
        formItem = form.addTextItem();
        break;
    }
    
    formItem.setTitle(question).setHelpText(helpText).setRequired(required);
    Utilities.sleep(500);
  });
} 

function onOpen() {
  const menu = SpreadsheetApp.getUi().createMenu('Form 🏭️');
  menu.addItem('Populate form', 'populateForm');
  menu.addToUi();
}
