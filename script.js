
function monitorFormResponses() {
  var form = FormApp.getActiveForm();
  var formResponses = form.getResponses();
  
  var stagCount = 0;
  var coupleCount = 0;
  var stagLimit = 300;
  var coupleLimit = 150;
  
  // Find the multiple choice item
  var items = form.getItems();
  var multipleChoiceItem;
  for (var i = 0; i < items.length; i++) {
    if (items[i].getType() === FormApp.ItemType.MULTIPLE_CHOICE) {
      multipleChoiceItem = items[i].asMultipleChoiceItem();
      if (multipleChoiceItem.getChoices().map(choice => choice.getValue()).includes('Stag') &&
          multipleChoiceItem.getChoices().map(choice => choice.getValue()).includes('Couple')) {
        break;
      }
    }
  }
  
  if (!multipleChoiceItem) {
    Logger.log('Multiple choice item with Stag and Couple options not found');
    return;
  }
  
  // Count responses
  for (var i = 0; i < formResponses.length; i++) {
    var response = formResponses[i].getResponseForItem(multipleChoiceItem);
    if (response) {
      if (response.getResponse() === 'Stag') {
        stagCount++;
      } else if (response.getResponse() === 'Couple') {
        coupleCount++;
      }
    }
  }
  
  // Find the sections to close
  var stagSection, coupleSection;
  var sections = form.getItems(FormApp.ItemType.SECTION_HEADER);
  for (var i = 0; i < sections.length; i++) {
    var title = sections[i].getTitle().toLowerCase();
    if (title.includes('stag')) {
      stagSection = sections[i];
    } else if (title.includes('couple')) {
      coupleSection = sections[i];
    }
  }
  
  // Close sections if limits are reached
  if (stagCount >= stagLimit && stagSection) {
    stagSection.asPageBreakItem().setGoToPage(FormApp.PageNavigationType.SUBMIT);
    Logger.log('Stag section closed');
  }
  
  if (coupleCount >= coupleLimit && coupleSection) {
    coupleSection.asPageBreakItem().setGoToPage(FormApp.PageNavigationType.SUBMIT);
    Logger.log('Couple section closed');
  }
}
