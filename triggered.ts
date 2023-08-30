function triggeredEnrich() {
    const document = SpreadsheetApp.openById(getChristinaSheet());
    const contacts = document.getSheetByName('Contacts') ?? document.getSheetByName('Contacts -> Enriching! 📡');
    if (!contacts) throw new Error("⛔ No Contacts sheet found!");
    contacts?.setTabColor('red');
    contacts.setName('Contacts -> Enriching! 📡');
    const propServ = PropertiesService.getScriptProperties();
    const previous = propServ.getProperty('triggeredEnrich-previous');
    const rangeObj = {top: previous ? Number(previous) : 0, end: 0};
    document.toast(JSON.stringify(rangeObj));
    const emailCol = contacts.getRange('M2:M' + contacts.getLastRow()).getValues().flat();
    while ((rangeObj.end - rangeObj.top) < 2) {
        rangeObj.top = emailCol.findIndex((element, row) => row > rangeObj.top && !element.includes('@'));
        rangeObj.end = emailCol.findIndex((element, row) => row > rangeObj.top && element.includes('@')) - 1;
    }
    propServ.setProperty('triggeredEnrich-previous', (rangeObj.top).toString());
    rangeObj.top += 2, rangeObj.end += 2;
    console.log('Found rows to enrich:', rangeObj);
    try {
        callEnricher('A' + rangeObj.top + ':A' + rangeObj.end, document, contacts);
    } catch(e) {
        console.error('Error message: ' + e);
    } finally {
        contacts.setTabColor(null);
        contacts.setName('Contacts');
    }
}