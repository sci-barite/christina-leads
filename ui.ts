function onOpen() {
    const UI = SpreadsheetApp.getUi();
    UI.createMenu('➕')
        .addItem('▶️🗃️  Start finding new contacts...', 'setFindContacts')
        .addItem('⏹️🗃️  Stop finding new contacts....', 'delFindContacts')
        .addItem('📑🗃️  Positions to search for...', 'editPositions')
        .addSeparator()
        .addItem('▶️📡  Start enriching contacts...', 'setEnrichContacts')
        .addItem('⏹️📡  Stop enriching contacts...', 'delEnrichContacts')
        .addToUi();
}