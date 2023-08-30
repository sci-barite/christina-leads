function onOpen() {
    const UI = SpreadsheetApp.getUi();
    UI.createMenu('➕')
        .addItem('▶️🗃️  Start finding new contacts...', 'setFindContacts')
        .addItem('⏹️🗃️  Stop finding new contacts....', 'delFindContacts')
        .addItem('📑🗃️  Positions to search for...', 'editPositions')
        .addSeparator()
        .addItem('▶️📡  Start enriching contacts...', 'setEnrichContacts')
        .addItem('⏹️📡  Stop enriching contacts...', 'delEnrichContacts')
        .addItem('✅📡  Enrich selected', 'enrichContacts')
        .addSeparator()
        .addItem('➕💼  Get companies hiring', 'getCompaniesHiring')
        .addSeparator()
        .addItem('❓🌍  Guess company websites', 'fromLiToWebsite')
        .addToUi();
}