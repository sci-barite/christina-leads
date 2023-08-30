function GETLINK(input: any){
  var sheet = SpreadsheetApp.getActiveSheet();
  var range = sheet.getRange(input);
  var value = range.getRichTextValue();
  var url = value?.getLinkUrl();
  return url;
}

function fromLiToWebsite() {
    const Window = SpreadsheetApp.getActiveSpreadsheet();
    const Spread = Window.getActiveSheet();
    const Header = Spread.getRange('1:1').getValues().flat();
    const Linked = Header.indexOf("Company Linkedin URL") + 1;
    if (!Linked) throw new Error('Column "Company Linkedin URL" not found!');
    const ComURL = Header.indexOf("Company Website") + 1;
    if (!ComURL) throw new Error('Column "Company Website" not found!');
    const Source = Spread.getRange(2, Linked, Spread.getLastRow() + 1, 1);
    const Target = Spread.getRange(2, ComURL, Spread.getLastRow() + 1, 1);
    const AllURL = Source.getValues();
    const Failed = [] as number[];
    AllURL.forEach((link, n) => {
      if (!link[0].includes('linkedin.com/company')) return;
      const CompID = link[0].split('/').pop();
      const Search = {fetch: UrlFetchApp.fetch('https://duckduckgo.com/?q=what+is+the+website+of+company+' + CompID, {muteHttpExceptions: true })};
      const DuckGo = Search.fetch.getContentText();
      const Result = DuckGo?.split('<h2 class="result__title">')?.[1];
      const NewURL = Result?.split('href="//duckduckgo.com/l/?uddg=')?.[1]?.split('%2F&amp;rut=')?.shift()?.replace('%3A%2F%2F', '://')?.split('%')?.[0];
      link[0] = (NewURL?.includes('linkedin') || NewURL?.includes('facebook')) ? '❓CANNOT GUESS' : NewURL;
      if (link[0].startsWith('❓')) Failed.push(n);
      if (n % 5 === 0) Window.toast('Guessed ' + (n + 1 - Failed.length) + ' out of ' + AllURL.length ,'❓🌍 GUESSING WEBSITES...');
    })
    Target.setValues(AllURL);
  }

function getCompaniesHiring() {
    const props = PropertiesService.getScriptProperties();
    const dbURL = props.getProperty('FWDB');
    const prevs = props.getProperty('PreviouslyFetched');
    const array = prevs ? JSON.parse(prevs) : [] as string[];
    SpreadsheetApp.getActiveSpreadsheet().toast(
        'Contacting FWDB Leads...', 
        '➕💼 GET COMPANIES HIRING'
    );
    const comps = JSON.parse(UrlFetchApp.fetch(dbURL + '?request=getCompaniesHiring' + 
        (prevs ? '&prevs=' + array.join('__') : '')).getContentText()) as string[];
    if (comps[0].startsWith('ERR')) {
        SpreadsheetApp.getActiveSpreadsheet().toast(
            comps[0], 
            '⛔💼 GET COMPANIES HIRING'
        );
        return;
    } 
    props.setProperty('PreviouslyFetched', JSON.stringify([...comps, ...array]));
    SpreadsheetApp.getActiveSpreadsheet().toast(
        'New sheet has ' + comps.length + ' companies hiring, you can find it marked in green, below.', 
        '✅💼 GET COMPANIES HIRING'
    );
}

function editPositions() {
    const props = PropertiesService.getScriptProperties();
    const inter = SpreadsheetApp.getUi();
    const poses = JSON.parse(props.getProperty('LeadsSearchTerms') ?? '[]') as string[];
    const reply = inter.alert(
        '📑🗃️ POSITIONS TO SEARCH FOR\n\nCurrently searched positions are:\n\n' +
        '"' + poses.join(', ') + '"\n\n' +
        'Click Yes to confirm, No to add/delete', inter.ButtonSet.YES_NO
    );
    if (reply === inter.Button.YES) return;
    const edit = inter.alert(
        '📑🗃️ POSITIONS TO SEARCH FOR\n\nCurrently searched positions are:\n\n' +
        '"' + poses.join(', ') + '"\n\n' +
        'Click Yes to add positions, No to delete', inter.ButtonSet.YES_NO_CANCEL
    );
    if (edit === inter.Button.CANCEL) return;
    const yOrN = edit === inter.Button.YES ? true : false;
    const answer = inter.prompt(
        '📑🗃️ POSITIONS TO SEARCH FOR\n\nCurrently searched positions are:\n\n' +
        '"' + poses.join(', ') + '"\n\n' +
        'Enter terms to ' + (yOrN ? 'add' : 'delete') + ', separated by commas:\n\n'
    );
    const edits = answer.getResponseText().split(',').map(term => term.trim());
    const newPoses = poses.concat(edits);
    if (!yOrN) edits.forEach(term => {
        const indexToDelete = poses.indexOf(term);
        if (indexToDelete < 0) return;
        poses.splice(indexToDelete, 1);
    });
    props.setProperty('LeadsSearchTerms', JSON.stringify(yOrN ? newPoses : poses));
    inter.alert(
        '📑🗃️ POSITIONS TO SEARCH FOR\n\nNew searched positions are:\n\n' +
        '"' + (yOrN ? newPoses.join(', ') : poses.join(', ')) + '"\n\n' +
        'Please use the function again if you need to make more changes.'
    );
}

function setFindContacts() {
    const props = PropertiesService.getScriptProperties();
    const inter = SpreadsheetApp.getUi();
    const reply = inter.alert(
        '▶️🗃️ START FINDING CONTACTS\n\nCurrently searched positions are:\n\n' +
        '"' + JSON.parse(props.getProperty('LeadsSearchTerms') ?? '[]').join(', ') + '"\n\n' +
        'You can change these with "📑🗃️ Positions to search for..." from the ➕ menu.\n\n' +
        'Start finding new contacts for the current sheet?', inter.ButtonSet.YES_NO
    );
    if (reply === inter.Button.NO) return;
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getName();
    props.setProperty('FindContactsTargetSheet', sheet);
    const trigger = ScriptApp.newTrigger('findContacts')
        .timeBased()
        .everyMinutes(1)
        .create();
    props.setProperty('FindContactsTrigger', trigger.getUniqueId());
    inter.alert(
        '▶️🗃️ START FINDING CONTACTS\n\nSet a trigger to find new contacts every minute through Apollo.\n\n' +
        'Please avoid editing the sheet until the script is finished.\n' +
        'You can stop the job with "Stop finding new contacts..." from the ➕ menu.\n' +
        'You can change the searched positions from "Script Properties -> Edit Script Properties -> LeadsSearchTerms" at:\n\n' +
        'https://script.google.com/u/0/home/projects/1B-5kDYVdu7jEXJ7h-jSGqlNUKoDRh5kYt37zmkpDPp4VQ6cFfq1yBMPM/settings'
    );
}

function delFindContacts() {
    const props = PropertiesService.getScriptProperties();
    const inter = SpreadsheetApp.getUi();
    const sheet = props.getProperty('FindContactsTargetSheet');
    if (!sheet) throw new Error('No target sheet set! Set it with "▶️🗃️ Start finding new contacts..."');
    const reply = inter.alert(
        '⏹️🗃️ STOP FINDING CONTACTS\n\nCurrently searched sheet is:\n\n' +
        '"' + sheet + '"\n' +
        'Stop finding new contacts for it?', inter.ButtonSet.YES_NO
    );
    if (reply === inter.Button.NO) return;
    props.setProperty('FindContactsTargetSheet', '');
    const triggerID = props.getProperty('FindContactsTrigger');
    const trigger = ScriptApp.getProjectTriggers().find(trigger => trigger.getUniqueId() === triggerID);
    if (!trigger) throw new Error('Trigger not found!');
    ScriptApp.deleteTrigger(trigger);
    SpreadsheetApp.getActiveSpreadsheet().toast('"Find Contacts" job for sheet "' + sheet + '" stopped successfully.', '⏹️🗃️ JOB STOPPED')
}

function setEnrichContacts() {
    const props = PropertiesService.getScriptProperties();
    const inter = SpreadsheetApp.getUi();
    const reply = inter.alert(
        '▶️📡 START ENRICHING CONTACTS\n\nCurrently searched positions are:\n\n' +
        JSON.parse(props.getProperty('LeadsSearchTerms') ?? '[]').join('\n') + '\n' +
        'Start enriching contacts for the current sheet?', inter.ButtonSet.YES_NO
    );
    if (reply === inter.Button.NO) return;
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getName();
    props.setProperty('EnrichContactsTargetSheet', sheet);
    const trigger = ScriptApp.newTrigger('triggeredEnrich')
        .timeBased()
        .everyMinutes(1)
        .create();
    props.setProperty('EnrichContactsTrigger', trigger.getUniqueId());
    inter.alert(
        '▶️📡 START ENRICHING CONTACTS\n\nSet a trigger to enrich contacts every minute through Apollo.\n\n' +
        'Please avoid editing the sheet until the script is finished.\n' +
        'You can stop the job with "Stop enriching contacts..." from the ➕ menu.'
    );
}

function delEnrichContacts() {
    const props = PropertiesService.getScriptProperties();
    const inter = SpreadsheetApp.getUi();
    const sheet = props.getProperty('EnrichContactsTargetSheet');
    if (!sheet) throw new Error('No target sheet set! Set it with "▶️📡 Start enriching contacts..."');
    const reply = inter.alert(
        '⏹️📡 STOP ENRICHING CONTACTS\n\nCurrently searched sheet is:\n\n' +
        '"' + sheet + '"\n' +
        'Stop enriching contacts for it?', inter.ButtonSet.YES_NO
    );
    if (reply === inter.Button.NO) return;
    props.setProperty('EnrichContactsTargetSheet', '');
    const triggerID = props.getProperty('EnrichContactsTrigger');
    const trigger = ScriptApp.getProjectTriggers().find(trigger => trigger.getUniqueId() === triggerID);
    if (!trigger) throw new Error('Trigger not found!');
    ScriptApp.deleteTrigger(trigger);
    SpreadsheetApp.getActiveSpreadsheet().toast('"Enrich Contacts" job for sheet "' + sheet + '" stopped successfully.', '⏹️📡 JOB STOPPED')
}

function findContacts() {
    const props = PropertiesService.getScriptProperties();
    const dbURL = props.getProperty('FWDB');
    const terms = JSON.parse(props.getProperty('LeadsSearchTerms') ?? '[]');
    const sName = props.getProperty('FindContactsTargetSheet');
    if (!sName) throw new Error('No taget sheet set! Set it with "▶️🗃️ Start finding new contacts..." from the ➕ menu.');
    const sheet = SpreadsheetApp.openById(getChristinaSheet()).getSheetByName(sName);
    if (!sheet) throw new Error('Target sheet "' + sName + '" not found! Set it correctly with "Start finding new contacts..." from the ➕ menu.');
    const found = JSON.parse(props.getProperty('SearchedOnApollo-' + sName.split(' ').join('')) ?? '[]');
    const comps = sheet.getRange('C2:C' + sheet.getLastRow()).getValues().flat().map((url, row) => [row, url]);
    //const leads = sheet.getRange('L2:L' + sheet.getLastRow()).getValues().flat().map((url, row) => [row, url]);
    for (const [row, comp] of comps) {
        if (found.includes(comp)) continue;
        const logging = ['From row ' + (row + 1), '🗃️ Trying "' + comp + '"'];
        console.log(logging[1] + ' ' + logging[0]);
        SpreadsheetApp.getActiveSpreadsheet().toast(logging[0], logging[1]);
        const companyRows = comps.filter(name => name[1] === comp);
        //const excludeList = leads.filter(people => companyRows.find(company => company[0] === people[0])).map(person => person[1]);
        const reply = UrlFetchApp.fetch(dbURL + '?request=apolloPeopleFind&domain=' + comp
            + '&targetRow=' + companyRows.at(-1)?.[0]
            + '&titles=' + terms.join('__')
            + '&sheet=' + sName);
        //    + '&excludeLinkedIn=' + excludeList.join('__'));
        reply.getContentText() !== 'GOOD' ? console.error(reply.getContentText()) : console.log('GOOD');
        found.push(comp);
        const results = [
            comp + ' at found[' + found.indexOf(comp) + '] row ' + (row + 1), 
            reply.getContentText().length < 12 ? 'FWDB Reply: ' + reply.getContentText() : '⛔ API issue: try in 1h'
        ];
        SpreadsheetApp.getActiveSpreadsheet().toast(results[0], results[1]);
        console.warn(results[1], results[0]);
        break;
    }
    props.setProperty('SearchedOnApollo-' + sName.split(' ').join(''), JSON.stringify(found));
}

function enrichContacts() {
    const spred = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = spred.getActiveSheet();
    if (!sheet) throw new Error('No "Contacts" sheet!');
    const A1not = sheet.getActiveRange()?.getA1Notation();
    console.log('Active range: ', A1not);
    if (!A1not) throw new Error('🐞 ERROR: No active selection found. Please select some cells/rows.');
    callEnricher(A1not, spred, sheet);
}
function callEnricher(A1range: string, spread: GoogleAppsScript.Spreadsheet.Spreadsheet, sheet: GoogleAppsScript.Spreadsheet.Sheet) {
    const A1row = A1range.split(':').map(A1 => parseInt(A1.substring(1)));
    if (A1row.some(row => isNaN(row))) throw new Error('🐞 ERROR: Parsed rows "' + A1row.toString() + '" are invalid.');
    const conts = sheet.getRange('I' + A1row[0] + ':O' + (A1row[1] || A1row[0])).getValues();
    const comps = sheet.getRange('C' + A1row[0] + ':C' + (A1row[1] || A1row[0])).getValues().flat();
    const first = comps[0];
    const diffs = comps.findIndex(company => company !== first);
    if (diffs !== -1 || conts.length > 10) {
        console.warn('Can only enrich up to 10 contacts, all from a single company. Cutting the list accordingly.');
        spread.toast('Can only enrich up to 10 contacts, all from a single company. Cutting the list accordingly.', '⚠️ CUTTING LIST');
    }
    const outgo = conts.slice(0, diffs > 0 ? diffs : 10).slice(0, 10);
    A1row[1] = A1row[1] - (conts.length - outgo.length);
    const names = outgo.map(cont => '- ' + cont[0] + ' ' + cont[1]).join('\n');
    spread.toast(names, '▶️ Enriching till ' + (A1row[1] || A1row[0]));
    console.log('▶️ Enriching till ' + (A1row[1] || A1row[0]) + '\n' + names);

    const reply = UrlFetchApp.fetch(PropertiesService.getScriptProperties().getProperty('FWDB') + '?request=apolloPeopleEnrich&domain=' + first
            + '&organization=' + sheet.getRange('B' + A1row[0]).getValue()
            + '&firstNames=' + outgo.map(cont => cont[0]).join('__')
            + '&lastNames=' + outgo.map(cont => cont[1]).join('__')
            + '&apolloIDs=' + outgo.map(cont => cont[5]).join('__')
            + '&topOfRange=' + A1row[0]
            + '&endOfRange=' + (A1row[1] || A1row[0]
            + '&sheet=' + sheet.getName())
    );
    spread.toast(reply.getContentText().split('\n')[0], '✅ Enrichment done!');
}

function getChristinaSheet() {
    return '1NlLgnvvYqpS31i2qJXljMcfvjEJ0V2yWmbzRAeleoSg';
}

function consolidateContacts() {
    const SS = SpreadsheetApp.getActiveSpreadsheet();
    const AS = SS.getActiveSheet();
    const LR = AS.getLastRow();

    const LinkedIn = AS.getRange('L2:L' + LR).getValues().flat();
    const Contacts = AS.getRange('L2:N' + LR).getValues();
    
    LinkedIn.forEach((contact, row) => {
        const enriched = LinkedIn.indexOf(contact, (row + 1));
        if (enriched === -1) return; 
        [Contacts[row][1], Contacts[row][2]] = [Contacts[enriched][1], Contacts[enriched][2]];
    });

    AS.getRange('L2:N' + LR).setValues(Contacts);
}

function eliminateBlues() {
    const SS = SpreadsheetApp.getActiveSpreadsheet();
    const AS = SS.getSheetByName("Contacts");
    if (!AS) throw new Error('Sheet "Contacts" not found!');
    const RN = AS.getRange("B:B"), CompaniesBGs = RN.getBackgrounds().flat();

    const Values : any[][] = [], RowNs : any[] = [];
  
    CompaniesBGs.forEach((color, row) => {
        if (color !== "#e0ffff") return;
        const rowN = row + 1;
        const Row = AS.getRange((rowN)+':'+(rowN));
        RowNs.push(rowN);
        Values.push(Row.getValues().flat());
        if (row % 20 === 0) SS.toast("Row " + (row), "🧮 Processing...")
    });

    RowNs.sort((a,b) => (b - a)); // Inverts the order to delete bottom-to-top, which doesn't require row number corrections.

    const lock = LockService.getScriptLock();
    try {
        // Acquire the lock (wait up to 15 seconds to acquire the lock)
        lock.waitLock(15000);
        SS.toast("Deleting" + RowNs.length + " rows. Please don't edit until the job is done!", "⚠️ DELETING")
        RowNs.forEach((row) => AS.deleteRow(row));  // Deleting first, otherwise conditional formatting inside the trash will slow this down.
        SS.toast("Deleted " + RowNs.length, "✅ DONE")
        throwToTrash(SS, AS, '🗑', RowNs.length, Values);
        SS.toast("Backed up " + RowNs.length, "🗃️ BACKUP DONE")
    } catch (e) {
        console.error('Error: ', e);
        SS.toast('Error message: ' + e, '⛔ ERROR')
    } finally {
        // Release the lock so other scripts can run
        lock.releaseLock();
    }

}

function throwToTrash(SS : GoogleAppsScript.Spreadsheet.Spreadsheet, DB : GoogleAppsScript.Spreadsheet.Sheet, 
    Target : string, Rows : number, Data : any [][], RichData? : any [][]) {
let Trash = SS.getSheetByName(Target);
if(!Trash) {
Trash = SS.insertSheet().setName(Target);
const Header = DB!.getRange("1:1"), CopyHeader = Trash.getRange("1:1");
Header.copyTo(CopyHeader);
Header.copyTo(CopyHeader, SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
Header.copyTo(CopyHeader, SpreadsheetApp.CopyPasteType.PASTE_COLUMN_WIDTHS, false);
}

const StartRow = Trash.getLastRow() + 1;
const TrashRange = Trash.getRange(StartRow, 1, Rows, Data[0].length), TemplateRange = DB.getRange(2,1,2,21);
Trash.insertRows(StartRow, Rows); // To make sure we don't run out of row, we always insert the same amount we copy.

TrashRange.setValues(Data);

Trash.setConditionalFormatRules([]); // First remove all rules, or we get doubles, triples, etc.
TemplateRange.copyTo(Trash.getDataRange().offset(1, 0), SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false); // Easier.

const TrashRules = Trash.getConditionalFormatRules();
for (let i = TrashRules.length - 1; i >= 0; i--) {  // Reverse loop, suggested by ChatGPT to avoid skipping indexes while splicing.
const bool = TrashRules[i].getBooleanCondition();
if (bool)
if (bool.getCriteriaType() === SpreadsheetApp.BooleanCriteria.CUSTOM_FORMULA) TrashRules.splice(i, 1);
}
Trash.setConditionalFormatRules(TrashRules); // Now apply the set without the formulas.
}