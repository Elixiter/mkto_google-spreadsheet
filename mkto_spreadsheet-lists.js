// ==================================================
// BEGIN frontmatter
//
//
// Author information
// --------------------
// Kyle Halstvedt
// Elixiter, Inc.
// kyle@elixiter.com
// --------------------
//
// Setup (*IMPORTANT*)
// --------------------
// 1. From within a new Google Spreadsheet file,
//    open Tools -> Script editor...
// 2. File -> New -> Project, and select
//    Create a script for: Spreadsheet
// 3. Paste ALL of the code in this file into
//    the editor
// 4. CONFIGURE the script by replacing the
//    "REPLACE_ME" strings in the configuration
//    section at the top with URL's/keys that
//    are specific to your Marketo instance.
//    If you have trouble with this, see
//    Marketo's documentation at
//    http://developers.marketo.com/documentation/rest/
// --------------------
//
// Usage
// --------------------
// 1. From a Spreadsheet file with the script
//    included, click the new "Marketo Import"
//    menu item, and click "Initialize sidebar..."
// 2. A right-hand sidebar should appear with the names
//    and IDs of the lists fetched from Marketo
// 3. Each list has an "Insert" button. Click it to insert
//    the members of the list.
// --------------------
//
// Todo / Caveats
// --------------------
// 1. Paginated API; currently only retrieves
//    the first 100 results from a list.
// --------------------
//
//
// END: frontmatter
// ==================================================

// Configuration
// -------------
// YOU MUST REPLACE THESE VALUES, keeping the quotation marks intact.
var restEndpoint = 'REPLACE_ME'; // Marketo REST API endpoint
var identityUrl = 'REPLACE_ME'; // Marketo REST API identity service
var consumerKey = 'REPLACE_ME'; // Marketo REST API client ID
var consumerSecret = 'REPLACE_ME'; // Marketo REST API client secret
// END: configuration

// when the document is opened, create the top menu
function onOpen() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();
  var menu = ui.createMenu('Marketo Import');
  menu.addItem('Initialize sidebar...', 'fetchLists');
  menu.addToUi();
}
// END: onOpen()

// fetch list names from the REST API
function fetchLists() {
  var listsArray = [];
  var bearerToken = JSON.parse(UrlFetchApp.fetch(identityUrl + 'oauth/token?grant_type=client_credentials&client_id=' + consumerKey + '&client_secret=' + consumerSecret).getContentText()).access_token;
  var requestUrl = restEndpoint + 'v1/lists.json' + '?access_token=' + bearerToken;
  var response = UrlFetchApp.fetch(requestUrl);
  var parsedResponse = JSON.parse(response.getContentText());
  for (var n in parsedResponse.result) {
    listsArray.push({
      id : parsedResponse.result[n].id,
      name : parsedResponse.result[n].name
    });
  }
  // create a sidebar with the array of list ID's/names
  createSidebar(listsArray);
}
// END: fetchLists()

// create sidebar to display list names and ID's
// also includes 'insert' button to copy the list to the current spreadsheet
function createSidebar(listsArray) {
  var app = UiApp.createApplication().setTitle('Marketo Lists');
  var scroll = app.createScrollPanel().setHeight('100%').setWidth('100%');
  var vertical = app.createVerticalPanel();
  // for each item in lists array, create sidebar element
  // each sidebar element is a HorizontalPanel with a button and two labels
  // entire sidebar is a ScrollPanel containing a VerticalPanel containing each HorizontalPanel
  for (var l in listsArray) {
    var horizontal = app.createHorizontalPanel();
    var button = app.createButton('Insert').setId(listsArray[l].id); // set button ID to MKTO list ID
    var idLabel = app.createLabel(listsArray[l].id);
    var nameLabel = app.createLabel(listsArray[l].name);
    var handler = app.createServerHandler('buttonHandler'); // specify handler function
    button.addClickHandler(handler); // attach handler to button click event
    horizontal.setVerticalAlignment(UiApp.VerticalAlignment.MIDDLE).setSpacing(10); // set panel format options
    horizontal.add(button).add(idLabel).add(nameLabel); // add button and labels all at once
//    horizontal.add(idLabel);
//    horizontal.add(nameLabel);
    vertical.add(horizontal);
  }
  // END: for (var l in listsArray)
  scroll.add(vertical);
  app.add(scroll);
  SpreadsheetApp.getUi().showSidebar(app); // finally, display the sidebar
}
// END: createSidebar()

// handle click events from sidebar buttons
function buttonHandler(eventInfo) {
  var list = fetchList(eventInfo.parameter.source); // source is ID for button element
  insertList(list);
}
// END: buttonHandler()

// retrieve list by ID from MKTO REST API
// returns an array of objects with keys:
//   id, email, firstName, lastName
function fetchList(id) {
  var listArray = [];
  var bearerToken = JSON.parse(UrlFetchApp.fetch(identityUrl + 'oauth/token?grant_type=client_credentials&client_id=' + consumerKey + '&client_secret=' + consumerSecret).getContentText()).access_token;
  var requestUrl = restEndpoint + 'v1/list/' + id + '/leads.json' + '?access_token=' + bearerToken;
  var response = UrlFetchApp.fetch(requestUrl);
  var parsedResponse = JSON.parse(response.getContentText());
  for (var n in parsedResponse.result) {
    listArray.push({
      id: parsedResponse.result[n].id,
      email: parsedResponse.result[n].email,
      firstName: parsedResponse.result[n].firstName,
      lastName: parsedResponse.result[n].lastName
    });
  }
  return listArray;
}
// END: fetchList()

// insert the contents of the "list" array
//   into the currently-active sheet
function insertList(list) {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getActiveSheet();
  sheet.appendRow(["ID", "Email", "First Name", "Last Name"]);
  sheet.setFrozenRows(1); // freeze the first row into a header
  for (var entry in list) {
    var row = list[entry];
    sheet.appendRow([row.id, row.email, row.firstName, row.lastName]);
  }
}
// END: insertList()
