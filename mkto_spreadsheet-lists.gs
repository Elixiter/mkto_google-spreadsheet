// ==================================================
// BEGIN frontmatter

// Author information
// ------------------
// Kyle Halstvedt  
// [Elixiter, Inc.](http://www.elixiter.com)  
// kyle@elixiter.com

// Setup (*IMPORTANT*)
// -------------------
// 1. From within a new Google Spreadsheet file,
//    open Tools -> Script editor...
// 2. Create a new "Blank Project" from the wizard
//    or from File -> New.
// 3. Paste the code from
//    [mkto_spreadsheet-lists.gs]
//    (https://raw.githubusercontent.com/Elixiter/mkto_google-spreadsheet/master/mkto_spreadsheet-lists.gs)
//    into the editor. Make sure you paste over any sample code in the editor.
// 4. Configure the script by replacing the
//    "REPLACE_ME" strings in the configuration
//    section at the top with URL's/keys that
//    are specific to your Marketo instance.
//    If you have trouble with this, see
//    [Marketo's documentation](http://developers.marketo.com/documentation/rest/),
//    specifically the sections on [authentication]
//    (http://developers.marketo.com/documentation/rest/authentication/)
//    and [the custom service]
//    (http://developers.marketo.com/documentation/rest/custom-service/).
// 5. Save the script with the floppy icon or File -> Save.

// Usage
// -----
// 1. Refresh the Spreadsheets window after editing script.
// 2. From the open Spreadsheet with the script
//    included, click the custom "Marketo Import"
//    menu item, and click "Initialize sidebar...". If there
//    is no menu item titled "Marketo Import", refresh the page, and
//    please make sure you have followed all of the steps in Setup.
//    In particular, ensure that the script is present in the script editor
//    for the current document, and that you have filled in your API credentials.
// 3. The first time you run it, you must give authorization to
//    make external http requests.
// 4. A right-hand sidebar should appear with the names
//    and IDs of the lists fetched from Marketo.
// 5. Each list has an "Insert" button. Click it to insert
//    all members of the list into a new spreadsheet.

// Caveats / Todo
// --------------
// 1. Any user with enough priveleges to run the script
//    is able to *read* the script, which contains
//    your REST API credentials (ID and secret key) in-the-clear.  
//    **DO NOT POST YOUR API CREDENTIALS PUBLICALLY!**
// 2. The script does not attempt to "update" or "synchronize" the lists,
//    it currently only creates new sheets. Updating lists is a planned feature.
// 3. I am using the atomic Sheet.appendRow() method to add each row, which is safe
//    but slow. Planned migration to Range.setValues() for performance.
// 4. If you exceed the API quota of 100 requests per 20 seconds, you will receive
//    an error message. This script does not currently support fetching lists
//    of greater than 10k leads, as it has no regulator to prevent reaching
//    the API limit and no way to resume fetching a list in the middle.
// 5. The UX is quite poor: there are no status or loading indicators. Improving
//    this is on the long-term roadmap.


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
  if (!isConfigured()) {
    throw new Error('You have not entered your REST API credentials.' + '\n' +
		    'Please follow the configuration instructions.');
  }
  var ui = SpreadsheetApp.getUi();
  var menu = ui.createMenu('Marketo Import');
  menu.addItem('Initialize sidebar...', 'initializeSidebar');
  menu.addToUi();
}
// END: onOpen()

function isConfigured() {
  if (restEndpoint == 'REPLACE_ME') {
    return false;
  }
  else if ( identityUrl == 'REPLACE_ME') {
    return false;
  }
  else if ( consumerKey == 'REPLACE_ME') {
    return false;
  }
  else if ( consumerSecret == 'REPLACE_ME') {
    return false;
  }
  else {
    return true;
  }
}

function initializeSidebar() {
  var scriptProperties = PropertiesService.getScriptProperties();
  // create hidden fields to cache authentication token
  var tokenField = scriptProperties.setProperty('tokenValue', '');
  var timeStampField = scriptProperties.setProperty('tokenTimeStamp', '');
  var expiryField = scriptProperties.setProperty('tokenExpiry', '9999');
  var lists = fetchListsR();
  createSidebar(lists);
}
// END: initializeSidebar()

function fetchListsR(args) {
  var args = args || {};
  var listsArray = args.listsArray || [];
  var rest = new MktoClient();
  var bearerToken = rest.getToken();
  var requestUrl = restEndpoint + 'v1/lists.json' + '?access_token=' + bearerToken;
  // if passed a next page token
  if (args.nextPage) {
    requestUrl += '&nextPageToken=' + args.nextPage;
  }
  var response = UrlFetchApp.fetch(requestUrl);
  var content = response.getContentText();
  var parsed = JSON.parse(content);
  if (parsed.success != true) {
    throw new Error('The API request failed.' + '\n' +
		    parsed.errors[0].code + '\n' +
		    parsed.errors[0].message);
  }
  for (var n = 0; n < parsed.result.length; n++) {
    listsArray.push({
      id: parsed.result[n].id,
      name: parsed.result[n].name
    });
  }
  // if there are more lists...
  if (parsed.nextPageToken) {
    // recurse
    fetchListsR({ listsArray: listsArray, nextPage: parsed.nextPageToken });
  }
  // done recursing
  else {
    return listsArray;
  }
}
// END: fetchListsR()

// create sidebar to display list names and ID's
// also includes 'insert' button to copy the list to the current spreadsheet
function createSidebar(listsArray) {
  var app = UiApp
    .createApplication()
    .setTitle('Marketo Lists (first 100 only)');
  var scroll = app
    .createScrollPanel()
    .setHeight('100%')
    .setWidth('100%');
  var vertical = app.createVerticalPanel();
  // for each item in lists array, create sidebar element
  // each sidebar element is a HorizontalPanel with a button and two labels
  // entire sidebar is a ScrollPanel containing a VerticalPanel with HorizontalPanels
  for (var l = 0; l < listsArray.length; l++) {
    var horizontal = app.createHorizontalPanel();
    var button = app
      .createButton('Insert')
      .setId(listsArray[l].id); // set button ID to MKTO list ID
    var idLabel = app.createLabel(listsArray[l].id);
    var nameLabel = app.createLabel(listsArray[l].name);
    // create hidden callback element to pass to button click handler
    var nameHidden = app.createHidden('nameHidden', listsArray[l].name);
    app.add(nameHidden);
    var handler = app
      .createServerHandler('buttonHandler')
      .addCallbackElement(nameHidden);
    button.addClickHandler(handler); // attach handler to button click event
    horizontal
      .setVerticalAlignment(UiApp.VerticalAlignment.MIDDLE)
      .setSpacing(10); // set panel format options
    horizontal
      .add(button)
      .add(idLabel)
      .add(nameLabel); // add button and labels all at once
    vertical.add(horizontal);
  }
  // END: for (l in listsArray)
  scroll.add(vertical);
  app.add(scroll);
  SpreadsheetApp
    .getUi()
    .showSidebar(app); // finally, display the sidebar
}
// END: createSidebar()

function insertHeader() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getActiveSheet(); // currently adding to active sheet
  sheet.appendRow(['ID', 'Email', 'First Name', 'Last Name']);
  sheet.setFrozenRows(1); // freeze the first row into a header
}
// END: insertHeader()

function resizeColumns() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getActiveSheet();
  sheet
    .autoResizeColumn(1)
    .autoResizeColumn(2)
    .autoResizeColumn(3)
    .autoResizeColumn(4);
}
// END: resizeColumns()

// handle click events from sidebar buttons
// TODO: add list name as callback element
function buttonHandler(eventInfo) {
  var id = eventInfo.parameter.source;
  var name = eventInfo.parameter.nameHidden;
  var label = id + ' | ' + name;
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getActiveSheet();
  var range = sheet.getDataRange();
  var sheets = spreadsheet.getSheets();
  // sheet exists
  for (var i = 0; i< sheets.length; i++) {
    if (label == sheets[i].getName()) {
      throw new Error('Sheet with name: "' + label + '" already exists.');
    }
  }
  // if there is data in the current sheet, create a new one
  if (!range.isBlank()) {
    // name is list ID
    spreadsheet.setActiveSheet(spreadsheet.insertSheet());
    // update reference to current sheet
    sheet = spreadsheet.getActiveSheet();
  }
  // sheet name is list's: <id> | <name>
  sheet.setName(label);
  insertHeader(); // insert top header to spreadsheet
  // call recusive fetch
  fetchAndInsertListR({ id: id });
  // resizeColumns(); // not available in new Google Spreadsheet yet
}
// END: buttonHandler()

// recursive list-grabbing
// single argument assumed to be an object
//   to simplify passing multiple named arguments
// if no nextPage is given, assume it's the
//   first page and keep on fetching
// args: {id, nextPage}
function fetchAndInsertListR(args) {
  var args = args || {};
  var rest = new MktoClient();
  var bearerToken = rest.getToken();

  // called with no id
  if (!args.id) {
    throw new Error('The list ID is undefined.');
  }

  // there is an id, make a request
  else {
    var listArray = [];
    var requestUrl = restEndpoint +
      'v1/list/' +
      args.id +
      '/leads.json' +
      '?access_token=' +
      bearerToken;
    // if passed next page token, append it
    if (args.nextPage) {
      requestUrl += '&nextPageToken=' + args.nextPage;
    }
    var response = UrlFetchApp.fetch(requestUrl);
    var parsedResponse = JSON.parse(response.getContentText());
    if (parsedResponse.success != true) {
      throw new Error('The API request failed.' + '\n' +
		      parsedResponse.errors[0].code + '\n' +
		      parsedResponse.errors[0].message);
    }
    // if this result was not empty...
    else if (parsedResponse.result.length > 0) {
      // construct the array from the response
      for (var n = 0; n < parsedResponse.result.length; n++) {
	listArray.push({
	  id: parsedResponse.result[n].id,
	  email: parsedResponse.result[n].email,
	  firstName: parsedResponse.result[n].firstName,
	  lastName: parsedResponse.result[n].lastName
	});
      }
      // insert array directly into the spreadsheet
      insertList(listArray);
      // ...and recurse
      if (parsedResponse.nextPageToken) {
	fetchAndInsertListR({ id: args.id, nextPage: parsedResponse.nextPageToken, listArray: listArray });
      }
    }
  }
}
// END: fetchAndInsertListR()

// insert the contents of the list array
//   into the currently-active sheet
function insertList(list) {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getActiveSheet();
  for (var n = 0; n < list.length; n++) {
    var row = list[n];
    sheet.appendRow([row.id, row.email, row.firstName, row.lastName]);
  }
}
// END: insertList()


// manage connecting with Marketo REST API
var MktoClient = function() {

  // fetch token from REST authentication endpoint
  var authenticate = function() {
    var request = identityUrl +
      'oauth/token?grant_type=client_credentials' +
      '&client_id=' + consumerKey +
      '&client_secret=' + consumerSecret;
    var response = UrlFetchApp.fetch(request);
    var content = response.getContentText();
    var parsed = JSON.parse(content);
    if (parsed.error) {
      throw new Error('The authentication request failed.' + '\n' +
		      parsed.error + '\n' +
		      parsed.error_description);
    }
    else {
      var tokenValue = parsed.access_token;
      var tokenExpiry = parsed.expires_in; // seconds
      var tokenTimeStamp = new Date().getTime() / 1000; // now, converted to seconds
      // construct token object
      var token = { value: tokenValue, timeStamp: tokenTimeStamp, expiry: tokenExpiry }
      cacheToken(token);
      return token;
    }
  }
  // END: authenticate()

  // cache token in hidden fields
  // expects a token: {value, timeStamp, expiry}
  var cacheToken = function(tokenObject) {
    // ensure we are writing strings
    var value = tokenObject.value || '';
    var timeStamp = tokenObject.timeStamp || '';
    var expiry = tokenObject.expiry || '';
    var scriptProperties = PropertiesService.getScriptProperties();
    // set the hidden fields
    //var app = UiApp.getActiveApplication();
    scriptProperties.setProperty('tokenValue', value);
    scriptProperties.setProperty('tokenTimeStamp', timeStamp);
    scriptProperties.setProperty('tokenExpiry', expiry);
  }
  // END: cacheToken()

  this.getToken = function() {
    var app = UiApp.getActiveApplication();
    var scriptProperties = PropertiesService.getScriptProperties();
    var tokenValue = scriptProperties.getProperty('tokenValue');
    var tokenTimeStamp = scriptProperties.getProperty('tokenTimeStamp');
    var tokenExpiry = scriptProperties.getProperty('tokenExpiry');
    var currentTime = new Date().getTime() / 1000; // converted to seconds
    // check if cached token exists, and if it has expired
    if (tokenValue != '' && tokenTimeStamp != '' && tokenExpiry != '' &&
	currentTime - tokenTimeStamp < tokenExpiry) {
      return tokenValue;
    }
    else {
      return authenticate().value;
    }
  }
  // END: getToken()

}
// END: MktoClient()
