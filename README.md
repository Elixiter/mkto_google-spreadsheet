mkto_google-spreadsheet
=======================

A repository of Google Apps Script for integrating with Marketo API's

// Author information
// --------------------
// Kyle Halstvedt
// Elixiter, Inc.
// August 20, 2014
// kyle@elixiter.com
// --------------------

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

// Todo / Caveats
// --------------------
// 1. Paginated API; currently only retrieves
//    the first 100 results from a list.
// --------------------
