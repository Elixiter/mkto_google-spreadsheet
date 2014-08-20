mkto_google-spreadsheet
=======================

A repository of Google Apps Script for integrating Spreadsheet with Marketo API's

__NOTE: DO NOT POST YOUR API CREDENTIALS PUBLICALLY__  
__You alone are responsible for the security of your API credentials__

This project is a work-in-progress. It is only minimally functional. Please salt heavily, and see the Todo / Caveats section at the bottom of this README.

Author information
--------------------
Kyle Halstvedt  
[Elixiter, Inc.](http://www.elixiter.com)  
kyle@elixiter.com

Setup (*IMPORTANT*)
--------------------
1. From within a new Google Spreadsheet file,
   open Tools -> Script editor...
2. File -> New -> Project, and select
  Create a script for: Spreadsheet
3. Paste ALL of the code in this file into
   the editor
4. CONFIGURE the script by replacing the
  "REPLACE_ME" strings in the configuration
  section at the top with URL's/keys that
  are specific to your Marketo instance.
  If you have trouble with this, see
  Marketo's documentation at
  http://developers.marketo.com/documentation/rest/

Usage
--------------------
1. From a Spreadsheet file with the script
  included, click the new "Marketo Import"
  menu item, and click "Initialize sidebar..."
2. A right-hand sidebar should appear with the names
  and IDs of the lists fetched from Marketo
3. Each list has an "Insert" button. Click it to insert
  the members of the list.

Todo / Caveats
--------------------
1. Paginated API; currently only retrieves
  the first 100 results from a list.
2. Any user with enough priveleges to run the script
   would be able to *read* the script, which contains
   your REST API credentials (ID and secret key) in-the-clear.  
   __DO NOT POST YOUR API CREDENTIALS PUBLICALLY__
