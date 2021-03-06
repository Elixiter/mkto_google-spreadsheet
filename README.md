mkto_google-spreadsheet
=======================

![Marketo Static Lists in your Google Spreadsheet]
(http://f.cl.ly/items/0e3T470p1i3L14351I3p/mkto_google-spreadsheet.png)

Import Marketo Static Lists to Google Spreadsheet - and maybe one day,
keep them up-to-date - using the [REST API]
(http://developers.marketo.com/documentation/rest/) and 
[Google Apps Script](https://developers.google.com/apps-script/).

**NOTE: DO NOT POST YOUR API CREDENTIALS PUBLICALLY!**  
**You alone are responsible for the security of your API credentials.**

This project is a work-in-progress. It is intended for the "new"
version of Google Spreadsheet (green check in bottom bar).
It is only minimally functional.
Please salt heavily, and see the Caveats / Todo section at the bottom of this README.

Author information
------------------
Kyle Halstvedt  
[Elixiter, Inc.](http://www.elixiter.com)  
kyle@elixiter.com

Setup (*IMPORTANT*)
-------------------
1. From within a new Google Spreadsheet file,
   open Tools -> Script editor...
2. Create a new "Blank Project" from the wizard
   or from File -> New.
3. Paste the code from
   [mkto_spreadsheet-lists.gs]
   (https://raw.githubusercontent.com/Elixiter/mkto_google-spreadsheet/master/mkto_spreadsheet-lists.gs)
   into the editor. Make sure you paste over any sample code in the editor.
4. Configure the script by replacing the
   "REPLACE_ME" strings in the configuration
   section at the top with URL's/keys that
   are specific to your Marketo instance.
   If you have trouble with this, see
   [Marketo's documentation](http://developers.marketo.com/documentation/rest/),
   specifically the sections on [authentication]
   (http://developers.marketo.com/documentation/rest/authentication/)
   and [the custom service]
   (http://developers.marketo.com/documentation/rest/custom-service/).
5. Save the script with the floppy icon or File -> Save.

Usage
-----
1. Refresh the Spreadsheets window after editing script.
2. From the open Spreadsheet with the script
   included, click the custom "Marketo Import"
   menu item, and click "Initialize sidebar...". If there
   is no menu item titled "Marketo Import", refresh the page, and
   please make sure you have followed all of the steps in Setup.
   In particular, ensure that the script is present in the script editor
   for the current document, and that you have filled in your API credentials.
3. The first time you run it, you must give authorization to
   make external http requests.
4. A right-hand sidebar should appear with the names
   and IDs of the Static Lists fetched from Marketo.
5. Each list has an "Insert" button. Click it to insert
   all members of the list into a new spreadsheet.

Caveats / Todo
--------------
1. Any user with enough priveleges to run the script
   is able to *read* the script, which contains
   your REST API credentials (ID and secret key) in-the-clear.  
   **DO NOT POST YOUR API CREDENTIALS PUBLICALLY!**
2. The Marketo API currently only permits fetching Static Lists. Sorry,
   no Smart Lists. :(
3. The script does not attempt to "update" or "synchronize" the lists,
   it currently only creates new sheets. Updating lists is a planned feature.
4. I am using the atomic Sheet.appendRow() method to add each row, which is safe
   but slow. Planned migration to Range.setValues() for performance.
5. If you exceed the API quota of 100 requests per 20 seconds, you will receive
   an error message. This script does not currently support fetching lists
   of greater than 10k leads, as it has no regulator to prevent reaching
   the API limit and no way to resume fetching a list in the middle.
6. The UX is quite poor: there are no status or loading indicators. Improving
   this is on the long-term roadmap.
