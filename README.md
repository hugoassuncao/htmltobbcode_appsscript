# htmltobbcode_appsscript
Convert HTML to BBCode in Google Sheets

This is a Google Apps Script projects that adds a custom function to Google Sheets and let's you convert HTML code into BBCode.
That's helpfull when you have CSV HTML content downloaded from one source and must convert it to BBCode to upload somewhere else.

To use it, create a new SpreadSheet in Google Sheets, go to Extensions > Apps Script, to create a bounded script, and replace all the content of the file with this code.
Once done this and saved the script, go back to the spreadsheet and use the HTMLTOBBCODE() formula to convert each cell or use the HTMLTOBBCODERANGE() to perform a conversion of a full range (use with the ARRAYFORMULA()).
