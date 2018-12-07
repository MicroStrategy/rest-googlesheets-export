# mstr-googlesheets-export
The following Google Script enables exporting MicroStrategy visualizations from a Dossier Chart or Grid to a Google Spreadsheet on Google Drive.

# Quick demonstration
[![IMAGE ALT TEXT HERE](https://img.youtube.com/vi/zQF2P51wAz0/0.jpg)](https://www.youtube.com/watch?v=zQF2P51wAz0)

# Setup
1. Create a project on https://scripts.google.com
2. Add the 2 .js files from this repo as '.gs' (instead of '.js') to the project
3. Update config.gs with MicroStrategy standard authentication username and password (leave all the other fields empty)
4. Go to Resources > Advanced Google Services : enable Google Sheets API
5. From previous screen, click on Google Cloud Platform API Dashboard. and also enable Sheets API for the project
6. Use menu *Publish* > *Deploy as web app* and publish the script
7. Should give you a url like https://script.google.com/macros/s/someweiredtechnicalidentifier/exec

Use the following scheme to run an export to a new Spreadsheet: https://script.google.com/macros/s/someweiredtechnicalidentifier/exec?searchvisname=YourVisualizationTitle&dossierurl=DossierURLinLibrary
Use the following scheme to run an export to an existing Spreadsheet: https://script.google.com/macros/s/someweiredtechnicalidentifier/exec?searchvisname=YourVisualizationTitle&dossierurl=DossierURLinLibrary&spreadsheetId=idOfSpreadsheet

# Limitations, Caveats, Next steps
This exporter has been designed for simple extracts: 1 Attribute and 1 Metric. It does not work for every visualizations. Successfully tested: Grid, KPI Widget
It does not yet support Google Authentication, or End User pass-through authentication. It uses instead a MicroStrategy account hardcoded into the config.gs file
