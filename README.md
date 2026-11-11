# Kati Roll by Bombay Spice - Kendall Park Sign Up

Static sign-up page (index.html, logo.png, food.jpg). Deploy on Vercel as a static site, no build step.

Each sign-up does two things:
1. Emails the details through Web3Forms (access key is in the form in index.html).
2. Adds a row to a Google Sheet through a Google Apps Script web app (optional, URL goes in SHEET_ENDPOINT near the bottom of index.html).

## Spreadsheet setup
1. Create a new Google Sheet (for example "Kati Roll Kendall Park Sign Ups").
2. Extensions > Apps Script. Delete the sample code and paste in apps-script.gs. Save.
3. Deploy > New deployment > type "Web app". Execute as: Me. Who has access: Anyone. Deploy and approve the permissions.
4. Copy the Web app URL (ends in /exec) and paste it into SHEET_ENDPOINT in index.html.
5. Share the Sheet with the owner. File > Download > Microsoft Excel (.xlsx) gives an Excel copy any time.
