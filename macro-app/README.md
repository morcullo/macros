# Macros — Google Sheets Edition

This version keeps a local browser copy and syncs the app state to a Google Sheet through Google Apps Script.

## Configure Google Sheets (one-time)

1. Create a new Google Sheet.
2. In the sheet, open **Extensions → Apps Script**.
3. Replace the contents of `Code.gs` with the included `Code.gs` file.
4. Save the project.
5. Choose **Deploy → New deployment**.
6. Select **Web app**.
7. Set **Execute as: Me**.
8. Set **Who has access: Anyone**.
9. Deploy and authorize the script if Google asks.
10. Copy the deployment URL ending in `/exec`.
11. Open `index.html` in a text editor and find:

   `const GOOGLE_SHEET_URL='';`

12. Paste your URL between the quotes, for example:

   `const GOOGLE_SHEET_URL='https://script.google.com/macros/s/XXXXXXXXXXXX/exec';`

That's it. **You do not need to enter or configure the URL anywhere inside the app.**

If you change Google Sheets deployments later, update the URL in `index.html` and redeploy/reload the app.

The script creates these tabs automatically:

- `AppState` — complete app backup used for reliable round-trip syncing
- `Meals` — one row per meal with totals
- `Food Items` — one row per food item
- `Weight` — weight history in lbs
- `Goals` — daily macro goals

## Sync behavior

- The app still keeps a local browser copy as a fallback.
- On startup, if `GOOGLE_SHEET_URL` is configured, the app loads the sheet's existing app state.
- Adding/editing/deleting meals, adding weight, and saving goals trigger a Google Sheets sync.
- The readable tabs are regenerated from the app state after each save, so they stay consistent.

### Important

The Apps Script web app is set to **Anyone** so the static app can communicate with it without a Google sign-in flow. Treat the deployment URL as private; anyone who obtains it could potentially send data to that sheet. For a personal tracker, this is the simplest setup. A future OAuth-based version can provide stronger access control.
