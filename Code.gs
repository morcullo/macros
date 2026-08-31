/**
 * Macros app <-> Google Sheets bridge.
 *
 * Setup:
 * 1. Create a Google Sheet.
 * 2. Extensions -> Apps Script.
 * 3. Paste this file into Code.gs and save.
 * 4. Deploy -> New deployment -> Web app.
 * 5. Execute as: Me. Who has access: Anyone.
 * 6. Paste the /exec URL into the GOOGLE_SHEET_URL constant in index.html. No URL entry is needed in the app.
 */
const STATE_SHEET = 'AppState';
const MEALS_SHEET = 'Meals';
const FOODS_SHEET = 'Food Items';
const WEIGHT_SHEET = 'Weight';
const GOALS_SHEET = 'Goals';

function doGet(e) {
  const state = readState_();
  const callback = e && e.parameter && e.parameter.callback;
  if (callback) return ContentService.createTextOutput(callback + '(' + JSON.stringify(state) + ')').setMimeType(ContentService.MimeType.JAVASCRIPT);
  return json_(state);
}

function doPost(e) {
  try {
    const raw = e && e.parameter && e.parameter.state;
    if (!raw) throw new Error('Missing state');
    const state = JSON.parse(raw);
    writeState_(state);
    return json_({ok:true});
  } catch (err) {
    return json_({ok:false, error:String(err)});
  }
}

function readState_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(STATE_SHEET);
  if (!sheet || sheet.getLastRow() < 2) return {__empty:true};
  const raw = sheet.getRange(2, 2).getValue();
  if (!raw) return {__empty:true};
  try { return JSON.parse(String(raw)); } catch (_) { return {__empty:true}; }
}

function writeState_(state) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const stateSheet = getOrCreate_(ss, STATE_SHEET);
  stateSheet.clearContents();
  stateSheet.getRange(1,1,2,2).setValues([
    ['updatedAt','stateJson'],
    [new Date(), JSON.stringify(state)]
  ]);
  stateSheet.getRange(1,1,1,2).setFontWeight('bold');
  stateSheet.setColumnWidth(1, 160);
  stateSheet.setColumnWidth(2, 700);

  writeMeals_(ss, state.meals || []);
  writeFoods_(ss, state.meals || []);
  writeWeights_(ss, state.weights || []);
  writeGoals_(ss, state.goals || {});
}

function writeMeals_(ss, meals) {
  const sh = getOrCreate_(ss, MEALS_SHEET); sh.clearContents();
  const rows = [['ID','Date','Time','Meal Type','Meal Name','Calories','Protein (g)','Carbs (g)','Fat (g)','Food Items']];
  meals.forEach(m => {
    const t = totals_(m.items || []);
    rows.push([m.id || '', m.date || '', m.time || '', m.mealType || '', m.name || '', t.calories, t.protein, t.carbs, t.fat, (m.items || []).length]);
  });
  sh.getRange(1,1,rows.length,rows[0].length).setValues(rows); sh.getRange(1,1,1,rows[0].length).setFontWeight('bold');
  sh.setFrozenRows(1); sh.autoResizeColumns(1, rows[0].length);
}

function writeFoods_(ss, meals) {
  const sh = getOrCreate_(ss, FOODS_SHEET); sh.clearContents();
  const rows = [['Meal ID','Meal Name','Food Item','Calories','Protein (g)','Carbs (g)','Fat (g)']];
  meals.forEach(m => (m.items || []).forEach(i => rows.push([m.id || '', m.name || '', i.name || '', +i.calories || 0, +i.protein || 0, +i.carbs || 0, +i.fat || 0])));
  sh.getRange(1,1,rows.length,rows[0].length).setValues(rows); sh.getRange(1,1,1,rows[0].length).setFontWeight('bold');
  sh.setFrozenRows(1); sh.autoResizeColumns(1, rows[0].length);
}

function writeWeights_(ss, weights) {
  const sh = getOrCreate_(ss, WEIGHT_SHEET); sh.clearContents();
  const rows = [['Date','Weight (lbs)']].concat(weights.map(w => [w.date || '', +w.value || 0]));
  sh.getRange(1,1,rows.length,2).setValues(rows); sh.getRange(1,1,1,2).setFontWeight('bold'); sh.setFrozenRows(1); sh.autoResizeColumns(1,2);
}

function writeGoals_(ss, goals) {
  const sh = getOrCreate_(ss, GOALS_SHEET); sh.clearContents();
  const rows = [['Goal','Value'],['Calories',+goals.calories||0],['Protein',+goals.protein||0],['Carbs',+goals.carbs||0],['Fat',+goals.fat||0]];
  sh.getRange(1,1,rows.length,2).setValues(rows); sh.getRange(1,1,1,2).setFontWeight('bold'); sh.autoResizeColumns(1,2);
}

function totals_(items) {
  return items.reduce((a,i) => ({
    calories:a.calories+(+i.calories||0), protein:a.protein+(+i.protein||0),
    carbs:a.carbs+(+i.carbs||0), fat:a.fat+(+i.fat||0)
  }), {calories:0,protein:0,carbs:0,fat:0});
}
function getOrCreate_(ss, name) { return ss.getSheetByName(name) || ss.insertSheet(name); }
function json_(obj) { return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON); }
