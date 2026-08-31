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
const FOOD_LIBRARY_SHEET = 'Food Library';
const WEIGHT_SHEET = 'Weight';
const GOALS_SHEET = 'Goals';

function doGet(e) {
  const action = e && e.parameter && e.parameter.action;
  const cb = e && e.parameter && e.parameter.callback;
  if (action === 'estimateFood') return estimateFood_(e);
  const state = readState_();
  if (cb) return ContentService.createTextOutput(cb + '(' + JSON.stringify(state) + ')').setMimeType(ContentService.MimeType.JAVASCRIPT);
  return json_(state);
}


function estimateFood_(e) {
  const cb = e && e.parameter && e.parameter.callback;
  const food = String((e && e.parameter && e.parameter.food) || '').trim();
  let result;
  try {
    if (!food) throw new Error('Enter a food item first.');
    const key = PropertiesService.getScriptProperties().getProperty('OPENAI_API_KEY');
    if (!key) throw new Error('OpenAI API key is not configured in Apps Script.');
    const prompt = [
      'Estimate nutrition for the food item below.',
      'Return ONLY valid JSON with exactly these keys: name, calories, protein, carbs, fat, assumption.',
      'Calories and macros should be for the serving/quantity stated by the user.',
      'If no serving size is stated, use a reasonable standard serving and explain that assumption briefly.',
      'Use numeric values for calories, protein, carbs, and fat. Do not include units in numeric fields.',
      'Do not invent a brand unless the user provided one.',
      'Food item: ' + food
    ].join('\n');
    const payload = {
      model: 'gpt-5-mini',
      input: prompt,
      max_output_tokens: 250
    };
    const response = UrlFetchApp.fetch('https://api.openai.com/v1/responses', {
      method: 'post',
      contentType: 'application/json',
      headers: {Authorization: 'Bearer ' + key},
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });
    const code = response.getResponseCode();
    const body = response.getContentText();
    if (code < 200 || code >= 300) throw new Error('OpenAI request failed (' + code + ').');
    const parsed = JSON.parse(body);
    let text = parsed.output_text || '';
    if (!text && Array.isArray(parsed.output)) {
      parsed.output.forEach(item => {
        if (item && item.type === 'message' && Array.isArray(item.content)) {
          item.content.forEach(part => { if (part && part.type === 'output_text') text += part.text || ''; });
        }
      });
    }
    const jsonText = String(text).trim().replace(/^```json\s*/i,'').replace(/\s*```$/,'');
    const foodData = JSON.parse(jsonText);
    const clean = {
      name: String(foodData.name || food),
      calories: Math.max(0, Number(foodData.calories) || 0),
      protein: Math.max(0, Number(foodData.protein) || 0),
      carbs: Math.max(0, Number(foodData.carbs) || 0),
      fat: Math.max(0, Number(foodData.fat) || 0),
      assumption: String(foodData.assumption || '')
    };
    result = {ok:true, food:clean};
  } catch (err) {
    result = {ok:false, error:String(err && err.message || err)};
  }
  const out = JSON.stringify(result);
  if (cb) return ContentService.createTextOutput(cb + '(' + out + ')').setMimeType(ContentService.MimeType.JAVASCRIPT);
  return json_(result);
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    const raw = e && e.parameter && e.parameter.state;
    if (!raw) throw new Error('Missing state');
    const incoming = JSON.parse(raw);
    lock.waitLock(10000);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const existing = readState_();
    const incomingStamp = Number(incoming && incoming._updatedAt || 0);
    const existingStamp = Number(existing && existing._updatedAt || 0);
    // Never let an older browser overwrite a newer cloud copy.
    if (!existing.__empty && existingStamp > incomingStamp) {
      return json_({ok:true, ignored:true, updatedAt:existingStamp});
    }
    writeState_(incoming);
    return json_({ok:true, updatedAt:incomingStamp});
  } catch (err) {
    return json_({ok:false, error:String(err)});
  } finally {
    try { lock.releaseLock(); } catch (_) {}
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
  writeFoodLibrary_(ss, state.foodLibrary || []);
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


function writeFoodLibrary_(ss, foods) {
  const sh = getOrCreate_(ss, FOOD_LIBRARY_SHEET); sh.clearContents();
  const seen = new Set();
  const rows = [['Food','Carbs','Protein','Fat','Calories']];
  (foods || []).forEach(i => {
    const name = String(i.name || '').trim();
    const key = name.toLowerCase();
    if (!name || seen.has(key)) return;
    seen.add(key);
    rows.push([name, +i.carbs || 0, +i.protein || 0, +i.fat || 0, +i.calories || 0]);
  });
  sh.getRange(1,1,rows.length,5).setValues(rows); sh.getRange(1,1,1,5).setFontWeight('bold');
  sh.setFrozenRows(1); sh.autoResizeColumns(1,5);
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
