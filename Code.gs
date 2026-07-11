/**
 * RC Location Map - Apps Script Backend
 *
 * SETUP:
 * 1. Go to script.google.com -> New Project, paste this code in as Code.gs
 * 2. In the Apps Script editor: Project Settings -> Script Properties -> add:
 *      GOOGLE_CLIENT_ID = <your OAuth client ID>.apps.googleusercontent.com
 * 3. In your Google Sheet, the Authorized_Users tab now has TWO extra columns
 *    to support CVC who don't have email (login by Name + PIN instead of Google):
 *
 *    Tab name: Authorized_Users
 *    Columns:  Email | Name | Role | Added By | Date Added | Login Code | PIN
 *    - For Google-login users (like yourself): fill in Email, leave Login Code/PIN blank.
 *    - For PIN-login users (CVC with no email): leave Email blank, fill in a short
 *      Login Code (e.g. "CVC01") and a PIN (e.g. a 4-6 digit number). Login Code must
 *      be unique per person.
 *
 *    Tab name: RC_Locations
 *    Columns:  Child ID | Latitude | Longitude | Pinned By | Date Pinned | Notes
 *
 *    Tab name: Child_Roster
 *    Columns:  Child ID | Name | District | Commune | Village | Photo URL
 *
 *    Tab name: Village_Centers
 *    Columns:  District | Commune | Village | Latitude | Longitude | Added By | Date Added
 *
 * 4. Set SHEET_ID below to your spreadsheet's ID.
 * 5. Deploy -> Manage deployments -> New version -> Deploy after any change to this file.
 */

const SHEET_ID = '1pYP1Ke3YjPUwaCks5Thqv_gBP7M3K8ZZIxQLlD9cXQQ';
const USERS_TAB = 'Authorized_Users';
const LOCATIONS_TAB = 'RC_Locations';
const ROSTER_TAB = 'Child_Roster';
const VILLAGE_CENTERS_TAB = 'Village_Centers';

function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    const action = body.action;

    if (action === 'login') {
      return handleLogin(body);
    } else if (action === 'getLoginOptions') {
      return handleGetLoginOptions();
    } else if (action === 'addLocation') {
      return handleAddLocation(body);
    } else if (action === 'deleteLocation') {
      return handleDeleteLocation(body);
    } else if (action === 'setVillageCenter') {
      return handleSetVillageCenter(body);
    } else {
      return jsonResponse({ ok: false, error: 'Unknown action' });
    }
  } catch (err) {
    return jsonResponse({ ok: false, error: 'Server error: ' + err.message });
  }
}

// ---------------------------------------------------------------------------
// Authentication: supports EITHER Google Sign-In (idToken) OR Name+PIN login
// (loginCode + pin). Every write action calls authenticate(body) first.
// ---------------------------------------------------------------------------

function verifyGoogleToken(idToken) {
  const url = 'https://oauth2.googleapis.com/tokeninfo?id_token=' + encodeURIComponent(idToken);
  const resp = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
  if (resp.getResponseCode() !== 200) return null;
  const data = JSON.parse(resp.getContentText());

  const expectedClientId = PropertiesService.getScriptProperties().getProperty('GOOGLE_CLIENT_ID');
  if (expectedClientId && data.aud !== expectedClientId) return null;

  if (!data.email || data.email_verified !== 'true') return null;
  return data.email.toLowerCase();
}

function getAllUserRows() {
  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(USERS_TAB);
  return sheet.getDataRange().getValues();
  // Columns: Email | Name | Role | Added By | Date Added | Login Code | PIN
}

function getAuthorizedUserByEmail(email) {
  const rows = getAllUserRows();
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]).toLowerCase().trim() === email) {
      return { email: email, name: rows[i][1], role: rows[i][2], loginCode: rows[i][5] || '' };
    }
  }
  return null;
}

function getAuthorizedUserByPin(loginCode, pin) {
  const rows = getAllUserRows();
  const codeTrimmed = String(loginCode || '').trim().toLowerCase();
  const pinTrimmed = String(pin || '').trim();
  for (let i = 1; i < rows.length; i++) {
    const rowCode = String(rows[i][5] || '').trim().toLowerCase();
    const rowPin = String(rows[i][6] || '').trim();
    if (rowCode && rowCode === codeTrimmed && rowPin && rowPin === pinTrimmed) {
      return { email: rows[i][0] || '', name: rows[i][1], role: rows[i][2], loginCode: rows[i][5] };
    }
  }
  return null;
}

// Main entry point used by every action that needs to know "who is this?".
// Returns { user } on success, or { error } on failure. Never throws.
function authenticate(body) {
  if (body.idToken) {
    const email = verifyGoogleToken(body.idToken);
    if (!email) return { error: 'Could not verify Google login.' };
    const user = getAuthorizedUserByEmail(email);
    if (!user) return { error: 'This account is not authorized to access RC locations.' };
    return { user: user };
  }
  if (body.loginCode && body.pin) {
    const user = getAuthorizedUserByPin(body.loginCode, body.pin);
    if (!user) return { error: 'Incorrect name or PIN.' };
    return { user: user };
  }
  return { error: 'No login credentials provided.' };
}

function handleGetLoginOptions() {
  const rows = getAllUserRows();
  const options = [];
  for (let i = 1; i < rows.length; i++) {
    const loginCode = rows[i][5];
    if (loginCode) {
      options.push({ loginCode: loginCode, name: rows[i][1] });
    }
  }
  return jsonResponse({ ok: true, options: options });
}

// ---------------------------------------------------------------------------

function getRoster() {
  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(ROSTER_TAB);
  const rows = sheet.getDataRange().getValues();
  const list = [];
  const byId = {};
  // Columns: Child ID | Name | District | Commune | Village | Photo URL
  for (let i = 1; i < rows.length; i++) {
    if (!rows[i][0]) continue;
    const entry = {
      childId: String(rows[i][0]).trim(),
      name: rows[i][1],
      district: rows[i][2],
      commune: rows[i][3],
      village: rows[i][4],
      photoUrl: rows[i][5] || ''
    };
    list.push(entry);
    byId[entry.childId] = entry;
  }
  return { list, byId };
}

function getVillageCenters() {
  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(VILLAGE_CENTERS_TAB);
  const rows = sheet.getDataRange().getValues();
  const list = [];
  // Columns: District | Commune | Village | Latitude | Longitude | Added By | Date Added
  for (let i = 1; i < rows.length; i++) {
    if (!rows[i][2]) continue;
    list.push({
      district: rows[i][0],
      commune: rows[i][1],
      village: rows[i][2],
      lat: rows[i][3],
      lng: rows[i][4]
    });
  }
  return list;
}

function handleLogin(body) {
  const auth = authenticate(body);
  if (auth.error) return jsonResponse({ ok: false, error: auth.error });
  const user = auth.user;

  const roster = getRoster();

  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(LOCATIONS_TAB);
  const rows = sheet.getDataRange().getValues();
  const locations = [];
  for (let i = 1; i < rows.length; i++) {
    if (!rows[i][0]) continue;
    const childId = String(rows[i][0]).trim();
    const rosterEntry = roster.byId[childId];
    locations.push({
      childId: childId,
      childName: rosterEntry ? rosterEntry.name : '',
      childPhotoUrl: rosterEntry ? rosterEntry.photoUrl : '',
      lat: rows[i][1],
      lng: rows[i][2],
      pinnedBy: rows[i][3],
      datePinned: rows[i][4],
      notes: rows[i][5] || ''
    });
  }

  return jsonResponse({ ok: true, user: user, locations: locations, roster: roster.list, villageCenters: getVillageCenters() });
}

function handleAddLocation(body) {
  const auth = authenticate(body);
  if (auth.error) return jsonResponse({ ok: false, error: auth.error });
  const user = auth.user;

  const childId = String(body.childId || '').trim();
  const lat = parseFloat(body.lat);
  const lng = parseFloat(body.lng);
  const notes = String(body.notes || '').trim();

  if (!childId || isNaN(lat) || isNaN(lng)) {
    return jsonResponse({ ok: false, error: 'Missing or invalid Child ID / coordinates.' });
  }

  const pinnedByLabel = user.email || (user.name + ' (' + user.loginCode + ')');
  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(LOCATIONS_TAB);
  sheet.appendRow([childId, lat, lng, pinnedByLabel, new Date(), notes]);

  return jsonResponse({ ok: true });
}

function handleDeleteLocation(body) {
  const auth = authenticate(body);
  if (auth.error) return jsonResponse({ ok: false, error: auth.error });

  const childId = String(body.childId || '').trim();
  const lat = parseFloat(body.lat);
  const lng = parseFloat(body.lng);

  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(LOCATIONS_TAB);
  const rows = sheet.getDataRange().getValues();

  for (let i = 1; i < rows.length; i++) {
    const rowChildId = String(rows[i][0]).trim();
    const rowLat = parseFloat(rows[i][1]);
    const rowLng = parseFloat(rows[i][2]);
    if (rowChildId === childId && Math.abs(rowLat - lat) < 0.0001 && Math.abs(rowLng - lng) < 0.0001) {
      sheet.deleteRow(i + 1);
      return jsonResponse({ ok: true });
    }
  }

  return jsonResponse({ ok: false, error: 'Could not find that exact pin to delete (it may have already been removed).' });
}

function handleSetVillageCenter(body) {
  const auth = authenticate(body);
  if (auth.error) return jsonResponse({ ok: false, error: auth.error });
  const user = auth.user;
  if (String(user.role).toLowerCase() !== 'admin') {
    return jsonResponse({ ok: false, error: 'Only Admins can set village center points.' });
  }

  const district = String(body.district || '').trim();
  const commune = String(body.commune || '').trim();
  const village = String(body.village || '').trim();
  const lat = parseFloat(body.lat);
  const lng = parseFloat(body.lng);

  if (!district || !commune || !village || isNaN(lat) || isNaN(lng)) {
    return jsonResponse({ ok: false, error: 'Missing district/commune/village or coordinates.' });
  }

  const addedByLabel = user.email || (user.name + ' (' + user.loginCode + ')');
  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(VILLAGE_CENTERS_TAB);
  const rows = sheet.getDataRange().getValues();

  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]).trim() === district &&
        String(rows[i][1]).trim() === commune &&
        String(rows[i][2]).trim() === village) {
      sheet.getRange(i + 1, 4).setValue(lat);
      sheet.getRange(i + 1, 5).setValue(lng);
      sheet.getRange(i + 1, 6).setValue(addedByLabel);
      sheet.getRange(i + 1, 7).setValue(new Date());
      return jsonResponse({ ok: true, updated: true });
    }
  }

  sheet.appendRow([district, commune, village, lat, lng, addedByLabel, new Date()]);
  return jsonResponse({ ok: true, updated: false });
}

function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
