/**
 * SLIDE — Apps Script backend
 * ---------------------------------------------------------------
 * Serves index.html and backs:
 *   - registerUser / loginUser        → "Users" tab
 *   - getOwnerDashboard / createShift /
 *     updateShift / setShiftStatus    → "Shifts" tab
 *   - proposeRate / respondToRateProposal → rate bargaining on "Shifts" tab
 *   - updateProfile                   → "Users" tab
 *   - uploadImage                     → Google Drive (logo + café photos)
 *
 * ⚠️ PROTOTYPE ONLY — passwords are stored in PLAIN TEXT in the
 * sheet. Fine for a demo where nobody enters a real password, not
 * remotely secure. Before this touches real user data, replace
 * registerUser/loginUser with proper auth (hashed + salted
 * passwords at minimum, ideally a real auth provider).
 * ---------------------------------------------------------------
 */

const SPREADSHEET_ID = '1HSJmuKbyL2lxXs5kjsJJ3EM5TyhKQv3R2xrgbdoL1UI';

const USERS_SHEET = 'Users';
const USERS_HEADERS = [
  'UserID', 'Role', 'FullName', 'Email', 'Password', 'CafeName', 'Phone', 'CreatedAt',
  'Bio', 'Address', 'LogoUrl', 'InstagramUrl', 'FacebookUrl', 'WebsiteUrl', 'CafeImages',
  // Running rating totals — RatingCount/RatingTotal let us compute an
  // average without rescanning every shift. Applies to both roles: a
  // barista's rating is the average of OwnerRatingOfBarista values left
  // on their completed shifts; a café's is the average of
  // BaristaRatingOfCafe values. Incremented by rateShift(), never
  // decremented (no edit/delete-a-rating flow yet).
  'RatingTotal', 'RatingCount'
];

const SHIFTS_SHEET = 'Shifts';
const SHIFTS_HEADERS = [
  'ShiftID', 'OwnerUserID', 'CafeName', 'Date', 'StartTime', 'EndTime', 'Rate',
  'Skills', 'BusyLevel', 'LunchIncluded', 'BreakMinutes', 'BreakNotes',
  'Status', 'AssignedBaristaName', 'AssignedBaristaUserID', 'CreatedAt',
  // Free-text details from the café about this specific shift (parking,
  // who to ask for, dress code, whatever) — visible to the barista
  // wherever they see the shift's details.
  'Notes',
  // Rate bargaining — a barista can counter-offer a rate on an open shift;
  // the owner then accepts (locks it in + assigns the shift) or declines
  // (clears the offer, shift stays open).
  'ProposedRate', 'ProposedByName', 'ProposedByUserID', 'NegotiationStatus',
  // Repost flow — when a barista backs out of a shift they'd already
  // claimed, we mark who cancelled it (so the owner's dashboard can tell
  // "barista cancelled" apart from a shift the owner cancelled themselves)
  // and whether the owner has already acted on that (reposted or
  // dismissed the prompt), so it doesn't keep nagging them.
  'CancelledBy', 'RepostedShiftID',
  // Post-shift ratings — each side rates the other once the shift is
  // completed. Kept per-shift (rather than just on the Users tab) so we
  // know a given shift has/hasn't been rated yet and can't be rated
  // twice; rateShift() also rolls each value into the other party's
  // running RatingTotal/RatingCount on Users.
  'BaristaRatingOfCafe', 'OwnerRatingOfBarista'
];

// The shared Drive folder every uploaded image (logos + café photos) is
// saved into, so uploads land somewhere the team can actually see and
// fetch them rather than a folder buried in whichever account ran the
// script. https://drive.google.com/drive/folders/1fhXOio7IVnr17RDaQ8fVhimAnpgYGDO4
//
// ⚠️ AUTHORIZATION: reading/writing to a folder the script didn't create
// itself requires the broad Drive scope, not the narrower drive.file
// scope Apps Script often auto-selects. Sharing the folder "anyone with
// link" does NOT grant the script permission — that's a separate thing.
// If uploads fail with a "you do not have permission to call DriveApp..."
// error, you need to:
//   1. Open the project, go to Project Settings (gear icon) and confirm
//      "Show appsscript.json manifest file in editor" is on.
//   2. Open appsscript.json and add/confirm this block:
//        "oauthScopes": [
//          "https://www.googleapis.com/auth/spreadsheets",
//          "https://www.googleapis.com/auth/drive"
//        ]
//      (the spreadsheets scope covers the Users/Shifts tabs; without an
//      explicit oauthScopes array Apps Script auto-detects scopes, and
//      it can under-detect Drive access for a folder the script didn't
//      create, which is exactly this case).
//   3. Re-authorize: open any function in the editor (e.g. debugShifts)
//      and click Run once — this triggers the consent screen for the
//      new scope. Approve it.
//   4. Redeploy the web app (Deploy > Manage deployments > Edit > New
//      version) so the running deployment picks up the new scope —
//      editing the manifest alone doesn't update a live deployment.
const UPLOAD_FOLDER_ID = '1fhXOio7IVnr17RDaQ8fVhimAnpgYGDO4';

// Hard cap on a single uploaded image, enforced both here and in the
// front end. Keeps Drive tidy and avoids Apps Script choking on huge
// base64 payloads passed over google.script.run.
const MAX_IMAGE_SIZE_BYTES = 5 * 1024 * 1024; // 5 MB

// UK National Living Wage (21+), effective 1 April 2026. This is a floor,
// not a recommendation — update it each April when the rate changes.
// https://www.gov.uk/national-minimum-wage-rates
const MINIMUM_WAGE = 12.71;

/**
 * Serves the web app.
 */
function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('SLIDE')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/* ================================================================
   Generic sheet + header helpers
   ================================================================ */

/**
 * Returns { sheet, headerIndex } for a sheet, creating it (and/or
 * appending any missing headers) if needed. headerIndex maps each
 * header name to its 0-based column index, so all the functions
 * below reference columns by name rather than by position — safe
 * to extend later without breaking existing rows.
 */
function getSheetWithHeaders_(sheetName, expectedHeaders) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    const looseMatch = ss.getSheets().find(s => s.getName().toLowerCase() === sheetName.toLowerCase());
    sheet = looseMatch || ss.insertSheet(sheetName);
  }
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(expectedHeaders);
    sheet.setFrozenRows(1);
  }

  const lastCol = sheet.getLastColumn();
  const existingHeaders = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  const headerIndex = {};
  
  // Trim headers to prevent trailing-space mismatches
  existingHeaders.forEach((h, i) => { 
    if (h) headerIndex[String(h).trim()] = i; 
  });

  const missing = expectedHeaders.filter(h => !(h in headerIndex));
  if (missing.length) {
    sheet.getRange(1, lastCol + 1, 1, missing.length).setValues([missing]);
    missing.forEach((h, i) => { headerIndex[String(h).trim()] = lastCol + i; });
  }

  return { sheet, headerIndex };
}

/**
 * Builds a full row array (matching current sheet width) from a
 * { headerName: value } map. Columns not present in valuesMap are
 * left as ''.
 */
function buildRow_(headerIndex, valuesMap) {
  const width = Object.keys(headerIndex).length;
  const row = new Array(width).fill('');
  Object.keys(valuesMap).forEach(key => {
    if (key in headerIndex) row[headerIndex[key]] = valuesMap[key];
  });
  return row;
}

/**
 * Partially updates an existing row in place (only the keys present
 * in valuesMap are touched).
 */
function updateRow_(sheet, rowNumber, headerIndex, valuesMap) {
  Object.keys(valuesMap).forEach(key => {
    if (key in headerIndex) {
      sheet.getRange(rowNumber, headerIndex[key] + 1).setValue(valuesMap[key]);
    }
  });
}

function rowToObject_(row, headerIndex) {
  const obj = {};
  Object.keys(headerIndex).forEach(key => { 
    obj[key] = row[headerIndex[key]]; 
  });
  return obj;
}

/* ================================================================
   USERS — registration, login, profile
   ================================================================ */

function findUserRow_(sheet, headerIndex, email) {
  const data = sheet.getDataRange().getValues();
  const target = String(email).toLowerCase();
  const col = headerIndex['Email'];
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][col]).toLowerCase() === target) {
      return { rowNumber: i + 1, row: data[i] };
    }
  }
  return null;
}

function findUserRowById_(sheet, headerIndex, userId) {
  const data = sheet.getDataRange().getValues();
  const col = headerIndex['UserID'];
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][col]) === String(userId)) {
      return { rowNumber: i + 1, row: data[i] };
    }
  }
  return null;
}

function userRowToSafeUser_(row, headerIndex) {
  const o = rowToObject_(row, headerIndex);
  const ratingCount = Number(o.RatingCount) || 0;
  const ratingTotal = Number(o.RatingTotal) || 0;
  return {
    userId: o.UserID,
    role: o.Role,
    fullName: o.FullName,
    email: o.Email,
    cafeName: o.CafeName,
    phone: o.Phone,
    createdAt: o.CreatedAt instanceof Date ? o.CreatedAt.toISOString() : o.CreatedAt,
    bio: o.Bio || '',
    address: o.Address || '',
    logoUrl: o.LogoUrl || '',
    instagramUrl: o.InstagramUrl || '',
    facebookUrl: o.FacebookUrl || '',
    websiteUrl: o.WebsiteUrl || '',
    cafeImages: o.CafeImages ? String(o.CafeImages).split('|').filter(Boolean) : [],
    ratingCount: ratingCount,
    ratingAverage: ratingCount > 0 ? Math.round((ratingTotal / ratingCount) * 10) / 10 : null
  };
}

/**
 * Builds a { userId: { ratingAverage, ratingCount } } map for every user
 * in one read, so callers that need to attach a rating to a bunch of
 * shifts at once (e.g. listOpenShifts attaching each café's rating)
 * don't do a lookup per row.
 */
function getUserRatingsById_() {
  const { sheet, headerIndex } = getSheetWithHeaders_(USERS_SHEET, USERS_HEADERS);
  const data = sheet.getDataRange().getValues();
  const idCol = headerIndex['UserID'];
  const totalCol = headerIndex['RatingTotal'];
  const countCol = headerIndex['RatingCount'];

  const map = {};
  for (let i = 1; i < data.length; i++) {
    const count = Number(data[i][countCol]) || 0;
    const total = Number(data[i][totalCol]) || 0;
    map[String(data[i][idCol])] = {
      ratingCount: count,
      ratingAverage: count > 0 ? Math.round((total / count) * 10) / 10 : null
    };
  }
  return map;
}

/**
 * Adds `rating` to a user's running total and bumps their count by one.
 * Used by rateShift() to roll a completed shift's rating into the other
 * party's overall average.
 */
function incrementUserRating_(userId, rating) {
  const { sheet, headerIndex } = getSheetWithHeaders_(USERS_SHEET, USERS_HEADERS);
  const found = findUserRowById_(sheet, headerIndex, userId);
  if (!found) return; // nothing we can do if the user row is missing
  const current = rowToObject_(found.row, headerIndex);
  const newTotal = (Number(current.RatingTotal) || 0) + rating;
  const newCount = (Number(current.RatingCount) || 0) + 1;
  updateRow_(sheet, found.rowNumber, headerIndex, { RatingTotal: newTotal, RatingCount: newCount });
}

function registerUser(payload) {
  try {
    const { sheet, headerIndex } = getSheetWithHeaders_(USERS_SHEET, USERS_HEADERS);
    const email = String(payload.email || '').trim().toLowerCase();
    const password = String(payload.password || '');
    const fullName = String(payload.fullName || '').trim();
    const phone = String(payload.phone || '').trim();
    const role = String(payload.role || '');

    if (!email || !password || !fullName || !role) {
      return { success: false, message: 'Missing required fields.' };
    }
    if (!phone) {
      return { success: false, message: 'A phone number is required.' };
    }
    if (password.length < 8) {
      return { success: false, message: 'Password should be at least 8 characters.' };
    }
    if (role === 'owner' && !String(payload.cafeName || '').trim()) {
      return { success: false, message: 'Café name is required for café owners.' };
    }
    if (findUserRow_(sheet, headerIndex, email)) {
      return { success: false, message: 'That email is already registered.' };
    }

    const userId = Utilities.getUuid();
    const createdAt = new Date();

    sheet.appendRow(buildRow_(headerIndex, {
      UserID: userId,
      Role: role,
      FullName: fullName,
      Email: email,
      Password: password, // plain text — see security note at top of file
      CafeName: payload.cafeName || '',
      Phone: phone,
      CreatedAt: createdAt
    }));

    return {
      success: true,
      user: {
        userId, role, fullName, email,
        cafeName: payload.cafeName || '', phone,
        createdAt: createdAt.toISOString(),
        bio: '', address: '', logoUrl: '', instagramUrl: '', facebookUrl: '', websiteUrl: '', cafeImages: [],
        ratingCount: 0, ratingAverage: null
      }
    };
  } catch (err) {
    return { success: false, message: 'Server error: ' + err.message };
  }
}

function loginUser(payload) {
  try {
    const { sheet, headerIndex } = getSheetWithHeaders_(USERS_SHEET, USERS_HEADERS);
    const email = String(payload.email || '').trim().toLowerCase();
    const password = String(payload.password || '');

    const found = findUserRow_(sheet, headerIndex, email);
    if (!found || String(found.row[headerIndex['Password']]) !== password) {
      return { success: false, message: 'Email or password is incorrect.' };
    }
    return { success: true, user: userRowToSafeUser_(found.row, headerIndex) };
  } catch (err) {
    return { success: false, message: 'Server error: ' + err.message };
  }
}

/**
 * Updates a user's editable profile fields. Works for both roles:
 *   - Barista payload:  { userId, role:'barista', fullName, phone }
 *   - Owner payload:    { userId, role:'owner', fullName, phone, cafeName,
 *                          bio, address, instagramUrl, facebookUrl,
 *                          websiteUrl, logoUrl, cafeImages: [urls...] }
 * Full name and phone are mandatory for everyone; café name is mandatory
 * for owners only. Café-specific fields are ignored for baristas.
 */
function updateProfile(payload) {
  try {
    const { sheet, headerIndex } = getSheetWithHeaders_(USERS_SHEET, USERS_HEADERS);
    const found = findUserRowById_(sheet, headerIndex, payload.userId);
    if (!found) return { success: false, message: 'User not found.' };

    const fullName = String(payload.fullName || '').trim();
    const phone = String(payload.phone || '').trim();
    if (!fullName) return { success: false, message: 'Full name is required.' };
    if (!phone) return { success: false, message: 'A phone number is required.' };

    const existingRole = rowToObject_(found.row, headerIndex).Role;
    const role = payload.role || existingRole;

    const updates = { FullName: fullName, Phone: phone };

    if (role === 'owner') {
      const cafeName = String(payload.cafeName || '').trim();
      if (!cafeName) return { success: false, message: 'Café name is required.' };
      updates.CafeName = cafeName;
      updates.Bio = payload.bio || '';
      updates.Address = payload.address || '';
      updates.LogoUrl = payload.logoUrl || '';
      updates.InstagramUrl = payload.instagramUrl || '';
      updates.FacebookUrl = payload.facebookUrl || '';
      updates.WebsiteUrl = payload.websiteUrl || '';
      updates.CafeImages = (payload.cafeImages || []).join('|');
    }

    updateRow_(sheet, found.rowNumber, headerIndex, updates);

    const updatedRow = sheet.getRange(found.rowNumber, 1, 1, sheet.getLastColumn()).getValues()[0];
    return { success: true, user: userRowToSafeUser_(updatedRow, headerIndex) };
  } catch (err) {
    return { success: false, message: 'Server error: ' + err.message };
  }
}

/* ================================================================
   SHIFTS — post, edit, list, status changes, rate bargaining
   ================================================================ */

function shiftRowToObject_(row, headerIndex) {
  const o = rowToObject_(row, headerIndex);
  
  // Safely format Date objects or ISO strings for the Date column into 'YYYY-MM-DD'
  let rawDate = o.Date;
  if (rawDate instanceof Date) {
    rawDate = Utilities.formatDate(rawDate, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  } else if (rawDate) {
    rawDate = String(rawDate).split('T')[0];
  }

  // Safely format Date/Time objects coming from Google Sheets time-only cells (1899 epoch) into 'HH:mm'
  function formatTimeCell_(val) {
    if (!val) return '';
    if (val instanceof Date) {
      return Utilities.formatDate(val, Session.getScriptTimeZone(), 'H:mm');
    }
    const str = String(val);
    // If it's an ISO string starting with the Google Sheets 1899 epoch, extract the time portion
    if (str.includes('1899-12-30') || str.includes('T')) {
      const parts = str.split('T');
      if (parts.length > 1) {
        const timePart = parts[1].replace('Z', '').substring(0, 5); // "HH:mm"
        // Remove leading zero if desired or return as-is
        return timePart.startsWith('0') ? timePart.substring(1) : timePart;
      }
    }
    return str;
  }

  return {
    shiftId: o.ShiftID,
    ownerUserId: o.OwnerUserID,
    cafeName: o.CafeName,
    date: rawDate,
    startTime: formatTimeCell_(o.StartTime),
    endTime: formatTimeCell_(o.EndTime),
    rate: Number(o.Rate) || 0,
    skills: o.Skills ? String(o.Skills).split('|').filter(Boolean) : [],
    busyLevel: Number(o.BusyLevel) || 3,
    lunchIncluded: o.LunchIncluded === true || o.LunchIncluded === 'TRUE' || o.LunchIncluded === 'true',
    breakMinutes: Number(o.BreakMinutes) || 0,
    breakNotes: o.BreakNotes || '',
    status: String(o.Status || 'open').trim().toLowerCase(),
    assignedBaristaName: o.AssignedBaristaName || '',
    assignedBaristaUserId: o.AssignedBaristaUserID || '',
    createdAt: o.CreatedAt instanceof Date ? o.CreatedAt.toISOString() : o.CreatedAt,
    proposedRate: o.ProposedRate ? Number(o.ProposedRate) : null,
    proposedByName: o.ProposedByName || '',
    proposedByUserId: o.ProposedByUserID || '',
    negotiationStatus: String(o.NegotiationStatus || '').trim().toLowerCase(),
    cancelledBy: String(o.CancelledBy || '').trim().toLowerCase(),
    repostedShiftId: o.RepostedShiftID || '',
    notes: o.Notes || '',
    baristaRatingOfCafe: o.BaristaRatingOfCafe ? Number(o.BaristaRatingOfCafe) : null,
    ownerRatingOfBarista: o.OwnerRatingOfBarista ? Number(o.OwnerRatingOfBarista) : null
  };
}

function findShiftRowById_(sheet, headerIndex, shiftId) {
  const data = sheet.getDataRange().getValues();
  const col = headerIndex['ShiftID'];
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][col]) === String(shiftId)) {
      return { rowNumber: i + 1, row: data[i] };
    }
  }
  return null;
}

/**
 * payload: { ownerUserId }
 * returns: { success, openShifts:[...], historyShifts:[...] }
 */
function getOwnerDashboard(payload) {
  try {
    // Fallback if frontend passes userId as a plain string instead of an object
    if (typeof payload === 'string') {
      payload = { ownerUserId: payload };
    }
    if (!payload || !payload.ownerUserId) {
      return { success: false, message: 'Owner User ID is missing.' };
    }

    const { sheet, headerIndex } = getSheetWithHeaders_(SHIFTS_SHEET, SHIFTS_HEADERS);
    const data = sheet.getDataRange().getValues();
    const ownerCol = headerIndex['OwnerUserID'];

    const open = [];
    const history = [];
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][ownerCol]).trim() !== String(payload.ownerUserId).trim()) continue;
      const shift = shiftRowToObject_(data[i], headerIndex);
      (shift.status === 'open' ? open : history).push(shift);
    }
    
    open.sort((a, b) => (a.date < b.date ? -1 : 1));
    history.sort((a, b) => (a.date < b.date ? 1 : -1));

    return { success: true, openShifts: open, historyShifts: history };
  } catch (err) {
    return { success: false, message: 'Server error: ' + err.message };
  }
}

/**
 * payload: { ownerUserId, cafeName, date, startTime, endTime, rate,
 *            skills:[...], busyLevel, lunchIncluded, breakMinutes, breakNotes }
 */
function createShift(payload) {
  try {
    const { sheet, headerIndex } = getSheetWithHeaders_(SHIFTS_SHEET, SHIFTS_HEADERS);

    if (!payload.date || !payload.startTime || !payload.endTime) {
      return { success: false, message: 'Date, start time, and end time are required.' };
    }
    if (!payload.rate || Number(payload.rate) <= 0) {
      return { success: false, message: 'Enter a valid hourly rate.' };
    }
    if (Number(payload.rate) < MINIMUM_WAGE) {
      return { success: false, message: `Rate can't be below the minimum wage (£${MINIMUM_WAGE.toFixed(2)}/hr).` };
    }
    if (!payload.skills || !payload.skills.length) {
      return { success: false, message: 'Select at least one skill needed for this shift.' };
    }

    const shiftId = Utilities.getUuid();
    const createdAt = new Date();

    sheet.appendRow(buildRow_(headerIndex, {
      ShiftID: shiftId,
      OwnerUserID: payload.ownerUserId,
      CafeName: payload.cafeName || '',
      Date: payload.date,
      StartTime: payload.startTime,
      EndTime: payload.endTime,
      Rate: payload.rate,
      Skills: (payload.skills || []).join('|'),
      BusyLevel: payload.busyLevel || 3,
      LunchIncluded: !!payload.lunchIncluded,
      BreakMinutes: payload.breakMinutes || 0,
      BreakNotes: payload.breakNotes || '',
      Notes: payload.notes || '',
      Status: 'open',
      AssignedBaristaName: '',
      CreatedAt: createdAt
    }));

    return {
      success: true,
      shift: {
        shiftId, ownerUserId: payload.ownerUserId, cafeName: payload.cafeName || '',
        date: payload.date, startTime: payload.startTime, endTime: payload.endTime,
        rate: payload.rate, skills: payload.skills, busyLevel: payload.busyLevel || 3,
        lunchIncluded: !!payload.lunchIncluded, breakMinutes: payload.breakMinutes || 0,
        breakNotes: payload.breakNotes || '', notes: payload.notes || '', status: 'open', assignedBaristaName: '',
        createdAt: createdAt.toISOString()
      }
    };
  } catch (err) {
    return { success: false, message: 'Server error: ' + err.message };
  }
}

/**
 * payload: same shape as createShift, plus { shiftId }.
 */
function updateShift(payload) {
  try {
    const { sheet, headerIndex } = getSheetWithHeaders_(SHIFTS_SHEET, SHIFTS_HEADERS);
    const found = findShiftRowById_(sheet, headerIndex, payload.shiftId);
    if (!found) return { success: false, message: 'Shift not found.' };

    if (payload.rate !== undefined && Number(payload.rate) < MINIMUM_WAGE) {
      return { success: false, message: `Rate can't be below the minimum wage (£${MINIMUM_WAGE.toFixed(2)}/hr).` };
    }

    updateRow_(sheet, found.rowNumber, headerIndex, {
      Date: payload.date,
      StartTime: payload.startTime,
      EndTime: payload.endTime,
      Rate: payload.rate,
      Skills: (payload.skills || []).join('|'),
      BusyLevel: payload.busyLevel || 3,
      LunchIncluded: !!payload.lunchIncluded,
      BreakMinutes: payload.breakMinutes || 0,
      BreakNotes: payload.breakNotes || '',
      Notes: payload.notes || ''
    });

    const updatedRow = sheet.getRange(found.rowNumber, 1, 1, sheet.getLastColumn()).getValues()[0];
    return { success: true, shift: shiftRowToObject_(updatedRow, headerIndex) };
  } catch (err) {
    return { success: false, message: 'Server error: ' + err.message };
  }
}

/**
 * payload: { shiftId, status: 'open'|'filled'|'completed'|'cancelled', assignedBaristaName? }
 */
function setShiftStatus(payload) {
  try {
    const { sheet, headerIndex } = getSheetWithHeaders_(SHIFTS_SHEET, SHIFTS_HEADERS);
    const found = findShiftRowById_(sheet, headerIndex, payload.shiftId);
    if (!found) return { success: false, message: 'Shift not found.' };

    const updates = { Status: payload.status };
    if (payload.assignedBaristaName !== undefined) {
      updates.AssignedBaristaName = payload.assignedBaristaName;
    }
    // This endpoint is only ever driven by the owner's own dashboard
    // controls, so a cancellation coming through here is always
    // owner-initiated — tag it as such, distinct from a barista backing
    // out of a shift (see cancelShiftAsBarista), so the dashboard knows
    // not to show a "barista cancelled — repost?" prompt for it.
    if (payload.status === 'cancelled') {
      updates.CancelledBy = 'owner';
    }
    updateRow_(sheet, found.rowNumber, headerIndex, updates);

    const updatedRow = sheet.getRange(found.rowNumber, 1, 1, sheet.getLastColumn()).getValues()[0];
    return { success: true, shift: shiftRowToObject_(updatedRow, headerIndex) };
  } catch (err) {
    return { success: false, message: 'Server error: ' + err.message };
  }
}

/**
 * A barista backs out of a shift they'd already claimed. Rather than
 * silently reopening it (which could surprise the owner, or clash with
 * other arrangements they've made since), this just marks the shift
 * cancelled-by-barista and leaves it to the owner's dashboard to prompt
 * "repost this shift?" — see repostShift() / dismissRepostPrompt().
 * payload: { shiftId, baristaUserId }
 */
function cancelShiftAsBarista(payload) {
  try {
    const { sheet, headerIndex } = getSheetWithHeaders_(SHIFTS_SHEET, SHIFTS_HEADERS);
    const found = findShiftRowById_(sheet, headerIndex, payload.shiftId);
    if (!found) return { success: false, message: 'This shift no longer exists.' };

    const current = shiftRowToObject_(found.row, headerIndex);
    if (current.status !== 'filled') {
      return { success: false, message: 'This shift is not currently assigned to you.' };
    }
    if (String(current.assignedBaristaUserId) !== String(payload.baristaUserId)) {
      return { success: false, message: 'This shift is not assigned to you.' };
    }

    updateRow_(sheet, found.rowNumber, headerIndex, {
      Status: 'cancelled',
      CancelledBy: 'barista'
      // Deliberately leave AssignedBaristaName/UserID in place — the
      // owner's dashboard uses it to show "cancelled by <name>".
    });

    const updatedRow = sheet.getRange(found.rowNumber, 1, 1, sheet.getLastColumn()).getValues()[0];
    return { success: true, shift: shiftRowToObject_(updatedRow, headerIndex) };
  } catch (err) {
    return { success: false, message: 'Server error: ' + err.message };
  }
}

/**
 * Returns the shifts a barista needs to see on their "your shifts" list:
 * anything currently assigned to them (status 'filled', for cancelling)
 * plus anything they've completed but haven't rated the café for yet.
 * Upcoming shifts first (soonest date), then completed ones (most
 * recent first).
 * payload: { baristaUserId }
 */
function getBaristaShifts(payload) {
  try {
    if (typeof payload === 'string') payload = { baristaUserId: payload };
    if (!payload || !payload.baristaUserId) {
      return { success: false, message: 'Barista User ID is missing.' };
    }

    const { sheet, headerIndex } = getSheetWithHeaders_(SHIFTS_SHEET, SHIFTS_HEADERS);
    const data = sheet.getDataRange().getValues();
    const assignedCol = headerIndex['AssignedBaristaUserID'];
    const ratingsById = getUserRatingsById_();

    const filled = [];
    const completed = [];
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][assignedCol]).trim() !== String(payload.baristaUserId).trim()) continue;
      const shift = shiftRowToObject_(data[i], headerIndex);
      const cafeRating = ratingsById[shift.ownerUserId] || { ratingAverage: null, ratingCount: 0 };
      shift.cafeRatingAverage = cafeRating.ratingAverage;
      shift.cafeRatingCount = cafeRating.ratingCount;
      if (shift.status === 'filled') filled.push(shift);
      // Only surface completed shifts that still need a rating from the
      // barista — once rated, there's no reason for it to keep showing
      // up here indefinitely.
      else if (shift.status === 'completed' && shift.baristaRatingOfCafe === null) completed.push(shift);
    }
    filled.sort((a, b) => (a.date === b.date ? (a.startTime < b.startTime ? -1 : 1) : (a.date < b.date ? -1 : 1)));
    completed.sort((a, b) => (a.date === b.date ? 0 : (a.date < b.date ? 1 : -1)));

    return { success: true, shifts: filled.concat(completed) };
  } catch (err) {
    return { success: false, message: 'Server error: ' + err.message };
  }
}

/**
 * Owner taps "Repost this shift" after a barista cancelled. Rather than
 * quietly un-cancelling the original row (which would erase the record
 * that it was ever cancelled), this creates a fresh open shift with the
 * same details and links back to it, so the cancelled row stays in
 * history as an accurate record and doesn't keep triggering the prompt.
 * payload: { shiftId } — the cancelled shift being reposted.
 */
function repostShift(payload) {
  try {
    const { sheet, headerIndex } = getSheetWithHeaders_(SHIFTS_SHEET, SHIFTS_HEADERS);
    const found = findShiftRowById_(sheet, headerIndex, payload.shiftId);
    if (!found) return { success: false, message: 'Shift not found.' };

    const original = shiftRowToObject_(found.row, headerIndex);
    if (original.status !== 'cancelled' || original.cancelledBy !== 'barista') {
      return { success: false, message: 'Only a barista-cancelled shift can be reposted.' };
    }
    if (original.repostedShiftId) {
      return { success: false, message: 'This shift has already been reposted.' };
    }

    const newShiftId = Utilities.getUuid();
    const createdAt = new Date();

    sheet.appendRow(buildRow_(headerIndex, {
      ShiftID: newShiftId,
      OwnerUserID: original.ownerUserId,
      CafeName: original.cafeName,
      Date: original.date,
      StartTime: original.startTime,
      EndTime: original.endTime,
      Rate: original.rate,
      Skills: (original.skills || []).join('|'),
      BusyLevel: original.busyLevel,
      LunchIncluded: !!original.lunchIncluded,
      BreakMinutes: original.breakMinutes,
      BreakNotes: original.breakNotes,
      Notes: original.notes,
      Status: 'open',
      AssignedBaristaName: '',
      CreatedAt: createdAt
    }));

    updateRow_(sheet, found.rowNumber, headerIndex, { RepostedShiftID: newShiftId });

    return {
      success: true,
      shift: {
        shiftId: newShiftId, ownerUserId: original.ownerUserId, cafeName: original.cafeName,
        date: original.date, startTime: original.startTime, endTime: original.endTime,
        rate: original.rate, skills: original.skills, busyLevel: original.busyLevel,
        lunchIncluded: original.lunchIncluded, breakMinutes: original.breakMinutes,
        breakNotes: original.breakNotes, notes: original.notes, status: 'open', assignedBaristaName: '',
        createdAt: createdAt.toISOString()
      }
    };
  } catch (err) {
    return { success: false, message: 'Server error: ' + err.message };
  }
}

/**
 * Owner dismisses the "repost this shift?" prompt without reposting.
 * Marks it handled (via a sentinel in RepostedShiftID) so it stops
 * showing up as needing attention.
 * payload: { shiftId }
 */
function dismissRepostPrompt(payload) {
  try {
    const { sheet, headerIndex } = getSheetWithHeaders_(SHIFTS_SHEET, SHIFTS_HEADERS);
    const found = findShiftRowById_(sheet, headerIndex, payload.shiftId);
    if (!found) return { success: false, message: 'Shift not found.' };

    updateRow_(sheet, found.rowNumber, headerIndex, { RepostedShiftID: 'dismissed' });
    const updatedRow = sheet.getRange(found.rowNumber, 1, 1, sheet.getLastColumn()).getValues()[0];
    return { success: true, shift: shiftRowToObject_(updatedRow, headerIndex) };
  } catch (err) {
    return { success: false, message: 'Server error: ' + err.message };
  }
}

/**
 * Records a 1–5 star rating on a completed shift, from either side, and
 * rolls it into the other party's running average on the Users tab.
 * Each side can only rate a given shift once — the shift row itself is
 * the record of whether that's already happened.
 * payload: { shiftId, raterRole: 'barista'|'owner', raterUserId, rating }
 */
function rateShift(payload) {
  try {
    const { sheet, headerIndex } = getSheetWithHeaders_(SHIFTS_SHEET, SHIFTS_HEADERS);
    const found = findShiftRowById_(sheet, headerIndex, payload.shiftId);
    if (!found) return { success: false, message: 'Shift not found.' };

    const rating = Math.round(Number(payload.rating));
    if (!rating || rating < 1 || rating > 5) {
      return { success: false, message: 'Rating must be between 1 and 5 stars.' };
    }

    const shift = shiftRowToObject_(found.row, headerIndex);
    if (shift.status !== 'completed') {
      return { success: false, message: 'Only completed shifts can be rated.' };
    }

    if (payload.raterRole === 'barista') {
      if (String(shift.assignedBaristaUserId) !== String(payload.raterUserId)) {
        return { success: false, message: 'This shift is not assigned to you.' };
      }
      if (shift.baristaRatingOfCafe !== null) {
        return { success: false, message: "You've already rated this shift." };
      }
      updateRow_(sheet, found.rowNumber, headerIndex, { BaristaRatingOfCafe: rating });
      incrementUserRating_(shift.ownerUserId, rating);
    } else if (payload.raterRole === 'owner') {
      if (String(shift.ownerUserId) !== String(payload.raterUserId)) {
        return { success: false, message: "This isn't your shift to rate." };
      }
      if (shift.ownerRatingOfBarista !== null) {
        return { success: false, message: "You've already rated this shift." };
      }
      if (!shift.assignedBaristaUserId) {
        return { success: false, message: 'This shift has no assigned barista to rate.' };
      }
      updateRow_(sheet, found.rowNumber, headerIndex, { OwnerRatingOfBarista: rating });
      incrementUserRating_(shift.assignedBaristaUserId, rating);
    } else {
      return { success: false, message: 'Unknown rater role.' };
    }

    const updatedRow = sheet.getRange(found.rowNumber, 1, 1, sheet.getLastColumn()).getValues()[0];
    return { success: true, shift: shiftRowToObject_(updatedRow, headerIndex) };
  } catch (err) {
    return { success: false, message: 'Server error: ' + err.message };
  }
}

/**
 * Barista counter-offers a rate on an open shift, instead of accepting the
 * café's posted rate outright. The offer can never undercut the minimum
 * wage — the front end pre-fills and floors the input at MINIMUM_WAGE, and
 * this is re-checked here since the client can't be trusted.
 * payload: { shiftId, baristaUserId, baristaName, proposedRate }
 */
function proposeRate(payload) {
  try {
    const { sheet, headerIndex } = getSheetWithHeaders_(SHIFTS_SHEET, SHIFTS_HEADERS);
    const found = findShiftRowById_(sheet, headerIndex, payload.shiftId);
    if (!found) return { success: false, message: 'This shift no longer exists.' };

    const current = shiftRowToObject_(found.row, headerIndex);
    if (current.status !== 'open') {
      return { success: false, message: 'This shift is no longer open.' };
    }

    const proposedRate = Number(payload.proposedRate);
    if (!proposedRate || proposedRate <= 0) {
      return { success: false, message: 'Enter a valid hourly rate.' };
    }
    if (proposedRate < MINIMUM_WAGE) {
      return { success: false, message: `Your rate can't be below the minimum wage (£${MINIMUM_WAGE.toFixed(2)}/hr).` };
    }

    updateRow_(sheet, found.rowNumber, headerIndex, {
      ProposedRate: proposedRate,
      ProposedByName: payload.baristaName || '',
      ProposedByUserID: payload.baristaUserId || '',
      NegotiationStatus: 'pending'
    });

    const updatedRow = sheet.getRange(found.rowNumber, 1, 1, sheet.getLastColumn()).getValues()[0];
    return { success: true, shift: shiftRowToObject_(updatedRow, headerIndex) };
  } catch (err) {
    return { success: false, message: 'Server error: ' + err.message };
  }
}

/**
 * Owner accepts or declines a barista's rate offer.
 *   - accept: locks in the proposed rate as the shift's rate, assigns the
 *     shift to that barista, and marks it filled.
 *   - decline: clears the offer; the shift stays open at its original rate
 *     for anyone (including the same barista) to try again.
 * payload: { shiftId, accept: boolean }
 */
function respondToRateProposal(payload) {
  try {
    const { sheet, headerIndex } = getSheetWithHeaders_(SHIFTS_SHEET, SHIFTS_HEADERS);
    const found = findShiftRowById_(sheet, headerIndex, payload.shiftId);
    if (!found) return { success: false, message: 'Shift not found.' };

    const current = shiftRowToObject_(found.row, headerIndex);
    if (current.negotiationStatus !== 'pending') {
      return { success: false, message: 'There is no pending rate offer on this shift.' };
    }

    if (payload.accept) {
      updateRow_(sheet, found.rowNumber, headerIndex, {
        Rate: current.proposedRate,
        Status: 'filled',
        AssignedBaristaName: current.proposedByName,
        AssignedBaristaUserID: current.proposedByUserId,
        NegotiationStatus: 'accepted'
      });
    } else {
      updateRow_(sheet, found.rowNumber, headerIndex, {
        NegotiationStatus: 'declined',
        ProposedRate: '',
        ProposedByName: '',
        ProposedByUserID: ''
      });
    }

    const updatedRow = sheet.getRange(found.rowNumber, 1, 1, sheet.getLastColumn()).getValues()[0];
    return { success: true, shift: shiftRowToObject_(updatedRow, headerIndex) };
  } catch (err) {
    return { success: false, message: 'Server error: ' + err.message };
  }
}

/**
 * Returns all shifts with status 'open', across all owners, soonest first.
 * This is what a barista sees when browsing for shifts to pick up.
 */
function listOpenShifts() {
  try {
    const { sheet, headerIndex } = getSheetWithHeaders_(SHIFTS_SHEET, SHIFTS_HEADERS);
    const data = sheet.getDataRange().getValues();
    const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
    const ratingsById = getUserRatingsById_();

    const open = [];
    for (let i = 1; i < data.length; i++) {
      const shift = shiftRowToObject_(data[i], headerIndex);
      if (shift.status === 'open' && String(shift.date) >= today) {
        const cafeRating = ratingsById[shift.ownerUserId] || { ratingAverage: null, ratingCount: 0 };
        shift.cafeRatingAverage = cafeRating.ratingAverage;
        shift.cafeRatingCount = cafeRating.ratingCount;
        open.push(shift);
      }
    }
    open.sort((a, b) => (a.date === b.date ? (a.startTime < b.startTime ? -1 : 1) : (a.date < b.date ? -1 : 1)));

    return { success: true, shifts: open, minimumWage: MINIMUM_WAGE };
  } catch (err) {
    return { success: false, message: 'Server error: ' + err.message };
  }
}

/**
 * Claims an open shift for a barista. Re-checks status at write time
 * so two baristas can't both claim the same shift.
 * payload: { shiftId, baristaUserId, baristaName }
 */
function acceptShift(payload) {
  try {
    const { sheet, headerIndex } = getSheetWithHeaders_(SHIFTS_SHEET, SHIFTS_HEADERS);
    const found = findShiftRowById_(sheet, headerIndex, payload.shiftId);
    if (!found) return { success: false, message: 'This shift no longer exists.' };

    const current = shiftRowToObject_(found.row, headerIndex);
    if (current.status !== 'open') {
      return { success: false, message: 'This shift was just taken by someone else.' };
    }

    updateRow_(sheet, found.rowNumber, headerIndex, {
      Status: 'filled',
      AssignedBaristaName: payload.baristaName || '',
      AssignedBaristaUserID: payload.baristaUserId || ''
    });

    const updatedRow = sheet.getRange(found.rowNumber, 1, 1, sheet.getLastColumn()).getValues()[0];
    return { success: true, shift: shiftRowToObject_(updatedRow, headerIndex) };
  } catch (err) {
    return { success: false, message: 'Server error: ' + err.message };
  }
}

/**
 * Returns every account with Role === 'barista' from the Users tab —
 * used for the owner's "Find a barista" directory, both the per-shift
 * booking flow and the dashboard-level browse (e.g. planning ahead
 * for time off, with no specific shift picked yet).
 *
 * Only ever returns safe, non-password fields. Real accounts only have
 * what was collected at sign-up (name, email, phone) — no rating,
 * distance, or skills yet, since that was never captured.
 */
function listBaristas() {
  try {
    const { sheet, headerIndex } = getSheetWithHeaders_(USERS_SHEET, USERS_HEADERS);
    const data = sheet.getDataRange().getValues();
    const roleCol = headerIndex['Role'];

    const baristas = [];
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][roleCol]) !== 'barista') continue;
      const safe = userRowToSafeUser_(data[i], headerIndex);
      baristas.push(safe);
    }
    baristas.sort((a, b) => (a.fullName || '').localeCompare(b.fullName || ''));

    return { success: true, baristas };
  } catch (err) {
    return { success: false, message: 'Server error: ' + err.message };
  }
}

/* ================================================================
   IMAGE UPLOAD — café logo + photos, stored in Drive
   ================================================================ */

/**
 * Returns the shared uploads folder by its fixed ID so every image lands
 * in a place the team can browse and that stays reachable regardless of
 * which account's Drive the script runs under.
 *
 * NOTE: DriveApp calls on a folder the script didn't create itself need
 * the broad Drive scope (https://www.googleapis.com/auth/drive), not the
 * narrower drive.file scope — sharing the folder "anyone with link" does
 * NOT substitute for this. If this throws a permission/authorization
 * error, see the oauthScopes block in appsscript.json and the
 * re-authorization note above it.
 */
function getUploadsFolder_() {
  try {
    return DriveApp.getFolderById(UPLOAD_FOLDER_ID);
  } catch (err) {
    throw new Error(
      'Could not open the uploads folder (ID ' + UPLOAD_FOLDER_ID + '). ' +
      'This is almost always a scope/authorization issue, not a sharing setting — ' +
      'see the oauthScopes note in appsscript.json. Original error: ' + err.message
    );
  }
}

/**
 * payload: { base64, mimeType, fileName }
 * returns: { success, url }
 */
function uploadImage(payload) {
  try {
    if (!payload.base64 || !payload.mimeType) {
      return { success: false, message: 'No image data received.' };
    }

    const bytes = Utilities.base64Decode(payload.base64);
    if (bytes.length > MAX_IMAGE_SIZE_BYTES) {
      const maxMb = (MAX_IMAGE_SIZE_BYTES / (1024 * 1024)).toFixed(0);
      return { success: false, message: `Image is too large — please use a file under ${maxMb}MB.` };
    }

    const blob = Utilities.newBlob(bytes, payload.mimeType, payload.fileName || 'upload');

    const folder = getUploadsFolder_();
    const file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

    // The classic "uc?export=view" link is unreliable for hotlinking as an
    // <img src> these days; the thumbnail endpoint fetches consistently for
    // publicly-shared files.
    const url = 'https://drive.google.com/thumbnail?id=' + file.getId() + '&sz=w1000';
    return { success: true, url };
  } catch (err) {
    return { success: false, message: 'Upload failed: ' + err.message };
  }
}

/**
 * TEMPORARY — run this once from the Apps Script editor (select it in the
 * function dropdown, then click Run) to force the Drive authorization
 * consent screen. debugShifts() never calls DriveApp, so running that
 * doesn't prompt for Drive access even after the scope is in the
 * manifest — Apps Script only asks you to approve a scope when a call
 * that needs it actually executes. Delete this function once uploads
 * are working.
 */
function testDriveAccess() {
  const folder = DriveApp.getFolderById(UPLOAD_FOLDER_ID);
  Logger.log('Drive access OK — folder name: ' + folder.getName());
}

function debugShifts() {
  const { sheet, headerIndex } = getSheetWithHeaders_(SHIFTS_SHEET, SHIFTS_HEADERS);
  Logger.log("Header Index Map: " + JSON.stringify(headerIndex));
  
  const data = sheet.getDataRange().getValues();
  Logger.log("Total rows in sheet: " + data.length);
  
  if (data.length > 1) {
    Logger.log("First row raw: " + JSON.stringify(data[1]));
    const parsed = shiftRowToObject_(data[1], headerIndex);
    Logger.log("Parsed first shift object: " + JSON.stringify(parsed));
  }
}