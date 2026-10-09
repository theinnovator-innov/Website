/**
 * INNOV v1.0 — Google Apps Script backend
 *
 * Setup:
 * 1. Set SPREADSHEET_ID and RESPONSE_SHEET_GID in Project Settings > Script properties.
 * 2. Deploy as a Web app (execute as you; access restricted to your organization where possible).
 * 3. Configure Microsoft Graph properties before enabling Outlook email sending.
 *
 * Important: this starter intentionally does NOT send a receipt automatically.
 */

const CONFIG = {
  // Existing response spreadsheet ID from the provided URL:
  SPREADSHEET_ID: '1spanhY8DgDMqQb9-NOvf7-mKeU-BTF2rwkk8T7rSH5I',
  // Existing Google Form response tab GID from the provided URL:
  RESPONSE_SHEET_GID: 71866861,
  SHEETS: {
    EVENTS: 'INNOV_Events',
    MEMBERS: 'INNOV_Members',
    ASSIGNMENTS: 'INNOV_Assignments',
    EMAIL_LOG: 'INNOV_EmailLog'
  },
  ROLES: ['Photographer', 'Writer', 'Graphics & Layout', 'Videographer']
};

function doGet() {
  setupSheets_();
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('INNOV | Event Management')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/** Called from the browser to initialize the dashboard. */
function getAppData() {
  setupSheets_();
  const events = readObjects_(CONFIG.SHEETS.EVENTS);
  const assignments = readObjects_(CONFIG.SHEETS.ASSIGNMENTS);
  const members = readObjects_(CONFIG.SHEETS.MEMBERS);
  const currentEmail = getCurrentUserEmail_();

  return {
    events: events.map(e => decorateEvent_(e, assignments)),
    assignments: assignments,
    members: members,
    roles: CONFIG.ROLES,
    currentEmail: currentEmail,
    emailConfigured: isOutlookConfigured_()
  };
}

/**
 * Imports new source form rows into INNOV_Events without overwriting decisions.
 * Header matching is heuristic. Verify the mapped columns in the source sheet.
 */
function syncFormResponses() {
  setupSheets_();
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const source = ss.getSheets().find(s => s.getSheetId() === Number(CONFIG.RESPONSE_SHEET_GID));
  if (!source) throw new Error('Could not find the form-response tab. Check RESPONSE_SHEET_GID.');
  const values = source.getDataRange().getDisplayValues();
  if (values.length < 2) return { imported: 0, message: 'No form responses found.' };

  const headers = values[0].map(h => String(h).trim().toLowerCase());
  const col = {
    timestamp: findHeader_(headers, ['timestamp', 'submitted']),
    organization: findHeader_(headers, ['organization', 'org name', 'student organization', 'club']),
    contact: findHeader_(headers, ['contact person', 'name of requester', 'requester', 'representative']),
    email: findHeader_(headers, ['email address', 'email', 'e-mail']),
    title: findHeader_(headers, ['event title', 'event name', 'name of event', 'title']),
    date: findHeader_(headers, ['event date', 'date of event', 'date']),
    venue: findHeader_(headers, ['venue', 'location', 'place']),
    description: findHeader_(headers, ['description', 'details', 'purpose']),
    raw: -1
  };

  if (col.title < 0) {
    throw new Error('Could not identify the event title column. Update the header matching in syncFormResponses().');
  }

  const eventSheet = ss.getSheetByName(CONFIG.SHEETS.EVENTS);
  const existing = readObjects_(CONFIG.SHEETS.EVENTS);
  const existingSourceRows = new Set(existing.map(e => String(e.sourceRow || '')));
  let imported = 0;

  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    if (!row.some(v => String(v).trim())) continue;
    const sourceRow = String(i + 1);
    if (existingSourceRows.has(sourceRow)) continue;
    const title = valueAt_(row, col.title);
    if (!title) continue;

    appendObject_(eventSheet, {
      eventId: 'EVT-' + Utilities.getUuid().slice(0, 8).toUpperCase(),
      sourceRow: sourceRow,
      submittedAt: valueAt_(row, col.timestamp),
      organization: valueAt_(row, col.organization),
      contactPerson: valueAt_(row, col.contact),
      email: valueAt_(row, col.email),
      title: title,
      eventDate: valueAt_(row, col.date),
      venue: valueAt_(row, col.venue),
      description: valueAt_(row, col.description),
      status: 'Pending Review',
      rejectionReason: '',
      receiptSent: 'No',
      receiptSentAt: '',
      receiptSentBy: '',
      createdAt: new Date().toISOString()
    });
    imported++;
  }
  return { imported: imported, message: imported + ' new request(s) imported.' };
}

function reviewEvent(eventId, decision, rejectionReason) {
  if (!['Approved for Staffing', 'Rejected'].includes(decision)) throw new Error('Invalid decision.');
  const event = findEvent_(eventId);
  if (!event) throw new Error('Event not found.');
  if (event.status !== 'Pending Review') throw new Error('Only pending requests can be reviewed.');

  if (decision === 'Rejected') {
    if (!String(rejectionReason || '').trim()) throw new Error('Please enter a rejection reason.');
    if (!event.email) throw new Error('The request has no sender email address.');
    sendOutlookEmail_(event.email, 'Update on your event request: ' + event.title,
      rejectionEmailBody_(event, rejectionReason));
    updateObject_(CONFIG.SHEETS.EVENTS, 'eventId', eventId, {
      status: 'Rejected',
      rejectionReason: String(rejectionReason).trim()
    });
    logEmail_(eventId, event.email, 'Rejection', getCurrentUserEmail_(), 'Sent');
  } else {
    updateObject_(CONFIG.SHEETS.EVENTS, 'eventId', eventId, { status: decision });
  }
  return getAppData();
}

function volunteerForEvent(eventId, role, memberName, memberEmail) {
  if (!CONFIG.ROLES.includes(role)) throw new Error('Choose a valid coverage role.');
  if (!String(memberName || '').trim() || !String(memberEmail || '').trim()) {
    throw new Error('Enter your name and email.');
  }
  const event = findEvent_(eventId);
  if (!event || event.status !== 'Approved for Staffing') throw new Error('This event is not open for staffing.');

  const assignments = readObjects_(CONFIG.SHEETS.ASSIGNMENTS);
  const duplicate = assignments.some(a =>
    String(a.eventId) === String(eventId) &&
    String(a.memberEmail).toLowerCase() === String(memberEmail).toLowerCase()
  );
  if (duplicate) throw new Error('You have already volunteered for this event.');

  appendObject_(SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID).getSheetByName(CONFIG.SHEETS.ASSIGNMENTS), {
    assignmentId: 'ASN-' + Utilities.getUuid().slice(0, 8).toUpperCase(),
    eventId: eventId,
    memberName: String(memberName).trim(),
    memberEmail: String(memberEmail).trim().toLowerCase(),
    role: role,
    createdAt: new Date().toISOString()
  });
  return getAppData();
}

/** Manual-only receipt sending. Called only by the board's Send Receipt button. */
function sendAcknowledgementReceipt(eventId, sendDespiteVacancies) {
  const event = findEvent_(eventId);
  if (!event) throw new Error('Event not found.');
  if (!['Approved for Staffing', 'Receipt Sent'].includes(event.status)) {
    throw new Error('Approve the event for staffing before sending a receipt.');
  }
  if (String(event.receiptSent).toLowerCase() === 'yes') {
    throw new Error('An acknowledgement receipt has already been sent for this event.');
  }
  if (!event.email) throw new Error('The request has no sender email address.');

  const assignments = readObjects_(CONFIG.SHEETS.ASSIGNMENTS).filter(a => String(a.eventId) === String(eventId));
  const rolesFilled = new Set(assignments.map(a => a.role));
  const vacantRoles = CONFIG.ROLES.filter(role => !rolesFilled.has(role));
  if (vacantRoles.length && !sendDespiteVacancies) {
    return { needsConfirmation: true, vacantRoles: vacantRoles };
  }

  const lineup = CONFIG.ROLES.map(role => {
    const names = assignments.filter(a => a.role === role).map(a => a.memberName);
    return role + ': ' + (names.length ? names.join(', ') : 'Vacant');
  }).join('\n');

  sendOutlookEmail_(event.email, 'Acknowledgement Receipt: ' + event.title,
    receiptEmailBody_(event, lineup, vacantRoles));
  const now = new Date().toISOString();
  updateObject_(CONFIG.SHEETS.EVENTS, 'eventId', eventId, {
    status: 'Receipt Sent',
    receiptSent: 'Yes',
    receiptSentAt: now,
    receiptSentBy: getCurrentUserEmail_()
  });
  logEmail_(eventId, event.email, 'Acknowledgement Receipt', getCurrentUserEmail_(), 'Sent');
  return { success: true, data: getAppData() };
}

function setupSheets_() {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  ensureSheet_(ss, CONFIG.SHEETS.EVENTS, [
    'eventId','sourceRow','submittedAt','organization','contactPerson','email','title','eventDate',
    'venue','description','status','rejectionReason','receiptSent','receiptSentAt','receiptSentBy','createdAt'
  ]);
  ensureSheet_(ss, CONFIG.SHEETS.MEMBERS, ['memberId','name','email','role','active']);
  ensureSheet_(ss, CONFIG.SHEETS.ASSIGNMENTS, ['assignmentId','eventId','memberName','memberEmail','role','createdAt']);
  ensureSheet_(ss, CONFIG.SHEETS.EMAIL_LOG, ['emailLogId','eventId','recipient','type','sentAt','sentBy','result']);
}

function ensureSheet_(ss, name, headers) {
  let sheet = ss.getSheetByName(name);
  if (!sheet) sheet = ss.insertSheet(name);
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
  }
}

function readObjects_(sheetName) {
  const sheet = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID).getSheetByName(sheetName);
  if (!sheet || sheet.getLastRow() < 2) return [];
  const values = sheet.getDataRange().getValues();
  const headers = values[0].map(String);
  return values.slice(1).filter(row => row.some(v => v !== '')).map(row => {
    const obj = {};
    headers.forEach((h, i) => obj[h] = row[i] instanceof Date ? row[i].toISOString() : row[i]);
    return obj;
  });
}

function appendObject_(sheet, obj) {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(String);
  sheet.appendRow(headers.map(h => obj[h] === undefined ? '' : obj[h]));
}

function updateObject_(sheetName, keyName, keyValue, changes) {
  const sheet = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID).getSheetByName(sheetName);
  const values = sheet.getDataRange().getValues();
  const headers = values[0].map(String);
  const keyCol = headers.indexOf(keyName);
  if (keyCol < 0) throw new Error('Missing key column ' + keyName);
  for (let r = 1; r < values.length; r++) {
    if (String(values[r][keyCol]) === String(keyValue)) {
      Object.keys(changes).forEach(k => {
        const c = headers.indexOf(k);
        if (c >= 0) sheet.getRange(r + 1, c + 1).setValue(changes[k]);
      });
      return;
    }
  }
  throw new Error('Record not found: ' + keyValue);
}

function findEvent_(eventId) {
  return readObjects_(CONFIG.SHEETS.EVENTS).find(e => String(e.eventId) === String(eventId));
}

function decorateEvent_(event, assignments) {
  const rows = assignments.filter(a => String(a.eventId) === String(event.eventId));
  const rolesFilled = new Set(rows.map(a => a.role));
  const vacantRoles = CONFIG.ROLES.filter(role => !rolesFilled.has(role));
  return Object.assign({}, event, {
    assignments: rows,
    vacantRoles: vacantRoles,
    filledRoleCount: CONFIG.ROLES.length - vacantRoles.length,
    totalRoles: CONFIG.ROLES.length,
    lineupComplete: vacantRoles.length === 0
  });
}

function findHeader_(headers, candidates) {
  for (const candidate of candidates) {
    const index = headers.findIndex(h => h === candidate || h.includes(candidate));
    if (index >= 0) return index;
  }
  return -1;
}
function valueAt_(row, index) { return index >= 0 ? String(row[index] || '').trim() : ''; }
function getCurrentUserEmail_() {
  try { return Session.getActiveUser().getEmail() || 'Unknown user'; }
  catch (e) { return 'Unknown user'; }
}

function isOutlookConfigured_() {
  const p = PropertiesService.getScriptProperties();
  return ['MS_TENANT_ID','MS_CLIENT_ID','MS_CLIENT_SECRET','MS_SENDER_EMAIL'].every(k => !!p.getProperty(k));
}

function sendOutlookEmail_(to, subject, body) {
  const p = PropertiesService.getScriptProperties();
  const tenant = p.getProperty('MS_TENANT_ID');
  const clientId = p.getProperty('MS_CLIENT_ID');
  const secret = p.getProperty('MS_CLIENT_SECRET');
  const sender = p.getProperty('MS_SENDER_EMAIL');
  if (!tenant || !clientId || !secret || !sender) {
    throw new Error('Outlook is not configured yet. Set the Microsoft Graph Script Properties in the setup guide.');
  }

  const tokenResponse = UrlFetchApp.fetch('https://login.microsoftonline.com/' + encodeURIComponent(tenant) + '/oauth2/v2.0/token', {
    method: 'post',
    payload: {
      client_id: clientId,
      client_secret: secret,
      scope: 'https://graph.microsoft.com/.default',
      grant_type: 'client_credentials'
    },
    muteHttpExceptions: true
  });
  const tokenData = JSON.parse(tokenResponse.getContentText() || '{}');
  if (tokenResponse.getResponseCode() >= 300 || !tokenData.access_token) {
    throw new Error('Could not authenticate with Microsoft Graph: ' + (tokenData.error_description || tokenResponse.getContentText()));
  }

  const response = UrlFetchApp.fetch('https://graph.microsoft.com/v1.0/users/' + encodeURIComponent(sender) + '/sendMail', {
    method: 'post',
    contentType: 'application/json',
    headers: { Authorization: 'Bearer ' + tokenData.access_token },
    payload: JSON.stringify({
      message: {
        subject: subject,
        body: { contentType: 'Text', content: body },
        toRecipients: [{ emailAddress: { address: to } }]
      },
      saveToSentItems: true
    }),
    muteHttpExceptions: true
  });
  if (response.getResponseCode() >= 300) {
    throw new Error('Outlook email failed (' + response.getResponseCode() + '): ' + response.getContentText());
  }
}

function rejectionEmailBody_(event, reason) {
  return 'Dear ' + (event.contactPerson || 'Organization representative') + ',\n\n' +
    'Thank you for submitting your event request for "' + event.title + '". After review, the Editorial Board is unable to approve the request at this time.\n\n' +
    'Reason: ' + reason + '\n\nRegards,\nEditorial Board';
}
function receiptEmailBody_(event, lineup, vacantRoles) {
  return 'Dear ' + (event.contactPerson || 'Organization representative') + ',\n\n' +
    'This email serves as acknowledgement of your event request "' + event.title + '".\n\n' +
    'Organization: ' + (event.organization || 'Not provided') + '\n' +
    'Event date: ' + (event.eventDate || 'Not provided') + '\n' +
    'Venue: ' + (event.venue || 'Not provided') + '\n\n' +
    'Coverage lineup:\n' + lineup + '\n\n' +
    (vacantRoles.length ? 'Note: The following coverage roles are currently vacant: ' + vacantRoles.join(', ') + '.\n\n' : '') +
    'Thank you,\nEditorial Board';
}
function logEmail_(eventId, recipient, type, sentBy, result) {
  appendObject_(SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID).getSheetByName(CONFIG.SHEETS.EMAIL_LOG), {
    emailLogId: 'MAIL-' + Utilities.getUuid().slice(0, 8).toUpperCase(),
    eventId: eventId, recipient: recipient, type: type,
    sentAt: new Date().toISOString(), sentBy: sentBy, result: result
  });
}
