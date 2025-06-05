/**
 * CSULMCS Loan Management System Script
 * A magical toy box automating loan applications with Google Sheets and Forms!
 * Handles Intent Form sign-ups, Application Form processing, Finance Officer assignments, and notifications.
 * Features:
 * - Applicants sign up with two guarantors via Intent Form.
 * - One email to applicant with GroupID post-Intent submission.
 * - Admins assign Finance Officers per group in Control tab, or auto-assign from FinanceOfficers sheet.
 * - Applications pend until Finance Officer is assigned.
 * - 7-day guarantor reminder countdown starts post-assignment.
 * - Manual notification trigger for newly assigned groups.
 * - Intent tab: One row per participant (applicant + guarantors).
 * - Control tab: Tracks loan details with statuses.
 * - Archive, FinanceOfficers, Logs tabs for completed apps, officer data, and actions.
 */

// --- SETTINGS ---
const CONTROL_SHEET_NAME = 'Control';
const ARCHIVE_SHEET_NAME = 'Archive';
const LOAN_INTENT_SHEET_NAME = 'Intent';
const FINANCE_OFFICERS_SHEET_NAME = 'FinanceOfficers';
const LOGS_SHEET_NAME = 'Logs';
const REMINDER_DAYS_LIMIT = 7;
const ROLE_APPLICANT = 'Applicant';
const ROLE_FINANCE = 'Finance Officer';
const ROLE_GUARANTOR = 'Guarantor';
const SENDER_NAME = 'CSULMCS Finance Team';
const EMAIL_FOOTER = `
  <hr style="border: 1px solid #ccc; margin: 20px 0;">
  <p style="font-size: 12px; color: #333; text-align: center;">
    Sent by the <strong>CSULMCS Loan Management System</strong>.<br>
    © 2025 CSULMCS. All rights reserved.
  </p>`;
const EMAIL_REGARDS = `<p style="font-family: Arial, sans-serif;">Regards,<br>CSULMCS Finance Team</p>`;

// --- HELPERS ---

/**
 * Gets column index by header in a sheet.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet The sheet to search.
 * @param {string} header Column header.
 * @returns {number} 1-based column index.
 */
function getColumnIndex(sheet, header) {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const index = headers.indexOf(header);
  if (index === -1) throw new Error(`Header '${header}' not found in ${sheet.getName()}`);
  return index + 1;
}

/**
 * Gets Control sheet, creating it if needed.
 * @returns {GoogleAppsScript.Spreadsheet.Sheet} Control sheet.
 */
function getControlSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(CONTROL_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(CONTROL_SHEET_NAME);
    sheet.getRange(1, 1, 1, 41).setValues([[
      'GroupID', 'CooperatorID', 'Name', 'Email', 'Phone', 'HomeAddress', 'LoanAmountFigures',
      'LoanAmountWords', 'RepaymentPeriod', 'Guarantor1Name', 'Guarantor1ID', 'Guarantor1Email', 'Guarantor1Phone',
      'Guarantor2Name', 'Guarantor2ID', 'Guarantor2Email', 'Guarantor2Phone', 'ApproverName', 'ApproverID',
      'ApproverEmail', 'ApproverPhone', 'Status', 'ApplicantLink', 'FinanceLink', 'ApplicationStatus', 'Locked',
      'Timestamp', 'Comments', 'BankName', 'AccountName', 'AccountNumber', 'ApplicantBalance', 'ApplicantRating',
      'Guarantor1Balance', 'Guarantor1Rating', 'Guarantor2Balance', 'Guarantor2Rating', 'FinanceOfficerName',
      'FinanceOfficerID', 'FinanceOfficerEmail', 'FinanceOfficerPhone', 'Notified'
    ]]);
  }
  return sheet;
}

/**
 * Gets or creates Archive sheet.
 * @returns {GoogleAppsScript.Spreadsheet.Sheet} Archive sheet.
 */
function getArchiveSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(ARCHIVE_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(ARCHIVE_SHEET_NAME);
    sheet.getRange(1, 1, 1, 41).setValues([getControlSheet().getRange(1, 1, 1, 41).getValues()[0]]);
  }
  return sheet;
}

/**
 * Gets or creates FinanceOfficers sheet.
 * @returns {GoogleAppsScript.Spreadsheet.Sheet} FinanceOfficers sheet.
 */
function getFinanceOfficersSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(FINANCE_OFFICERS_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(FINANCE_OFFICERS_SHEET_NAME);
    sheet.getRange(1, 1, 1, 4).setValues([['Name', 'ID', 'Email', 'Phone']]);
  }
  return sheet;
}

/**
 * Gets or creates Logs sheet.
 * @returns {GoogleAppsScript.Spreadsheet.Sheet} Logs sheet.
 */
function getLogsSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(LOGS_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(LOGS_SHEET_NAME);
    sheet.getRange(1, 1, 1, 3).setValues([['Timestamp', 'Action', 'Details']]);
  }
  return sheet;
}

/**
 * Logs an action to Logs sheet.
 * @param {string} action Action performed.
 * @param {string} details Details of the action.
 */
function logAction(action, details) {
  const sheet = getLogsSheet();
  sheet.appendRow([new Date(), action, details]);
}

/**
 * Generates a unique GroupID (e.g., LC0001).
 * @returns {string} GroupID.
 */
function generateGroupID() {
  const sheet = getControlSheet();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return 'LC0001';
  const lastID = sheet.getRange(lastRow, 1).getValue();
  const num = parseInt(lastID.replace('LC', ''), 10);
  return 'LC' + (num + 1).toString().padStart(4, '0');
}

/**
 * Gets form entry IDs for prefilling.
 * @param {string} formId Form ID.
 * @returns {Object} Entry IDs for form fields.
 */
function getFormEntryIds(formId) {
  try {
    const form = FormApp.openById(formId);
    const items = form.getItems();
    const entryIds = {};
    items.forEach(item => {
      const title = item.getTitle();
      const id = item.asTextItem()?.getId() || item.asListItem()?.getId() || item.asMultipleChoiceItem()?.getId();
      if (id) entryIds[title] = `entry.${id}`;
    });
    return {
      'Loan ID': entryIds['Loan ID'] || '',
      'Role': entryIds['Role'] || '',
      'Email': entryIds['Email'] || ''
    };
  } catch (e) {
    logAction('Error', `Failed to get form entry IDs: ${e.message}`);
    return {};
  }
}

/**
 * Generates a prefilled form link.
 * @param {string} groupID Group ID.
 * @param {string} role Role (Applicant, Guarantor, Finance Officer).
 * @param {string} email Email address.
 * @returns {string} Prefilled form URL.
 */
function generatePrefilledLink(groupID, role, email) {
  const formId = PropertiesService.getScriptProperties().getProperty('LOAN_APPLICATION_FORM_ID');
  if (!formId) {
    logAction('Error', 'Application Form ID not set');
    return '';
  }
  try {
    const formUrl = FormApp.openById(formId).getPublishedUrl().replace('/viewform', '/viewform?');
    const entryIds = getFormEntryIds(formId);
    if (!entryIds['Loan ID'] || !entryIds['Role'] || !entryIds['Email']) {
      logAction('Error', 'Missing form entry IDs');
      return '';
    }
    return `${formUrl}${entryIds['Loan ID']}=${encodeURIComponent(groupID)}&${entryIds['Role']}=${encodeURIComponent(role)}&${entryIds['Email']}=${encodeURIComponent(email)}`;
  } catch (e) {
    logAction('Error', `Failed to generate prefilled link: ${e.message}`);
    return '';
  }
}

/**
 * Builds an HTML table from an object.
 * @param {Object} dataObj Key-value pairs for table.
 * @returns {string} HTML table string.
 */
function buildHtmlTableFromObject(dataObj) {
  let rows = '';
  for (const key in dataObj) {
    if (dataObj.hasOwnProperty(key)) {
      rows += `<tr><td style="border:1px solid #ddd;padding:8px;font-family:Arial,sans-serif;"><strong>${key}</strong></td><td style="border:1px solid #ddd;padding:8px;font-family:Arial,sans-serif;">${dataObj[key] || ''}</td></tr>`;
    }
  }
  return `<table style="border-collapse:collapse;border:2px solid #ddd;font-family:Arial,sans-serif;">${rows}</table>`;
}

/**
 * Sends an email with HTML content.
 * @param {string} to Recipient email.
 * @param {string} subject Email subject.
 * @param {string} bodyHtml HTML email body.
 */
function sendEmail(to, subject, bodyHtml) {
  if (!to || !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(to)) {
    logAction('Error', `Invalid email address: ${to}`);
    return;
  }
  try {
    MailApp.sendEmail({
      to: to,
      subject: subject,
      htmlBody: bodyHtml + EMAIL_REGARDS + EMAIL_FOOTER,
      name: SENDER_NAME
    });
    logAction('Email Sent', `To: ${to}, Subject: ${subject}`);
  } catch (e) {
    logAction('Error', `Failed to send email to ${to}: ${e.message}`);
  }
}

/**
 * Checks if a Finance Officer is assigned for a group.
 * @param {string} groupID Group ID.
 * @returns {boolean} True if assigned.
 */
function checkFinanceOfficerExists(groupID) {
  const sheet = getControlSheet();
  const data = sheet.getDataRange().getValues();
  const cols = {
    groupID: getColumnIndex(sheet, 'GroupID'),
    foName: getColumnIndex(sheet, 'FinanceOfficerName'),
    foId: getColumnIndex(sheet, 'FinanceOfficerID'),
    foEmail: getColumnIndex(sheet, 'FinanceOfficerEmail'),
    foPhone: getColumnIndex(sheet, 'FinanceOfficerPhone')
  };
  return data.some(row => row[cols.groupID - 1] === groupID && row[cols.foName - 1] && row[cols.foId - 1] && row[cols.foEmail - 1] && row[cols.foPhone - 1]);
}

/**
 * Gets Finance Officer details for a group.
 * @param {string} groupID Group ID.
 * @returns {Object|null} Officer details or null.
 */
function getFinanceOfficer(groupID) {
  const sheet = getControlSheet();
  const data = sheet.getDataRange().getValues();
  const cols = {
    groupID: getColumnIndex(sheet, 'GroupID'),
    foName: getColumnIndex(sheet, 'FinanceOfficerName'),
    foId: getColumnIndex(sheet, 'FinanceOfficerID'),
    foEmail: getColumnIndex(sheet, 'FinanceOfficerEmail'),
    foPhone: getColumnIndex(sheet, 'FinanceOfficerPhone')
  };
  const row = data.find(row => row[cols.groupID - 1] === groupID);
  if (!row || !row[cols.foName - 1]) return null;
  return {
    name: row[cols.foName - 1],
    id: row[cols.foId - 1],
    email: row[cols.foEmail - 1],
    phone: row[cols.foPhone - 1]
  };
}

/**
 * Auto-assigns a Finance Officer to a group.
 * @param {string} groupID Group ID.
 */
function autoAssignFinanceOfficer(groupID) {
  const foSheet = getFinanceOfficersSheet();
  const foData = foSheet.getDataRange().getValues().slice(1).filter(row => row[0] && row[1] && row[2] && row[3]);
  if (foData.length === 0) {
    logAction('Error', `No Finance Officers available for GroupID: ${groupID}`);
    return;
  }
  const randomIndex = Math.floor(Math.random() * foData.length);
  const fo = foData[randomIndex];
  const sheet = getControlSheet();
  const data = sheet.getDataRange().getValues();
  const cols = {
    groupID: getColumnIndex(sheet, 'GroupID'),
    foName: getColumnIndex(sheet, 'FinanceOfficerName'),
    foId: getColumnIndex(sheet, 'FinanceOfficerID'),
    foEmail: getColumnIndex(sheet, 'FinanceOfficerEmail'),
    foPhone: getColumnIndex(sheet, 'FinanceOfficerPhone')
  };
  for (let i = 1; i < data.length; i++) {
    if (data[i][cols.groupID - 1] === groupID) {
      sheet.getRange(i + 1, cols.foName, 1, 4).setValues([[fo[0], fo[1], fo[2], fo[3]]]);
      logAction('Auto-Assign', `GroupID: ${groupID}, Officer: ${fo[0]}`);
      break;
    }
  }
}

/**
 * Sets up Form IDs by prompting for form titles.
 */
function setUpFormIds() {
  const ui = SpreadsheetApp.getUi();
  const intentFormTitle = ui.prompt('Enter Intent Form Title').getResponseText();
  const appFormTitle = ui.prompt('Enter Application Form Title').getResponseText();
  try {
    const forms = DriveApp.getFilesByType(MimeType.GOOGLE_FORMS);
    let intentFormId, appFormId;
    while (forms.hasNext()) {
      const form = forms.next();
      const formApp = FormApp.openById(form.getId());
      if (formApp.getTitle() === intentFormTitle) intentFormId = form.getId();
      if (formApp.getTitle() === appFormTitle) appFormId = form.getId();
    }
    if (!intentFormId || !appFormId) {
      ui.alert('Error: One or both forms not found. Set IDs manually in Script Properties.');
      logAction('Error', `Forms not found: Intent=${intentFormTitle}, Application=${appFormTitle}`);
      return;
    }
    PropertiesService.getScriptProperties().setProperties({
      'LOAN_INTENT_FORM_ID': intentFormId,
      'LOAN_APPLICATION_FORM_ID': appFormId
    });
    ui.alert('Success: Form IDs set.');
    logAction('Setup', `Form IDs set: Intent=${intentFormId}, Application=${appFormId}`);
  } catch (e) {
    ui.alert('Error: Failed to set Form IDs. Check logs.');
    logAction('Error', `Setting Form IDs: ${e.message}`);
  }
}

/**
 * Updates Application Form dropdowns with participant data.
 */
function syncParticipantDetails() {
  const intentSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(LOAN_INTENT_SHEET_NAME);
  const controlSheet = getControlSheet();
  if (!intentSheet) {
    logAction('Error', 'Intent sheet not found');
    return;
  }
  const intentData = intentSheet.getDataRange().getValues().slice(1);
  const controlData = controlSheet.getDataRange().getValues().slice(1);
  const formId = PropertiesService.getScriptProperties().getProperty('LOAN_APPLICATION_FORM_ID');
  if (!formId) {
    logAction('Error', 'Application Form ID not set');
    return;
  }
  try {
    const form = FormApp.openById(formId);
    const items = form.getItems();
    const fieldMap = {
      'Cooperator ID': intentData.filter(row => row[6] === ROLE_APPLICANT).map(row => row[2]),
      'Name': intentData.filter(row => row[6] === ROLE_APPLICANT).map(row => row[3]),
      'Phone': intentData.filter(row => row[6] === ROLE_APPLICANT).map(row => row[4]),
      'Email': intentData.filter(row => row[6] === ROLE_APPLICANT).map(row => row[5]),
      'Guarantor 1 Name': intentData.filter(row => row[6] === ROLE_GUARANTOR).map(row => row[3]),
      'Guarantor 1 Cooperator ID': intentData.filter(row => row[6] === ROLE_GUARANTOR).map(row => row[2]),
      'Guarantor 1 Email': intentData.filter(row => row[6] === ROLE_GUARANTOR).map(row => row[5]),
      'Guarantor 1 Phone': intentData.filter(row => row[6] === ROLE_GUARANTOR).map(row => row[4]),
      'Guarantor 2 Name': intentData.filter(row => row[6] === ROLE_GUARANTOR).map(row => row[3]),
      'Guarantor 2 Cooperator ID': intentData.filter(row => row[6] === ROLE_GUARANTOR).map(row => row[2]),
      'Guarantor 2 Email': intentData.filter(row => row[6] === ROLE_GUARANTOR).map(row => row[5]),
      'Guarantor 2 Phone': intentData.filter(row => row[6] === ROLE_GUARANTOR).map(row => row[4]),
      'Approver ID': controlData.map(row => row[getColumnIndex(controlSheet, 'FinanceOfficerID') - 1]).filter(id => id),
      'Approver Name': controlData.map(row => row[getColumnIndex(controlSheet, 'FinanceOfficerName') - 1]).filter(name => name),
      'Approver Email': controlData.map(row => row[getColumnIndex(controlSheet, 'FinanceOfficerEmail') - 1]).filter(email => email),
      'Approver Phone': controlData.map(row => row[getColumnIndex(controlSheet, 'FinanceOfficerPhone') - 1]).filter(phone => phone)
    };
    items.forEach(item => {
      const title = item.getTitle();
      if (fieldMap[title] && item.asListItem()) {
        item.asListItem().setChoiceValues([...new Set(fieldMap[title])]);
      }
    });
    logAction('Sync', 'Updated Application Form dropdowns');
  } catch (e) {
    logAction('Error', `Failed to sync dropdowns: ${e.message}`);
  }
}

/**
 * Handles Intent Form submissions.
 * @param {GoogleAppsScript.Forms.FormSubmitEvent} e Form submission event.
 */
function onIntentFormSubmit(e) {
  const intentSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(LOAN_INTENT_SHEET_NAME);
  if (!intentSheet) {
    logAction('Error', 'Intent sheet not found');
    return;
  }
  if (!e) {
    logAction('Error', 'Form submission event is undefined');
    return;
  }
  Logger.log(JSON.stringify(e));
  if (!e.namedValues) {
    logAction('Error', `No namedValues in event: ${JSON.stringify(e)}`);
    return;
  }
  const responses = e.namedValues;
  const applicant = {
    id: responses['Your Cooperator ID']?.[0] || '',
    name: responses['Your Name']?.[0] || '',
    phone: responses['Your Phone']?.[0] || '',
    email: responses['Your Email']?.[0] || ''
  };
  const guarantor1 = {
    id: responses['Guarantor 1 Cooperator ID']?.[0] || '',
    name: responses['Guarantor 1 Name']?.[0] || '',
    phone: responses['Guarantor 1 Phone']?.[0] || '',
    email: responses['Guarantor 1 Email']?.[0] || ''
  };
  const guarantor2 = {
    id: responses['Guarantor 2 Cooperator ID']?.[0] || '',
    name: responses['Guarantor 2 Name']?.[0] || '',
    phone: responses['Guarantor 2 Phone']?.[0] || '',
    email: responses['Guarantor 2 Email']?.[0] || ''
  };

  if (!applicant.id || !applicant.email || !guarantor1.id || !guarantor1.email || !guarantor2.id || !guarantor2.email) {
    logAction('Error', 'Incomplete Intent Form submission');
    return;
  }

  const groupID = generateGroupID();
  const timestamp = new Date();
  const rows = [
    [timestamp, groupID, applicant.id, applicant.name, applicant.phone, applicant.email, ROLE_APPLICANT],
    [timestamp, groupID, guarantor1.id, guarantor1.name, guarantor1.phone, guarantor1.email, ROLE_GUARANTOR],
    [timestamp, groupID, guarantor2.id, guarantor2.name, guarantor2.phone, guarantor2.email, ROLE_GUARANTOR]
  ];

  const existingData = intentSheet.getDataRange().getValues().slice(1);
  for (const row of rows) {
    if (existingData.some(r => r[1] === row[1] && r[2] === row[2])) {
      logAction('Error', `Duplicate entry: GroupID=${row[1]}, CooperatorID=${row[2]}`);
      continue;
    }
    intentSheet.appendRow(row);
  }

  const controlSheet = getControlSheet();
  controlSheet.appendRow([
    groupID, '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', 'PendingFinanceOfficer', 'FALSE', timestamp, '',
    '', '', '', '', '', '', '', '', '', '', '', '', 'FALSE'
  ]);

  autoAssignFinanceOfficer(groupID);

  const prefillLink = generatePrefilledLink(groupID, ROLE_APPLICANT, applicant.email);
  const intentData = {
    'GroupID': groupID,
    'Applicant Name': applicant.name,
    'Applicant Cooperator ID': applicant.id,
    'Applicant Email': applicant.email,
    'Applicant Phone': applicant.phone,
    'Guarantor 1 Name': guarantor1.name,
    'Guarantor 1 Cooperator ID': guarantor1.id,
    'Guarantor 1 Email': guarantor1.email,
    'Guarantor 1 Phone': guarantor1.phone,
    'Guarantor 2 Name': guarantor2.name,
    'Guarantor 2 Cooperator ID': guarantor2.id,
    'Guarantor 2 Email': guarantor2.email,
    'Guarantor 2 Phone': guarantor2.phone
  };
  const intentTable = buildHtmlTableFromObject(intentData);
  sendEmail(
    applicant.email,
    `Loan Intent Submitted - ${groupID}`,
    `<p>Dear ${applicant.name},</p><p>Your loan intent (GroupID: ${groupID}) has been received. Please wait for Finance Officer assignment before proceeding.</p>${intentTable}<p>Application Form: <a href="${prefillLink}">Click here</a></p>`
  );
  logAction('Intent', `Submitted: GroupID=${groupID}`);
}

/**
 * Handles Application Form submissions.
 * @param {GoogleAppsScript.Forms.FormSubmitEvent} e Form submission event.
 */
function onApplicationFormSubmit(e) {
  const sheet = getControlSheet();
  const responses = e.namedValues;
  const role = responses['Role']?.[0] || '';
  let groupID = responses['Loan ID']?.[0] || '';
  const data = {
    applicantId: responses['Cooperator ID']?.[0] || '',
    applicantName: responses['Name']?.[0] || '',
    applicantEmail: responses['Email']?.[0] || '',
    applicantPhone: responses['Phone']?.[0] || '',
    homeAddress: responses['Home Address']?.[0] || '',
    loanAmountFigures: responses['Loan Amount (Figures)']?.[0] || '',
    loanAmountWords: responses['Loan Amount (Words)']?.[0] || '',
    repaymentPeriod: responses['Repayment Period']?.[0] || '',
    bankName: responses['Bank Name']?.[0] || '',
    accountName: responses['Account Name']?.[0] || '',
    accountNumber: responses['Account Number']?.[0] || '',
    guarantor1Name: responses['Guarantor 1 Name']?.[0] || '',
    guarantor1Id: responses['Guarantor 1 Cooperator ID']?.[0] || '',
    guarantor1Email: responses['Guarantor 1 Email']?.[0] || '',
    guarantor1Phone: responses['Guarantor 1 Phone']?.[0] || '',
    guarantor2Name: responses['Guarantor 2 Name']?.[0] || '',
    guarantor2Id: responses['Guarantor 2 Cooperator ID']?.[0] || '',
    guarantor2Email: responses['Guarantor 2 Email']?.[0] || '',
    guarantor2Phone: responses['Guarantor 2 Phone']?.[0] || '',
    approverName: responses['Approver Name']?.[0] || '',
    approverId: responses['Approver ID']?.[0] || '',
    approverEmail: responses['Approver Email']?.[0] || '',
    approverPhone: responses['Approver Phone']?.[0] || '',
    status: responses['Status']?.[0] || '',
    applicantBalance: responses['Applicant Balance']?.[0] || '',
    applicantRating: responses['Applicant Rating']?.[0] || '',
    guarantor1Balance: responses['Guarantor 1 Balance']?.[0] || '',
    guarantor1Rating: responses['Guarantor 1 Rating']?.[0] || '',
    guarantor2Balance: responses['Guarantor 2 Balance']?.[0] || '',
    guarantor2Rating: responses['Guarantor 2 Rating']?.[0] || '',
    comments: responses['Comments']?.[0] || ''
  };

  if (!role || (role === ROLE_APPLICANT && !data.applicantId) || (role === ROLE_GUARANTOR && (!data.guarantor1Id || !data.guarantor2Id)) || (role === ROLE_FINANCE && !data.approverId)) {
    logAction('Error', `Incomplete Application Form data: Role=${role}, GroupID=${groupID}`);
    return;
  }

  if (role === ROLE_APPLICANT && !checkFinanceOfficerExists(groupID)) {
    sendEmail(data.applicantEmail, `Submission Blocked - ${groupID}`, `<p>Dear ${data.applicantName},</p><p>No Finance Officer assigned for GroupID: ${groupID}. Please contact the admin.</p>`);
    logAction('Error', `Blocked: No Finance Officer for GroupID: ${groupID}`);
    return;
  }

  const financeOfficer = getFinanceOfficer(groupID) || { name: '', id: '', email: '', phone: '' };
  const controlData = sheet.getDataRange().getValues();
  const cols = {
    groupID: getColumnIndex(sheet, 'GroupID'),
    cooperatorID: getColumnIndex(sheet, 'CooperatorID'),
    status: getColumnIndex(sheet, 'ApplicationStatus'),
    locked: getColumnIndex(sheet, 'Locked'),
    notified: getColumnIndex(sheet, 'Notified')
  };
  let rowIndex = controlData.findIndex(row => row[cols.groupID - 1] === groupID) + 1;

  if (role === ROLE_APPLICANT) {
    if (controlData.some(row => row[cols.cooperatorID - 1] === data.applicantId && !['FinanceReviewed', 'Expired'].includes(row[cols.status - 1]))) {
      sendEmail(data.applicantEmail, `Submission Blocked - Active Loan`, `<p>Dear ${data.applicantName},</p><p>You have an active loan application. Please contact the admin.</p>`);
      logAction('Error', `Blocked: Active loan for CooperatorID: ${data.applicantId}`);
      return;
    }
    if (!groupID || rowIndex === 0) {
      groupID = generateGroupID();
      rowIndex = sheet.getLastRow() + 1;
    } else if (sheet.getRange(rowIndex, cols.locked).getValue() === 'TRUE') {
      sendEmail(data.applicantEmail, `Application Locked - ${groupID}`, `<p>Dear ${data.applicantName},</p><p>Application ${groupID} is locked. Please contact the admin.</p>`);
      logAction('Error', `Blocked: Application locked for GroupID: ${groupID}`);
      return;
    }
    const newRow = Array(41).fill('');
    Object.keys(data).forEach(key => {
      const colName = key.replace(/([A-Z])/g, ' $1').replace(/^./, str => str.toUpperCase());
      try {
        const col = getColumnIndex(sheet, colName);
        newRow[col - 1] = data[key];
      } catch (e) {}
    });
    newRow[cols.groupID - 1] = groupID;
    newRow[cols.cooperatorID - 1] = data.applicantId;
    newRow[cols.status - 1] = 'ApplicantSubmitted';
    newRow[cols.locked - 1] = 'FALSE';
    newRow[cols.notified - 1] = 'TRUE';
    newRow[getColumnIndex(sheet, 'Timestamp') - 1] = new Date();
    newRow[getColumnIndex(sheet, 'FinanceOfficerName') - 1] = financeOfficer.name;
    newRow[getColumnIndex(sheet, 'FinanceOfficerID') - 1] = financeOfficer.id;
    newRow[getColumnIndex(sheet, 'FinanceOfficerEmail') - 1] = financeOfficer.email;
    newRow[getColumnIndex(sheet, 'FinanceOfficerPhone') - 1] = financeOfficer.phone;
    sheet.getRange(rowIndex, 1, 1, 41).setValues([newRow]);
  } else if (role === ROLE_GUARANTOR) {
    if (!groupID || rowIndex === 0) {
      logAction('Error', `No GroupID for Role: ${role}`);
      return;
    }
    if (sheet.getRange(rowIndex, cols.locked).getValue() === 'TRUE') {
      sendEmail(data.guarantor1Email || data.guarantor2Email, `Application Locked - ${groupID}`, `<p>Dear ${data.guarantor1Name || data.guarantor2Name},</p><p>Application ${groupID} is locked.</p>`);
      logAction('Error', `Blocked: ${role} submission for GroupID: ${groupID}`);
      return;
    }
    const guarantorNum = data.guarantor1Email === responses['Email']?.[0] ? '1' : '2';
    sheet.getRange(rowIndex, getColumnIndex(sheet, `Guarantor${guarantorNum}Name`), 1, 4).setValues([[
      data[`guarantor${guarantorNum}Name`], data[`guarantor${guarantorNum}Id`], 
      data[`guarantor${guarantorNum}Email`], data[`guarantor${guarantorNum}Phone`]
    ]]);
    sheet.getRange(rowIndex, cols.notified).setValue('TRUE');
  } else if (role === ROLE_FINANCE) {
    if (!groupID || rowIndex === 0) {
      logAction('Error', `No GroupID for Role: ${role}`);
      return;
    }
    if (sheet.getRange(rowIndex, cols.locked).getValue() === 'TRUE') {
      logAction('Error', `Blocked: ${role} submission for GroupID: ${groupID}`);
      return;
    }
    const fields = ['ApproverName', 'ApproverID', 'ApproverEmail', 'ApproverPhone', 'Status', 'Comments', 'BankName', 
                    'AccountName', 'AccountNumber', 'ApplicantBalance', 'ApplicantRating', 'Guarantor1Balance', 
                    'Guarantor1Rating', 'Guarantor2Balance', 'Guarantor2Rating'];
    const values = fields.map(field => data[field.toLowerCase()] || '');
    sheet.getRange(rowIndex, getColumnIndex(sheet, 'ApproverName'), 1, fields.length).setValues([values]);
    sheet.getRange(rowIndex, cols.status).setValue('FinanceReviewed');
    sheet.getRange(rowIndex, cols.locked).setValue('TRUE');
    sheet.getRange(rowIndex, cols.notified).setValue('TRUE');
  } else {
    logAction('Error', `Invalid role: ${role}`);
    return;
  }

  const links = {
    applicant: generatePrefilledLink(groupID, ROLE_APPLICANT, data.applicantEmail),
    guarantor1: generatePrefilledLink(groupID, ROLE_GUARANTOR, data.guarantor1Email),
    guarantor2: generatePrefilledLink(groupID, ROLE_GUARANTOR, data.guarantor2Email),
    finance: generatePrefilledLink(groupID, ROLE_FINANCE, financeOfficer.email)
  };
  sheet.getRange(rowIndex, getColumnIndex(sheet, 'ApplicantLink'), 1, 2).setValues([[links.applicant, links.finance]]);

  const emailData = {
    'Loan ID': groupID,
    'Applicant Name': data.applicantName,
    'Applicant Cooperator ID': data.applicantId,
    'Applicant Email': data.applicantEmail,
    'Applicant Phone': data.applicantPhone,
    'Home Address': data.homeAddress,
    'Loan Amount (Figures)': data.loanAmountFigures,
    'Loan Amount (Words)': data.loanAmountWords,
    'Repayment Period': data.repaymentPeriod,
    'Bank Name': data.bankName,
    'Account Name': data.accountName,
    'Account Number': data.accountNumber,
    'Guarantor 1 Name': data.guarantor1Name,
    'Guarantor 1 Cooperator ID': data.guarantor1Id,
    'Guarantor 1 Email': data.guarantor1Email,
    'Guarantor 1 Phone': data.guarantor1Phone,
    'Guarantor 2 Name': data.guarantor2Name,
    'Guarantor 2 Cooperator ID': data.guarantor2Id,
    'Guarantor 2 Email': data.guarantor2Email,
    'Guarantor 2 Phone': data.guarantor2Phone,
    'Status': data.status || 'Pending'
  };
  const table = buildHtmlTableFromObject(emailData);

  switch (role) {
    case ROLE_APPLICANT:
      sendEmail(data.applicantEmail, `Application Submitted - ${groupID}`, 
        `<p>Dear ${data.applicantName},</p><p>Your application (GroupID: ${groupID}) has been submitted.</p>${table}<p>Edit: <a href="${links.applicant}">Click here</a></p>`);
      sendEmail(data.guarantor1Email, `Action Required - ${groupID}`, 
        `<p>Dear ${data.guarantor1Name},</p><p>Please submit details for ${data.applicantName}'s loan (GroupID: ${groupID}).</p>${table}<p><a href="${links.guarantor1}">Submit</a></p>`);
      sendEmail(data.guarantor2Email, `Action Required - ${groupID}`, 
        `<p>Dear ${data.guarantor2Name},</p><p>Please submit details for ${data.applicantName}'s loan (GroupID: ${groupID}).</p>${table}<p><a href="${links.guarantor2}">Submit</a></p>`);
      if (financeOfficer.email) {
        sendEmail(financeOfficer.email, `Review Application - ${groupID}`, 
          `<p>Dear ${financeOfficer.name},</p><p>Please review ${data.applicantName}'s loan application (GroupID: ${groupID}).</p>${table}<p><a href="${links.finance}">Review</a></p>`);
      }
      break;
    case ROLE_GUARANTOR:
      const guarantorNum = data.guarantor1Email === responses['Email']?.[0] ? '1' : '2';
      const gName = guarantorNum === '1' ? data.guarantor1Name : data.guarantor2Name;
      const gEmail = guarantorNum === '1' ? data.guarantor1Email : data.guarantor2Email;
      sendEmail(gEmail, `Details Submitted - ${groupID}`, 
        `<p>Dear ${gName},</p><p>Your details for ${data.applicantName}'s loan (GroupID: ${groupID}) have been submitted.</p>${table}`);
      [data.applicantEmail, financeOfficer.email].filter(email => email).forEach(email => {
        sendEmail(email, `Guarantor ${guarantorNum} Submitted - ${groupID}`, 
          `<p>Dear ${email === data.applicantEmail ? data.applicantName : financeOfficer.name},</p><p>${gName} has submitted details for GroupID: ${groupID}.</p>${table}`);
      });
      break;
    case ROLE_FINANCE:
      [data.applicantEmail, data.guarantor1Email, data.guarantor2Email].filter(email => email).forEach(email => {
        sendEmail(email, `Application Reviewed - ${groupID}`, 
          `<p>Dear ${email === data.applicantEmail ? data.applicantName : email === data.guarantor1Email ? data.guarantor1Name : data.guarantor2Name},</p><p>Loan application ${groupID} has been reviewed. Status: ${data.status}.</p>${table}`);
      });
      archiveApplication(groupID);
      break;
  }
  logAction('Application', `Submitted: GroupID=${groupID}, Role=${role}`);
}

/**
 * Sends daily reminders for pending actions.
 */
function sendDailyReminders() {
  const sheet = getControlSheet();
  const data = sheet.getDataRange().getValues();
  const now = new Date();
  let emailCount = 0;
  const cols = {
    groupID: getColumnIndex(sheet, 'GroupID'),
    name: getColumnIndex(sheet, 'Name'),
    email: getColumnIndex(sheet, 'Email'),
    guarantor1Name: getColumnIndex(sheet, 'Guarantor1Name'),
    guarantor1Email: getColumnIndex(sheet, 'Guarantor1Email'),
    guarantor2Name: getColumnIndex(sheet, 'Guarantor2Name'),
    guarantor2Email: getColumnIndex(sheet, 'Guarantor2Email'),
    financeOfficerName: getColumnIndex(sheet, 'FinanceOfficerName'),
    financeOfficerEmail: getColumnIndex(sheet, 'FinanceOfficerEmail'),
    status: getColumnIndex(sheet, 'ApplicationStatus'),
    financeLink: getColumnIndex(sheet, 'FinanceLink'),
    timestamp: getColumnIndex(sheet, 'Timestamp')
  };

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const groupID = row[cols.groupID - 1];
    if (!checkFinanceOfficerExists(groupID) || row[cols.status - 1] !== 'ApplicantSubmitted') continue;
    const timestamp = new Date(row[cols.timestamp - 1]);
    const daysSince = Math.floor((now - timestamp) / (1000 * 60 * 60 * 24));
    if (daysSince >= REMINDER_DAYS_LIMIT) {
      sheet.getRange(i + 1, cols.status).setValue('Expired');
      sheet.getRange(i + 1, getColumnIndex(sheet, 'Locked')).setValue('TRUE');
      sendEmail(row[cols.email - 1], `Application Expired - ${groupID}`, 
        `<p>Dear ${row[cols.name - 1]},</p><p>Your application (GroupID: ${groupID}) has expired due to inactivity.</p>`);
      archiveApplication(groupID);
      logAction('Expired', `GroupID=${groupID}`);
      continue;
    }
    const reminderData = {
      'Loan ID': groupID,
      'Applicant Name': row[cols.name - 1],
      'Status': row[cols.status - 1]
    };
    const table = buildHtmlTableFromObject(reminderData);
    const emails = [
      { email: row[cols.guarantor1Email - 1], name: row[cols.guarantor1Name - 1], role: ROLE_GUARANTOR },
      { email: row[cols.guarantor2Email - 1], name: row[cols.guarantor2Name - 1], role: ROLE_GUARANTOR },
      { email: row[cols.financeOfficerEmail - 1], name: row[cols.financeOfficerName - 1], role: ROLE_FINANCE, linkCol: cols.financeLink }
    ];
    emails.forEach(({ email, name, role, linkCol }) => {
      if (email && emailCount < 50) {
        const link = linkCol ? row[linkCol - 1] : generatePrefilledLink(groupID, role, email);
        if (link) {
          sendEmail(email, `Reminder - ${groupID}`, 
            `<p>Dear ${name},</p><p>${role === ROLE_FINANCE ? 'Please review' : 'Please submit details for'} loan application ${groupID}.</p>${table}<p><a href="${link}">${role === ROLE_FINANCE ? 'Review' : 'Submit'}</a></p>`);
          emailCount++;
        }
      }
    });
  }
  logAction('Reminders', `Sent ${emailCount} reminders`);
}

/**
 * Notifies applicants of new Finance Officer assignments.
 */
function notifyNewFinanceOfficerAssignments() {
  const sheet = getControlSheet();
  const data = sheet.getDataRange().getValues();
  const cols = {
    groupID: getColumnIndex(sheet, 'GroupID'),
    name: getColumnIndex(sheet, 'Name'),
    email: getColumnIndex(sheet, 'Email'),
    financeOfficerName: getColumnIndex(sheet, 'FinanceOfficerName'),
    financeOfficerEmail: getColumnIndex(sheet, 'FinanceOfficerEmail'),
    status: getColumnIndex(sheet, 'ApplicationStatus'),
    notified: getColumnIndex(sheet, 'Notified')
  };
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const groupID = row[cols.groupID - 1];
    if (row[cols.notified - 1] === 'TRUE' || row[cols.status - 1] !== 'PendingFinanceOfficer' || !row[cols.financeOfficerEmail - 1]) continue;
    const link = generatePrefilledLink(groupID, ROLE_APPLICANT, row[cols.email - 1]);
    const table = buildHtmlTableFromObject({
      'Loan ID': groupID,
      'Applicant Name': row[cols.name - 1],
      'Finance Officer': row[cols.financeOfficerName - 1]
    });
    sendEmail(
      row[cols.email - 1],
      `Finance Officer Assigned - ${groupID}`,
      `<p>Dear ${row[cols.name - 1]},</p><p>A Finance Officer has been assigned to your loan application (GroupID: ${groupID}). You may now proceed.</p>${table}<p><a href="${link}">Submit Application</a></p>`
    );
    sheet.getRange(i + 1, cols.notified).setValue('TRUE');
    logAction('Notification', `Notified applicant for GroupID: ${groupID}`);
  }
}

/**
 * Archives a completed or expired application.
 * @param {string} groupID Group ID.
 */
function archiveApplication(groupID) {
  const controlSheet = getControlSheet();
  const archiveSheet = getArchiveSheet();
  const data = controlSheet.getDataRange().getValues();
  const cols = { groupID: getColumnIndex(controlSheet, 'GroupID') };
  for (let i = 1; i < data.length; i++) {
    if (data[i][cols.groupID - 1] === groupID) {
      archiveSheet.appendRow(data[i]);
      controlSheet.deleteRow(i + 1);
      logAction('Archived', `GroupID: ${groupID}`);
      break;
    }
  }
}

/**
 * Resets an application to PendingFinanceOfficer status.
 */
function manualReset() {
  const ui = SpreadsheetApp.getUi();
  const groupID = ui.prompt('Enter GroupID to reset').getResponseText();
  const sheet = getControlSheet();
  const data = sheet.getDataRange().getValues();
  const cols = {
    groupID: getColumnIndex(sheet, 'GroupID'),
    status: getColumnIndex(sheet, 'ApplicationStatus'),
    locked: getColumnIndex(sheet, 'Locked'),
    notified: getColumnIndex(sheet, 'Notified'),
    timestamp: getColumnIndex(sheet, 'Timestamp')
  };
  const rowIndex = data.findIndex(row => row[cols.groupID - 1] === groupID) + 1;
  if (rowIndex === 0) {
    ui.alert(`Error: GroupID ${groupID} not found`);
    logAction('Error', `Reset failed for GroupID: ${groupID}`);
    return;
  }
  sheet.getRange(rowIndex, cols.status, 1, 3).setValues([['PendingFinanceOfficer', 'FALSE', 'FALSE']]);
  sheet.getRange(rowIndex, cols.timestamp).setValue(new Date());
  ui.alert(`Success: GroupID ${groupID} reset to PendingFinanceOfficer`);
  logAction('Reset', `GroupID: ${groupID}`);
}

/**
 * Creates a custom menu in the spreadsheet.
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('CSULMCS Loan System')
    .addItem('Set Up Form IDs', 'runSetUpFormIds')
    .addItem('Set Up Triggers', 'setupTriggers')
    .addItem('Notify New Assignments', 'notifyNewFinanceOfficerAssignments')
    .addItem('Reset Application', 'manualReset')
    .addToUi();
}

/**
 * Sets up triggers for the system.
 */
function setupTriggers() {
  const triggers = [
    { name: 'sendDailyReminders', type: 'timeBased', everyDays: 1, atHour: 8 },
    { name: 'onIntentFormSubmit', type: 'formSubmit', formIdKey: 'LOAN_INTENT_FORM_ID' },
    { name: 'onApplicationFormSubmit', type: 'formSubmit', formIdKey: 'LOAN_APPLICATION_FORM_ID' },
    { name: 'syncParticipantDetails', type: 'formSubmit', formIdKey: 'LOAN_APPLICATION_FORM_ID' }
  ];
  triggers.forEach(trigger => {
    if (!ScriptApp.getProjectTriggers().some(t => t.getHandlerFunction() === trigger.name)) {
      if (trigger.type === 'timeBased') {
        ScriptApp.newTrigger(trigger.name)
          .timeBased()
          .everyDays(trigger.everyDays)
          .atHour(trigger.atHour)
          .create();
        logAction('Success', `Created time-based trigger: ${trigger.name}`);
      } else if (trigger.type === 'formSubmit') {
        const formId = PropertiesService.getScriptProperties().getProperty(trigger.formIdKey);
        if (formId) {
          ScriptApp.newTrigger(trigger.name)
            .forForm(formId)
            .onFormSubmit()
            .create();
          logAction('Success', `Created form trigger: ${trigger.name}`);
        } else {
          logAction('Error', `Form ID not set for trigger: ${trigger.name}`);
        }
      }
    }
  });
  logAction('Success', 'All triggers setup completed');
}

/**
 * Wrapper to run setUpFormIds from Sheet context.
 */
function runSetUpFormIds() {
  setUpFormIds();
}
