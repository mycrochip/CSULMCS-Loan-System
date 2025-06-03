/**
 * CSULMCS Loan System Script
 * This is like a magic toy box that runs our loan system!
 * It handles sign-ups (Intent Form), loan applications (Application Form),
 * sends emails, and tracks everything in a Google Sheet.
 * Features:
 * - Applicants sign up with two guarantors in one Intent Form submission.
 * - One email to applicant with GroupID after Intent Form submission.
 * - Admins assign Finance Officers per loan group in the Control tab.
 * - Applications pend until a Finance Officer is assigned (e.g., waiting for funds).
 * - 7-day guarantor reminder countdown starts only after Finance Officer assignment.
 * - Manual trigger to notify newly assigned groups via custom menu.
 * - Notifications sent only to groups starting application flow after officer assignment.
 * - Intent tab: One row per participant (applicant + guarantors), unique by GroupID + CooperatorID.
 * - Control tab: Tracks loan details, including Finance Officer per GroupID.
 * - Tabs: Intent, Control, Archive.
 * - New applications can be submitted with new GroupIDs without unlocking existing ones.
 */

// --- SETTINGS (Like picking colors for our toy) ---
const CONTROL_SHEET_NAME = 'Control'; // Name of the Control tab
const ARCHIVE_SHEET_NAME = 'Archive'; // Name of the Archive tab
const LOAN_APPLICATION_FORM_ID = 'your_loan_application_form_id'; // Replace with Application Form ID
const LOAN_INTENT_FORM_ID = 'your_loan_intent_form_id'; // Replace with Intent Form ID
const LOAN_INTENT_SHEET_NAME = 'Intent'; // Name of the Intent tab
const REMINDER_DAYS_LIMIT = 7; // Guarantors have 7 days to respond
const ROLE_APPLICANT = 'Applicant'; // Role for the person asking for a loan
const ROLE_FINANCE = 'Finance Officer'; // Role for the checker
const ROLE_GUARANTOR1 = 'Guarantor1'; // Role for first friend
const ROLE_GUARANTOR2 = 'Guarantor2'; // Role for second friend
const SENDER_NAME = 'CSULMCS Finance Team'; // Name on emails
const EMAIL_FOOTER = `
  <hr style="border: 1px solid #ccc; margin: 20px 0;">
  <p style="font-size: 12px; color: #666; text-align: center;">
    This email was sent by the <strong>CSULMCS Loan Management System</strong>.<br>
    © 2025 CSULMCS. All rights reserved.
  </p>`; // Bottom of every email
const EMAIL_REGARDS = `<p>Regards,<br>CSULMCS Finance Team</p>`; // Email signature

// --- HELPER FUNCTIONS (Like little tools in our toy box) ---

/**
 * Gets the Control tab, like opening the loan tracking page.
 * @returns {GoogleAppsScript.Spreadsheet.Sheet} The Control sheet.
 */
function getControlSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet(); // Open our notebook
  const sheet = ss.getSheetByName(CONTROL_SHEET_NAME); // Find Control page
  if (!sheet) throw new Error('Control sheet not found'); // Oops, page missing!
  return sheet;
}

/**
 * Gets or creates the Archive tab, like a box for old toys.
 * @returns {GoogleAppsScript.Spreadsheet.Sheet} The Archive sheet.
 */
function getArchiveSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet(); // Open notebook
  let sheet = ss.getSheetByName(ARCHIVE_SHEET_NAME); // Look for Archive page
  if (!sheet) { // If not found, make a new one
    sheet = ss.insertSheet(ARCHIVE_SHEET_NAME);
    // Add labels, like writing a list
    sheet.getRange(1, 1, 1, 40).setValues([[
      'GroupID', 'CooperatorID', 'Name', 'Email', 'Phone', 'HomeAddress', 'LoanAmountFigures',
      'LoanAmountWords', 'RepaymentPeriod', 'Guarantor1Name', 'Guarantor1ID', 'Guarantor1Email', 'Guarantor1Phone',
      'Guarantor2Name', 'Guarantor2ID', 'Guarantor2Email', 'Guarantor2Phone', 'ApproverName',
      'ApproverID', 'ApproverEmail', 'ApproverPhone', 'Status', 'ApplicantLink', 'FinanceLink', 'ApplicationStatus',
      'Locked', 'Timestamp', 'Comments', 'BankName', 'AccountName', 'AccountNumber', 'ApplicantBalance',
      'ApplicantRating', 'Guarantor1Balance', 'Guarantor1Rating', 'Guarantor2Balance', 'Guarantor2Rating',
      'FinanceOfficerName', 'FinanceOfficerID', 'FinanceOfficerEmail', 'FinanceOfficerPhone'
    ]]);
  }
  return sheet;
}

/**
 * Makes a new GroupID, like giving a new toy a number (LC0001, LC0002, etc.).
 * @returns {string} A new GroupID.
 */
function generateGroupID() {
  const sheet = getControlSheet(); // Open Control page
  const lastRow = sheet.getLastRow(); // Find the last toy
  if (lastRow < 2) return 'LC0001'; // If no toys, start with LC0001
  const lastID = sheet.getRange(lastRow, 1).getValue(); // Get last toy’s number
  const num = parseInt(lastID.replace(/[^\d]/g, ''), 10); // Get the number part
  return 'LC' + (num + 1).toString().padStart(4, '0'); // Add 1, make it LC0002
}

/**
 * Gets special codes for form fields, like finding toy labels.
 * @param {string} formId The form’s ID.
 * @returns {Object} Codes for Loan ID, Role, and Email fields.
 */
function getFormEntryIDs(formId) {
  try {
    const form = FormApp.openById(formId); // Open the form
    const items = form.getItems(); // Get all questions
    const entryIDs = {};
    items.forEach(item => {
      const title = item.getTitle(); // Question name
      const id = item.asTextItem()?.getId() || item.asMultipleChoiceItem()?.getId();
      if (id) entryIDs[title] = `entry.${id}`; // Save code like entry.123456
    });
    return {
      'Loan ID': entryIDs['Loan ID'] || '',
      'Role': entryIDs['Role'] || '',
      'Email': entryIDs['Email'] || ''
    };
  } catch (e) {
    Logger.log(`Error getting form codes: ${e.message}`); // Oops, something broke
    return {};
  }
}

/**
 * Makes a special link to fill the form, like a magic key.
 * @param {string} groupID The loan code (e.g., LC0001).
 * @param {string} role Who’s filling it (Applicant, Guarantor1, etc.).
 * @param {string} email Their email.
 * @returns {string} A web link.
 */
function generatePrefilledLink(groupID, role, email) {
  const formUrl = FormApp.openById(LOAN_APPLICATION_FORM_ID).getPublishedUrl().replace('/viewform', '/viewform?');
  const entryIDs = getFormEntryIDs(LOAN_APPLICATION_FORM_ID); // Get form codes
  if (!entryIDs['Loan ID'] || !entryIDs['Role'] || !entryIDs['Email']) {
    Logger.log('Missing form codes'); // Oops, no codes
    return '';
  }
  // Make the link with the right info
  return `${formUrl}${entryIDs['Loan ID']}=${encodeURIComponent(groupID)}&${entryIDs['Role']}=${encodeURIComponent(role)}&${entryIDs['Email']}=${encodeURIComponent(email)}`;
}

/**
 * Turns info into a pretty table, like drawing a neat chart.
 * @param {Object} dataObj The info to show.
 * @returns {string} HTML code for a table.
 */
function buildHtmlTableFromObject(dataObj) {
  let rows = '';
  for (const key in dataObj) {
    if (dataObj.hasOwnProperty(key)) {
      rows += `<tr><td style="border:1px solid #ddd;padding:8px;"><strong>${key}</strong></td><td style="border:1px solid #ddd;padding:8px;">${dataObj[key] || ''}</td></tr>`;
    }
  }
  return `<table style="border-collapse:collapse;border:2px solid #ddd;">${rows}</table>`;
}

/**
 * Sends an email, like mailing a letter.
 * @param {string} to Who gets the letter.
 * @param {string} subject The letter’s title.
 * @param {string} bodyHtml The letter’s message.
 */
function sendEmail(to, subject, bodyHtml) {
  if (!to || !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(to)) {
    Logger.log(`Bad email address: ${to}`); // Oops, wrong address
    return;
  }
  try {
    MailApp.sendEmail({
      to: to,
      subject: subject,
      htmlBody: bodyHtml + EMAIL_REGARDS + EMAIL_FOOTER, // Add signature and footer
      name: SENDER_NAME // From the toy team
    });
  } catch (e) {
    Logger.log(`Error sending email to ${to}: ${e.message}`);
  }
}

/**
 * Checks if a Finance Officer is assigned for a GroupID in the Control tab.
 * @param {string} groupID The loan code.
 * @returns {boolean} True if assigned, false if not.
 */
function checkFinanceOfficerExists(groupID) {
  const sheet = getControlSheet(); // Open Control tab
  const data = sheet.getDataRange().getValues(); // Read all loans
  return data.some(row => row[0] === groupID && row[37] && row[38] && row[39] && row[40]); // Check Finance Officer fields
}

/**
 * Gets the Finance Officer’s info for a GroupID from the Control tab.
 * @param {string} groupID The loan code.
 * @returns {Object|null} Their details or null if not found.
 */
function getFinanceOfficer(groupID) {
  const sheet = getControlSheet(); // Open Control tab
  const data = sheet.getDataRange().getValues(); // Read loans
  const row = data.find(row => row[0] === groupID); // Find GroupID
  if (!row || !row[37] || !row[38] || !row[39] || !row[40]) return null; // No Finance Officer
  return {
    name: row[37], // FinanceOfficerName
    id: row[38], // FinanceOfficerID
    email: row[39], // FinanceOfficerEmail
    phone: row[40] // FinanceOfficerPhone
  };
}

/**
 * Updates dropdowns in the Application Form, like filling a toy’s choices.
 */
function syncParticipantDetails() {
  const intentSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(LOAN_INTENT_SHEET_NAME);
  const controlSheet = getControlSheet();
  const intentData = intentSheet.getDataRange().getValues().slice(1); // Skip header
  const controlData = controlSheet.getDataRange().getValues().slice(1); // Skip header
  try {
    const form = FormApp.openById(LOAN_APPLICATION_FORM_ID); // Open Application Form
    const items = form.getItems(); // Get questions

    // Choices for each question
    const fieldMap = {
      'Cooperator ID': intentData.filter(row => row[6] === 'Applicant').map(row => row[2]),
      'Name': intentData.filter(row => row[6] === 'Applicant').map(row => row[3]),
      'Phone': intentData.filter(row => row[6] === 'Applicant').map(row => row[4]),
      'Email': intentData.filter(row => row[6] === 'Applicant').map(row => row[5]),
      'Guarantor 1 Name': intentData.filter(row => row[6] === 'Guarantor').map(row => row[3]),
      'Guarantor 1 Cooperator ID': intentData.filter(row => row[6] === 'Guarantor').map(row => row[2]),
      'Guarantor 1 Email': intentData.filter(row => row[6] === 'Guarantor').map(row => row[5]),
      'Guarantor 1 Phone': intentData.filter(row => row[6] === 'Guarantor').map(row => row[4]),
      'Guarantor 2 Name': intentData.filter(row => row[6] === 'Guarantor').map(row => row[3]),
      'Guarantor 2 Cooperator ID': intentData.filter(row => row[6] === 'Guarantor').map(row => row[2]),
      'Guarantor 2 Email': intentData.filter(row => row[6] === 'Guarantor').map(row => row[5]),
      'Guarantor 2 Phone': intentData.filter(row => row[6] === 'Guarantor').map(row => row[4]),
      'Approver ID': controlData.filter(row => row[38]).map(row => row[38]),
      'Approver Name': controlData.filter(row => row[37]).map(row => row[37]),
      'Approver Email': controlData.filter(row => row[39]).map(row => row[39]),
      'Approver Phone': controlData.filter(row => row[40]).map(row => row[40])
    };

    items.forEach(item => {
      const title = item.getTitle(); // Question name
      if (fieldMap[title] && item.asListItem()) {
        item.asListItem().setChoiceValues(fieldMap[title]); // Set choices
      }
    });
  } catch (e) {
    Logger.log(`Error syncing participant details: ${e.message}`);
  }
}

/**
 * Handles Intent Form submissions, like adding new sign-up notes.
 * @param {GoogleAppsScript.Forms.FormSubmitEvent} e The form event.
 */
function onIntentFormSubmit(e) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(LOAN_INTENT_SHEET_NAME); // Intent page
  const formResponses = e.namedValues; // Form answers

  // Get answers
  const applicantID = formResponses['Your Cooperator ID']?.[0] || '';
  const applicantName = formResponses['Your Name']?.[0] || '';
  const applicantPhone = formResponses['Your Phone']?.[0] || '';
  const applicantEmail = formResponses['Your Email']?.[0] || '';
  const guarantor1ID = formResponses['Guarantor 1 Cooperator ID']?.[0] || '';
  const guarantor1Name = formResponses['Guarantor 1 Name']?.[0] || '';
  const guarantor1Phone = formResponses['Guarantor 1 Phone']?.[0] || '';
  const guarantor1Email = formResponses['Guarantor 1 Email']?.[0] || '';
  const guarantor2ID = formResponses['Guarantor 2 Cooperator ID']?.[0] || '';
  const guarantor2Name = formResponses['Guarantor 2 Name']?.[0] || '';
  const guarantor2Phone = formResponses['Guarantor 2 Phone']?.[0] || '';
  const guarantor2Email = formResponses['Guarantor 2 Email']?.[0] || '';

  // Check if all answers are there
  if (!applicantID || !applicantName || !applicantPhone || !applicantEmail ||
      !guarantor1ID || !guarantor1Name || !guarantor1Phone || !guarantor1Email ||
      !guarantor2ID || !guarantor2Name || !guarantor2Phone || !guarantor2Email) {
    Logger.log('Missing answers in Intent Form');
    return;
  }

  // Make a new GroupID
  const groupID = generateGroupID();

  // Add three rows: applicant, guarantor1, guarantor2
  const timestamp = new Date();
  const rows = [
    [timestamp, groupID, applicantID, applicantName, applicantPhone, applicantEmail, 'Applicant'],
    [timestamp, groupID, guarantor1ID, guarantor1Name, guarantor1Phone, guarantor1Email, 'Guarantor'],
    [timestamp, groupID, guarantor2ID, guarantor2Name, guarantor2Phone, guarantor2Email, 'Guarantor']
  ];

  // Check for duplicates (GroupID + CooperatorID)
  const existingData = sheet.getDataRange().getValues().slice(1); // Skip header
  for (const newRow of rows) {
    const newGroupID = newRow[1];
    const newCooperatorID = newRow[2];
    if (existingData.some(row => row[1] === newGroupID && row[2] === newCooperatorID)) {
      Logger.log(`Duplicate found: GroupID ${newGroupID}, CooperatorID ${newCooperatorID}`);
      continue; // Skip duplicates
    }
    sheet.appendRow(newRow); // Add new row
  }

  // Create a placeholder row in Control sheet
  const controlSheet = getControlSheet();
  controlSheet.appendRow([groupID, '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', 'PendingFinanceOfficer', 'FALSE', timestamp]);

  // Send one email to applicant with GroupID
  const prefillApplicantLink = generatePrefilledLink(groupID, ROLE_APPLICANT, applicantEmail);
  const intentData = {
    'GroupID': groupID,
    'Applicant Name': applicantName,
    'Applicant Cooperator ID': applicantID,
    'Applicant Email': applicantEmail,
    'Applicant Phone': applicantPhone,
    'Guarantor 1 Name': guarantor1Name,
    'Guarantor 1 Cooperator ID': guarantor1ID,
    'Guarantor 1 Email': guarantor1Email,
    'Guarantor 1 Phone': guarantor1Phone,
    'Guarantor 2 Name': guarantor2Name,
    'Guarantor 2 Cooperator ID': guarantor2ID,
    'Guarantor 2 Email': guarantor2Email,
    'Guarantor 2 Phone': guarantor2Phone
  };
  const intentTable = buildHtmlTableFromObject(intentData);
  sendEmail(applicantEmail,
    `Loan Intent Submitted - ${groupID}`,
    `<p>Dear ${applicantName},</p><p>Your loan sign-up (GroupID: ${groupID}) was received!</p>
    ${intentTable}
    <p>Please wait for an admin to assign a Finance Officer before using this link to apply: <a href="${prefillApplicantLink}">Loan Application Form</a></p>`);
}

/**
 * Handles Application Form submissions, like adding loan details.
 * @param {GoogleAppsScript.Events.SheetsOnFormSubmit} e The form event.
 */
function onApplicationFormSubmit(e) {
  const sheet = getControlSheet(); // Open Control tab
  const formResponses = e.namedValues; // Get answers

  // Get answers
  const role = formResponses['Role']?.[0] || '';
  let groupID = formResponses['Loan ID']?.[0] || '';
  const applicantID = formResponses['Cooperator ID']?.[0] || '';
  const applicantName = formResponses['Name']?.[0] || '';
  const applicantPhone = formResponses['Phone']?.[0] || '';
  const applicantEmail = formResponses['Email']?.[0] || '';
  const applicantAddress = formResponses['Home Address']?.[0] || '';
  const loanAmountFigures = formResponses['Loan Amount (Figures)']?.[0] || '';
  const loanAmountWords = formResponses['Loan Amount (Words)']?.[0] || '';
  const loanDuration = formResponses['Repayment Period']?.[0] || '';
  const bankName = formResponses['Bank Name']?.[0] || '';
  const accountName = formResponses['Account Name']?.[0] || '';
  const accountNumber = formResponses['Account Number']?.[0] || '';
  const guarantor1Name = formResponses['Guarantor 1 Name']?.[0] || '';
  const guarantor1ID = formResponses['Guarantor 1 Cooperator ID']?.[0] || '';
  const guarantor1Email = formResponses['Guarantor 1 Email']?.[0] || '';
  const guarantor1Phone = formResponses['Guarantor 1 Phone']?.[0] || '';
  const guarantor2Name = formResponses['Guarantor 2 Name']?.[0] || '';
  const guarantor2ID = formResponses['Guarantor 2 Cooperator ID']?.[0] || '';
  const guarantor2Email = formResponses['Guarantor 2 Email']?.[0] || '';
  const guarantor2Phone = formResponses['Guarantor 2 Phone']?.[0] || '';
  const approverName = formResponses['Approver Name']?.[0] || '';
  const approverID = formResponses['Approver ID']?.[0] || '';
  const approverEmail = formResponses['Approver Email']?.[0] || '';
  const approverPhone = formResponses['Approver Phone']?.[0] || '';
  const decision = formResponses['Status']?.[0] || '';
  const applicantBalance = formResponses['Applicant Balance']?.[0] || '';
  const applicantRating = formResponses['Applicant Rating']?.[0] || '';
  const guarantor1Balance = formResponses['Guarantor 1 Balance']?.[0] || '';
  const guarantor1Rating = formResponses['Guarantor 1 Rating']?.[0] || '';
  const guarantor2Balance = formResponses['Guarantor 2 Balance']?.[0] || '';
  const guarantor2Rating = formResponses['Guarantor 2 Rating']?.[0] || '';
  const comments = formResponses['Comments']?.[0] || '';

  // Check if answers are complete
  if (!role || (role === ROLE_APPLICANT && (!applicantID || !applicantEmail || !applicantPhone || !applicantAddress || !loanAmountFigures || !loanAmountWords || !loanDuration || !guarantor1ID || !guarantor1Email || !guarantor2ID || !guarantor2Email)) ||
      ((role === ROLE_GUARANTOR1 || role === ROLE_GUARANTOR2) && (!guarantor1ID || !guarantor1Email || !guarantor1Phone || !guarantor1Name || !guarantor2ID || !guarantor2Email || !guarantor2Phone || !guarantor2Name)) ||
      (role === ROLE_FINANCE && (!approverID || !approverEmail || !approverPhone || !decision))) {
    Logger.log(`Missing answers: Role=${role}, GroupID=${groupID}`);
    return;
  }

  // Check if Finance Officer is assigned
  if (role === ROLE_APPLICANT && !checkFinanceOfficerExists(groupID)) {
    sendEmail(applicantEmail,
      `Loan Application Blocked - No Finance Officer`,
      `<p>Dear ${applicantName},</p><p>No Finance Officer is assigned for GroupID ${groupID}. Please wait or contact an admin.</p>`);
    Logger.log(`Blocked submission: No Finance Officer for ${groupID}`);
    return;
  }

  // Get Finance Officer details
  const financeOfficer = getFinanceOfficer(groupID);
  const approverDetails = financeOfficer || { name: '', id: '', email: '', phone: '' };

  const data = sheet.getDataRange().getValues(); // Read Control tab
  let rowIndex = -1;
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === groupID) {
      rowIndex = i + 1;
      break;
    }
  }

  // Handle Applicant submission
  if (role === ROLE_APPLICANT) {
    for (let i = 1; i < data.length; i++) {
      if (data[i][1] === applicantID && !['FinanceReviewed', 'Expired'].includes(data[i][24])) {
        sendEmail(applicantEmail,
          `Active Application Denied - ${data[i][0]}`,
          `<p>Dear ${applicantName},</p><p>You have an active loan (GroupID: ${data[i][0]}). Contact an admin to reset.</p>`);
        Logger.log(`Blocked: Active GroupID=${data[i][0]}`);
        return;
      }
    }
    if (!groupID || rowIndex === -1) {
      groupID = generateGroupID(); // New loan
    } else if (sheet.getRange(rowIndex, 26).getValue() === 'TRUE') {
      sendEmail(applicantEmail,
        `Submission Locked - ${groupID}`,
        `<p>Dear ${applicantName},</p><p>Application (GroupID: ${groupID}) is locked. Contact an admin.</p>`);
      Logger.log(`Blocked re-edit for GroupID ${groupID}`);
      return;
    }
    const newRow = [
      groupID, applicantID, applicantName, applicantEmail, applicantPhone, applicantAddress,
      loanAmountFigures, loanAmountWords, loanDuration,
      guarantor1Name, guarantor1ID, guarantor1Email, guarantor1Phone,
      guarantor2Name, guarantor2ID, guarantor2Email, guarantor2Phone,
      approverName, approverID, approverEmail, approverPhone, decision,
      '', '', 'ApplicantSubmitted', 'FALSE', new Date(),
      comments, bankName, accountName, accountNumber,
      applicantBalance, applicantRating, guarantor1Balance, guarantor1Rating,
      guarantor2Balance, guarantor2Rating,
      approverDetails.name, approverDetails.id, approverDetails.email, approverDetails.phone
    ];
    if (rowIndex === -1) {
      sheet.appendRow(newRow); // Add new loan
      rowIndex = sheet.getLastRow();
    } else {
      sheet.getRange(rowIndex, 1, 1, 40).setValues([newRow]); // Update existing
    }
  } else if (role === ROLE_GUARANTOR1 || role === ROLE_GUARANTOR2) {
    if (!groupID || rowIndex === -1) {
      Logger.log(`Guarantor submission without GroupID: ${role}`);
      return;
    }
    if (sheet.getRange(rowIndex, 26).getValue() === 'TRUE') {
      sendEmail(role === ROLE_GUARANTOR1 ? guarantor1Email : guarantor2Email,
        `Submission Locked - ${groupID}`,
        `<p>Dear ${role === ROLE_GUARANTOR1 ? guarantor1Name : guarantor2Name},</p><p>Application (GroupID: ${groupID}) is locked. Contact an admin.</p>`);
      Logger.log(`Blocked guarantor for GroupID ${groupID}`);
      return;
    }
    const guarantorData = role === ROLE_GUARANTOR1
      ? { name: guarantor1Name, id: guarantor1ID, email: guarantor1Email, phone: guarantor1Phone, col: 10 }
      : { name: guarantor2Name, id: guarantor2ID, email: guarantor2Email, phone: guarantor2Phone, col: 14 };
    sheet.getRange(rowIndex, guarantorData.col, 1, 4).setValues([[
      guarantorData.name, guarantorData.id, guarantorData.email, guarantorData.phone
    ]]);
  } else if (role === ROLE_FINANCE) {
    if (!groupID || rowIndex === -1) {
      Logger.log(`Finance submission without GroupID: ${role}`);
      return;
    }
    if (sheet.getRange(rowIndex, 26).getValue() === 'TRUE') {
      Logger.log(`Submission locked for GroupID ${groupID}, Role ${role}`);
      return;
    }
    sheet.getRange(rowIndex, 18, 1, 23).setValues([[
      approverName, approverID, approverEmail, approverPhone, decision,
      '', '', '', '', '', '', comments,
      bankName, accountName, accountNumber,
      applicantBalance, applicantRating, guarantor1Balance, guarantor1Rating,
      guarantor2Balance, guarantor2Rating,
      approverDetails.name, approverDetails.id
    ]]);
    sheet.getRange(rowIndex, 25).setValue('FinanceReviewed');
    sheet.getRange(rowIndex, 26).setValue('TRUE');
  } else {
    Logger.log(`Invalid role: ${role}`);
    return;
  }

  // Make links for everyone
  const prefillApplicantLink = generatePrefilledLink(groupID, ROLE_APPLICANT, applicantEmail);
  const prefillGuarantor1Link = generatePrefilledLink(groupID, ROLE_GUARANTOR1, guarantor1Email);
  const prefillGuarantor2Link = generatePrefilledLink(groupID, ROLE_GUARANTOR2, guarantor2Email);
  const prefillFinanceLink = generatePrefilledLink(groupID, ROLE_FINANCE, approverDetails.email);
  sheet.getRange(rowIndex, 23, 1, 2).setValues([[prefillApplicantLink, prefillFinanceLink]]);

  // Prepare tables for emails
  const guarantor1Data = {
    'Loan ID': groupID,
    'Applicant Cooperator ID': applicantID,
    'Applicant Name': applicantName,
    'Applicant Phone': applicantPhone,
    'Applicant Home Address': applicantAddress,
    'Loan Amount (Figures)': loanAmountFigures,
    'Loan Amount (Words)': loanAmountWords,
    'Repayment Period': loanDuration,
    'Guarantor1 ID': guarantor1ID,
    'Status': sheet.getRange(rowIndex, 25).getValue()
  };
  const guarantor2Data = {
    'Loan ID': groupID,
    'Applicant Cooperator ID': applicantID,
    'Applicant Name': applicantName,
    'Applicant Phone': applicantPhone,
    'Applicant Home Address': applicantAddress,
    'Loan Amount (Figures)': loanAmountFigures,
    'Loan Amount (Words)': loanAmountWords,
    'Repayment Period': loanDuration,
    'Guarantor2 ID': guarantor2ID,
    'Status': sheet.getRange(rowIndex, 25).getValue()
  };
  const financeData = {
    'Loan ID': groupID,
    'Applicant Cooperator ID': applicantID,
    'Applicant Name': applicantName,
    'Applicant Email': applicantEmail,
    'Applicant Phone': applicantPhone,
    'Applicant Home Address': applicantAddress,
    'Loan Amount (Figures)': loanAmountFigures,
    'Loan Amount (Words)': loanAmountWords,
    'Repayment Period': loanDuration,
    'Bank Name': bankName,
    'Account Name': accountName,
    'Account Number': accountNumber,
    'Guarantor1 Name': guarantor1Name,
    'Guarantor1 ID': guarantor1ID,
    'Guarantor1 Email': guarantor1Email,
    'Guarantor1 Phone': guarantor1Phone,
    'Guarantor2 Name': guarantor2Name,
    'Guarantor2 ID': guarantor2ID,
    'Guarantor2 Email': guarantor2Email,
    'Guarantor2 Phone': guarantor2Phone,
    'Approver Name': approverName,
    'Approver ID': approverID,
    'Approver Email': approverEmail,
    'Approver Phone': approverPhone,
    'Applicant Balance': applicantBalance,
    'Applicant Rating': applicantRating,
    'Guarantor1 Balance': guarantor1Balance,
    'Guarantor1 Rating': guarantor1Rating,
    'Guarantor2 Balance': guarantor2Balance,
    'Guarantor2 Rating': guarantor2Rating,
    'Status': decision || 'Pending',
    'Comments': comments
  };
  const guarantor1Table = buildHtmlTableFromObject(guarantor1Data);
  const guarantor2Table = buildHtmlTableFromObject(guarantor2Data);
  const financeTable = buildHtmlTableFromObject(financeData);

  // Send emails based on role
  switch (role) {
    case ROLE_APPLICANT:
      let recipient = applicantEmail;
      let subject = `Loan Application Submitted - ${groupID}`;
      let htmlBody = `<p>Dear ${applicantName},</p><p>Your loan application was received.</p>
        ${financeTable}<p></p>
        <p>Your guarantors and the Finance Officer were notified. Edit your application: <a href="${prefillApplicantLink}">Edit Application</a></p>
        <p>Guarantors must respond within ${REMINDER_DAYS_LIMIT} days.</p>`;
      sendEmail(recipient, subject, htmlBody);

      recipient = guarantor1Email;
      subject = `Action Needed: Guarantor Details - ${groupID}`;
      htmlBody = `<p>Dear ${guarantor1Name},</p><p>You’re Guarantor 1 for ${applicantName}’s loan (GroupID: ${groupID}). Submit your details within ${REMINDER_DAYS_LIMIT} days: <a href="${prefillGuarantor1Link}">Complete Details</a></p>
        ${guarantor1Table}`;
      sendEmail(recipient, subject, htmlBody);

      recipient = guarantor2Email;
      subject = `Action Needed: Guarantor Details - ${groupID}`;
      htmlBody = `<p>Dear ${guarantor2Name},</p><p>You’re Guarantor 2 for ${applicantName}’s loan (GroupID: ${groupID}). Submit your details within ${REMINDER_DAYS_LIMIT} days: <a href="${prefillGuarantor2Link}">Complete Details</a></p>
        ${guarantor2Table}`;
      sendEmail(recipient, subject, htmlBody);

      if (approverDetails.email) {
        subject = `Review Loan Application - ${groupID}`;
        htmlBody = `<p>Dear ${approverDetails.name},</p><p>Review ${applicantName}’s loan application (GroupID: ${groupID}).</p>
        ${financeTable}<p></p>
        <p>Review here: <a href="${prefillFinanceLink}">Review Application</a></p>`;
        sendEmail(approverDetails.email, subject, htmlBody);
      }
      break;

    case ROLE_GUARANTOR1:
    case ROLE_GUARANTOR2:
      const guarantorName = role === ROLE_GUARANTOR1 ? guarantor1Name : guarantor2Name;
      const guarantorEmail = role === ROLE_GUARANTOR1 ? guarantor1Email : guarantor2Email;
      const guarantorTable = role === ROLE_GUARANTOR1 ? guarantor1Table : guarantor2Table;
      [applicantEmail, approverDetails.email].filter(email => email).forEach(email => {
        sendEmail(email,
          `Guarantor Details Submitted - ${groupID}`,
          `<p>Dear ${email === applicantEmail ? applicantName : approverDetails.name},</p><p>${guarantorName} (${role}) submitted details for GroupID: ${groupID}.</p>
          ${financeTable}`);
      });
      recipient = guarantorEmail;
      subject = `Guarantor Details Submitted - ${groupID}`;
      htmlBody = `<p>Dear ${guarantorName},</p><p>Your details as ${role} for ${applicantName}’s loan (GroupID: ${groupID}) were submitted.</p>
        ${guarantorTable}`;
      sendEmail(recipient, subject, htmlBody);
      break;

    case ROLE_FINANCE:
      const recipients = [applicantEmail, guarantor1Email, guarantor2Email].filter(email => email);
      subject = `Loan Application Reviewed - ${groupID}`;
      recipients.forEach(recipientEmail => {
        if (recipientEmail === applicantEmail || recipientEmail === approverDetails.email) {
          htmlBody = `<p>Dear ${recipientEmail === applicantEmail ? applicantName : approverDetails.name},</p><p>Loan application (GroupID: ${groupID}) was reviewed by ${approverDetails.name}.</p>
          ${financeTable}<p>Status: ${financeData['Status']}</p>`;
        } else {
          const guarantorTable = recipientEmail === guarantor1Email ? guarantor1Table : guarantor2Table;
          htmlBody = `<p>Dear ${recipientEmail === guarantor1Email ? guarantor1Name : guarantor2Name},</p><p>Loan application (GroupID: ${groupID}) was reviewed by ${approverDetails.name}.</p>
            ${guarantorTable}<p>Status: ${financeData['Status']}</p>`;
        }
        sendEmail(recipientEmail, subject, htmlBody);
      });
      archiveApplication(groupID);
      break;
  }
}

/**
 * Sends daily reminders, like reminding friends to play.
 * Only for applications with Finance Officer assigned and within 7 days.
 */
function sendDailyReminders() {
  const sheet = getControlSheet(); // Open Control tab
  const data = sheet.getDataRange().getValues(); // Read all loans
  const now = new Date(); // Today’s date

  // Check each loan
  for (let i = 1; i < data.length; i++) {
    const row = data[i]; // Current loan
    const groupID = row[0]; // Group ID
    const applicantName = row[2]; // Applicant
    const applicantEmail = row[3]; // Email
    const guarantor1Name = row[9]; // Guarantor 1
    const guarantor1Email = row[11]; // Email
    const guarantor2Name = row[13]; // Guarantor 2
    const guarantor2Email = row[15]; // Email
    const approverName = row[37]; // Finance Officer
    const approverEmail = row[39]; // Email
    const status = row[24]; // Status
    const prefillGuarantor1Link = generatePrefilledLink(groupID, ROLE_GUARANTOR1, guarantor1Email);
    const prefillGuarantor2Link = generatePrefilledLink(groupID, ROLE_GUARANTOR2, guarantor2Email);
    const prefillFinanceLink = row[23]; // Finance link
    const timestamp = new Date(row[26]); // When submitted
    const daysSince = Math.floor((now - timestamp) / (1000 * 60 * 60 * 24)); // Days passed

    // Skip if no Finance Officer or already done
    if (!checkFinanceOfficerExists(groupID) || status !== 'ApplicantSubmitted') continue;

    // If too late (7 days), mark as expired
    if (daysSince > REMINDER_DAYS_LIMIT) {
      sheet.getRange(i + 1, 25).setValue('Expired'); // Mark expired
      sheet.getRange(i + 1, 26).setValue('TRUE'); // Lock it
      archiveApplication(groupID); // Move to Archive
      sendEmail(applicantEmail,
        `Loan Application Expired - ${groupID}`,
        `<p>Dear ${applicantName},</p><p>Your loan (GroupID: ${groupID}) expired because guarantors didn’t respond. Contact an admin to resubmit.</p>`);
      continue;
    }

    // Make a reminder table
    const reminderData = {
      'Loan ID': groupID,
      'Applicant Name': applicantName,
      'Status': status
    };
    const reminderTable = buildHtmlTableFromObject(reminderData);

    // Remind Guarantor 1
    if (guarantor1Email && prefillGuarantor1Link) {
      sendEmail(guarantor1Email,
        `Reminder: Guarantor Details - ${groupID}`,
        `<p>Dear ${guarantor1Name},</p><p>Please submit your details for ${applicantName}’s loan.</p>
        ${reminderTable}<p></p>
        <p>Submit within ${REMINDER_DAYS_LIMIT - daysSince} days: <a href="${prefillGuarantor1Link}">Complete Details</a></p>`);
    }

    // Remind Guarantor 2
    if (guarantor2Email && prefillGuarantor2Link) {
      sendEmail(guarantor2Email,
        `Reminder: Guarantor Details - ${groupID}`,
        `<p>Dear ${guarantor2Name},</p><p>Please submit your details for ${applicantName}’s loan.</p>
        ${reminderTable}<p></p>
        <p>Submit within ${REMINDER_DAYS_LIMIT - daysSince} days: <a href="${prefillGuarantor2Link}">Complete Details</a></p>`);
    }

    // Remind Finance Officer
    if (approverEmail && prefillFinanceLink) {
      const financeReminderData = {
        'Loan ID': groupID,
        'Applicant Name': applicantName,
        'Status': status,
        'Applicant Cooperator ID': row[1],
        'Applicant Email': applicantEmail,
        'Applicant Phone': row[4],
        'Applicant Home Address': row[5],
        'Loan Amount (Figures)': row[6],
        'Loan Amount (Words)': row[7],
        'Repayment Period': row[8]
      };
      const financeTable = buildHtmlTableFromObject(financeReminderData);
      sendEmail(approverEmail,
        `Reminder: Review Loan Application - ${groupID}`,
        `<p>Dear ${approverName},</p><p>Please review ${applicantName}’s loan (GroupID: ${groupID}).</p>
        ${financeTable}<p></p>
        <p>Review: <a href="${prefillFinanceLink}">Review Application</a></p>`);
    }
  }
}

/**
 * Sends notifications to newly assigned Finance Officer groups.
 * Only for groups with Finance Officer and no application flow started.
 */
function notifyNewFinanceOfficerAssignments() {
  const sheet = getControlSheet(); // Open Control sheet
  const intentSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(LOAN_INTENT_SHEET_NAME); // Intent sheet
  const controlData = sheet.getDataRange().getValues().slice(1); // Skip header
  const intentData = intentSheet.getDataRange().getValues().slice(1); // Skip header

  const ui = SpreadsheetApp.getUi(); // Get UI for prompt
  const response = ui.alert('Notify New Assignments',
    'Send notifications to groups with newly assigned Finance Officers?',
    ui.ButtonSet.YES_NO); // Ask admin
  if (response !== ui.Button.YES) {
    ui.alert('Notifications cancelled.');
    return;
  }

  let notifiedCount = 0;

  // Process each loan group
  for (let i = 0; i < controlData.length; i++) {
    const row = controlData[i];
    const groupID = row[0]; // Group ID
    const applicantEmail = row[3]; // Applicant email
    const applicantName = row[2]; // Applicant name
    const status = row[24]; // Application Status
    const financeOfficer = getFinanceOfficer(groupID); // Finance Officer

    // Skip if no Finance Officer or application already started
    if (!financeOfficer || status === 'ApplicantSubmitted' || status === 'FinanceReviewed' || status === 'Expired' || status === 'Notified') continue;

    // Get guarantor details from Intent sheet
    const groupIntentData = intentData.filter(row => row[1] === groupID);
    const applicantRow = groupIntentData.find(row => row[6] === 'Applicant');
    const guarantorRows = groupIntentData.filter(row => row[6] === 'Guarantor');
    const guarantor1Row = guarantorRows[0];
    const guarantor2Row = guarantorRows[1];

    if (!applicantRow || !guarantor1Row || !guarantor2Row) {
      Logger.log(`Missing intent data for GroupID: ${groupID}`);
      continue;
    }

    const guarantor1Email = guarantor1Row[5] || '';
    const guarantor1Name = guarantor1Row[3] || '';
    const guarantor2Email = guarantor2Row[5] || '';
    const guarantor2Name = guarantor2Row[3] || '';

    // Generate links
    const prefillApplicantLink = generatePrefilledLink(groupID, ROLE_APPLICANT, applicantEmail);
    const prefillGuarantor1Link = generatePrefilledLink(groupID, ROLE_GUARANTOR1, guarantor1Email);
    const prefillGuarantor2Link = generatePrefilledLink(groupID, ROLE_GUARANTOR2, guarantor2Email);
    const prefillFinanceLink = generatePrefilledLink(groupID, ROLE_FINANCE, financeOfficer.email);

    // Prepare notification data
    const notificationData = {
      'Loan ID': groupID,
      'Applicant Name': applicantName,
      'Finance Officer Name': financeOfficer.name,
      'Finance Officer Email': financeOfficer.email
    };
    const notificationTable = buildHtmlTableFromObject(notificationData);

    // Notify Applicant
    sendEmail(applicantEmail,
      `Finance Officer Assigned - ${groupID}`,
      `<p>Dear ${applicantName},</p><p>A Finance Officer has been assigned to your loan group (GroupID: ${groupID}). You can now submit your application.</p>
      ${notificationTable}<p></p>
      <p>Submit here: <a href="${prefillApplicantLink}">Loan Application Form</a></p>`);

    // Notify Guarantor 1
    if (guarantor1Email) {
      sendEmail(guarantor1Email,
        `Finance Officer Assigned - ${groupID}`,
        `<p>Dear ${guarantor1Name},</p><p>A Finance Officer has been assigned to ${applicantName}’s loan (GroupID: ${groupID}). Be ready to submit your details.</p>
        ${notificationTable}<p></p>
        <p>Your link: <a href="${prefillGuarantor1Link}">${groupID} Application</a></p>`);
    }

    // Notify Guarantor 2
    if (guarantor2Email) {
      sendEmail(guarantor2Email,
        `Finance Officer Assigned - ${groupID}`,
        `<p>Dear ${guarantor2Name},</p><p>A Finance Officer has been assigned to ${applicantName}’s loan (GroupID: ${groupID}). Be ready to submit your details.</p>
        ${notificationTable}<p></p>
        <p>Your link: <a href="${prefillGuarantor2Link}">${groupID} Application</a></p>`);
    }

    // Notify Finance Officer
    sendEmail(financeOfficer.email,
      `Assigned to Loan Group - ${groupID}`,
      `<p>Dear ${financeOfficer.name},</p><p>You’ve been assigned as Finance Officer for GroupID ${groupID} (Applicant: ${applicantName}).</p>
      ${notificationTable}<p></p>
      <p>Review application: <a href="${prefillFinanceLink}">Review Application</a></p>`);

    // Update status to reflect notification sent
    sheet.getRange(i + 2, 25).setValue('Notified');

    Logger.log(`Sent notifications for GroupID ${groupID}`);
    notifiedCount++;
  }

  ui.alert(`Notifications sent to ${notifiedCount} newly assigned groups!`);
}

/**
 * Moves a loan to the Archive tab, like putting away a toy.
 * @param {string} groupID The loan code.
 * @returns {string} What happened.
 */
function archiveApplication(groupID) {
  const controlSheet = getControlSheet(); // Open Control sheet
  const archiveSheet = getArchiveSheet(); // Open Archive sheet
  const data = controlSheet.getDataRange().getValues(); // Read all loans

  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === groupID && ['FinanceReviewed', 'Expired'].includes(data[i][24])) {
      archiveSheet.appendRow(data[i]); // Add to Archive
      controlSheet.deleteRow(i + 1); // Remove from Control
      Logger.log(`Archived GroupID ${groupID}`);
      return `Archived GroupID ${groupID}`;
    }
  }
  return `GroupID ${groupID} not found or not ready to archive`;
}

/**
 * Resets a loan application, like giving a toy a fresh start.
 * @param {string} groupID The loan code.
 * @returns {string} What happened.
 */
function manualResetUnlock(groupID) {
  const sheet = getControlSheet(); // Open Control tab
  const data = sheet.getDataRange().getValues(); // Read loans
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === groupID) {
      const applicantName = data[i][2];
      const applicantEmail = data[i][3];
      const guarantor1Name = data[i][9];
      const guarantor1Email = data[i][11];
      const guarantor2Name = data[i][13];
      const guarantor2Email = data[i][15];
      const approverName = data[i][37];
      const approverEmail = data[i][39];
      const guarantor1Data = {
        'Loan ID': groupID,
        'Applicant Cooperator ID': data[i][1],
        'Applicant Name': applicantName,
        'Applicant Phone': data[i][4],
        'Applicant Home Address': data[i][5],
        'Loan Amount (Figures)': data[i][6],
        'Loan Amount (Words)': data[i][7],
        'Repayment Period': data[i][8],
        'Guarantor1 ID': data[i][10],
        'Status': 'Pending'
      };
      const guarantor2Data = {
        'Loan ID': groupID,
        'Applicant Cooperator ID': data[i][1],
        'Applicant Name': applicantName,
        'Applicant Phone': data[i][4],
        'Applicant Home Address': data[i][5],
        'Loan Amount (Figures)': data[i][6],
        'Loan Amount (Words)': data[i][7],
        'Repayment Period': data[i][8],
        'Guarantor2 ID': data[i][14],
        'Status': 'Pending'
      };
      const financeData = {
        'Loan ID': groupID,
        'Applicant Cooperator ID': data[i][1],
        'Applicant Name': applicantName,
        'Applicant Email': applicantEmail,
        'Applicant Phone': data[i][4],
        'Applicant Home Address': data[i][5],
        'Loan Amount (Figures)': data[i][6],
        'Loan Amount (Words)': data[i][7],
        'Repayment Period': data[i][8],
        'Guarantor1 Name': guarantor1Name,
        'Guarantor1 ID': data[i][10],
        'Guarantor1 Email': guarantor1Email,
        'Guarantor1 Phone': data[i][12],
        'Guarantor2 Name': guarantor2Name,
        'Guarantor2 ID': data[i][14],
        'Guarantor2 Email': guarantor2Email,
        'Guarantor2 Phone': data[i][16],
        'Approver Name': approverName,
        'Approver ID': data[i][38],
        'Approver Email': approverEmail,
        'Approver Phone': data[i][40],
        'Status': data[i][21] || 'Pending',
        'Comments': data[i][27]
      };
      const guarantor1Table = buildHtmlTableFromObject(guarantor1Data);
      const guarantor2Table = buildHtmlTableFromObject(guarantor2Data);
      const financeTable = buildHtmlTableFromObject(financeData);

      sheet.getRange(i + 1, 25).setValue(''); // Clear status
      sheet.getRange(i + 1, 26).setValue('FALSE'); // Unlock
      const prefillApplicantLink = generatePrefilledLink(groupID, ROLE_APPLICANT, applicantEmail);
      const prefillGuarantor1Link = generatePrefilledLink(groupID, ROLE_GUARANTOR1, guarantor1Email);
      const prefillGuarantor2Link = generatePrefilledLink(groupID, ROLE_GUARANTOR2, guarantor2Email);
      const subject = `Application Reset - ${groupID}`;
      const recipients = [applicantEmail, guarantor1Email, guarantor2Email, approverEmail].filter(email => email);
      recipients.forEach(recipient => {
        let htmlBody = '';
        if (recipient === guarantor1Email) {
          htmlBody = `<p>Dear ${guarantor1Name},</p><p>Application (GroupID: ${groupID}) was reset. Resubmit: <a href="${prefillGuarantor1Link}">Complete Details</a></p>
            ${guarantor1Table}`;
        } else if (recipient === guarantor2Email) {
          htmlBody = `<p>Dear ${guarantor2Name},</p><p>Application (GroupID: ${groupID}) was reset. Resubmit: <a href="${prefillGuarantor2Link}">Complete Details</a></p>
            ${guarantor2Table}`;
        } else {
          htmlBody = `<p>Dear ${recipient === applicantEmail ? applicantName : approverName},</p><p>Application (GroupID: ${groupID}) was reset.</p>
          ${financeTable}<p></p>
          <p>Applicant: <a href="${prefillApplicantLink}">Resubmit</a></p>
          <p>Guarantor 1: <a href="${prefillGuarantor1Link}">Resubmit</a></p>
          <p>Guarantor 2: <a href="${prefillGuarantor2Link}">Resubmit</a></p>`;
        }
        sendEmail(recipient, subject, htmlBody);
      });
      Logger.log(`Reset GroupID ${groupID}`);
      return `Reset GroupID ${groupID}`;
    }
  }
  return `GroupID ${groupID} not found`;
}

/**
 * Adds buttons to the spreadsheet, like putting controls on our toy.
 */
function createMenu() {
  const ui = SpreadsheetApp.getUi(); // Get the notebook’s buttons
  ui.createMenu('CSULMCS Loan System')
    .addItem('Reset Application', 'showManualReset')
    .addItem('Archive Application', 'showArchive')
    .addItem('Notify New Assignments', 'notifyNewFinanceOfficerAssignments')
    .addToUi(); // Add to UI
}

/**
 * Shows a box to reset a loan, like pressing a reset button.
 */
function showManualReset() {
  const ui = SpreadsheetApp.getUi(); // Open a little window
  const response = ui.prompt('Reset Application', 'Enter GroupID:', ui.ButtonSet.OK_CANCEL); // Ask for ID
  if (response.getSelectedButton() === ui.Button.OK) {
    const groupID = response.getResponseText().trim(); // Get the ID
    if (!groupID) {
      ui.alert('Invalid ID. Enter a GroupID.');
      return;
    }
    const result = manualResetUnlock(groupID); // Reset the group
    ui.alert(result); // Show what happened
  }
}

/**
 * Shows a box to archive a loan, like putting away a toy.
 */
function showArchive() {
  const ui = SpreadsheetApp.getUi(); // Open a window
  const response = ui.prompt('Archive Application', 'Enter GroupID:', ui.ButtonSet.OK_CANCEL); // Ask for ID
  if (response.getSelectedButton() === ui.Button.OK) {
    const groupID = response.getResponseText().trim(); // Get the ID
    if (!groupID) {
      ui.alert('Invalid ID. Enter a GroupID.');
      return;
    }
    const result = archiveApplication(groupID); // Archive the group
    ui.alert(result); // Show what happened
  }
}

/**
 * Sets up the toy’s automatic actions, like winding it up.
 */
function setupTriggers() {
  try {
    const ss = SpreadsheetApp.getActive(); // Open notebook
    // Clear existing triggers to avoid duplicates
    ScriptApp.getProjectTriggers().forEach(trigger => ScriptApp.deleteTrigger(trigger));

    // Trigger for Application Form submissions
    ScriptApp.newTrigger('onApplicationFormSubmit')
      .forSpreadsheet(ss)
      .onFormSubmit()
      .create();
    // Trigger for Intent Form submissions
    ScriptApp.newTrigger('onIntentFormSubmit')
      .forForm(LOAN_INTENT_FORM_ID)
      .onFormSubmit()
      .create();
    // Daily reminders
    ScriptApp.newTrigger('sendDailyReminders')
      .timeBased()
      .everyDays(1)
      .atHour(8)
      .create();
    // Update dropdowns when Intent changes
    ScriptApp.newTrigger('syncParticipantDetails')
      .forForm(LOAN_INTENT_FORM_ID)
      .onFormSubmit()
      .create();
    Logger.log('Triggers set up successfully!');
  } catch (e) {
    Logger.log(`Error setting up triggers: ${e.message}`);
  }
}

/**
 * Initializes the system on spreadsheet open.
 */
function onOpen() {
  createMenu();
}