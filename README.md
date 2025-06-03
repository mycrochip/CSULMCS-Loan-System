# CSULMCS Loan Management System

Welcome to the **CSULMCS Loan Management System**! This is like a magical toy box that helps manage loan applications for the CSULMCS community. Built with Google Sheets and Google Apps Script, it automates loan intent submissions, application processing, Finance Officer assignments, and notifications, making the process as fun as playing with your favorite toys!

## Table of Contents

- Overview
- Features
- Prerequisites
- Setup Instructions
- Usage
- File Structure
- Troubleshooting
- Contributing
- License
- Contact

## Overview

The CSULMCS Loan System streamlines loan applications by allowing applicants to sign up with two guarantors via an Intent Form, pend applications until a Finance Officer is assigned, and process applications through an Application Form. Admins assign Finance Officers per loan group, trigger notifications manually, and manage applications via a Google Sheet. A 7-day guarantor reminder countdown starts only after a Finance Officer is assigned, ensuring flexibility and control.

This system uses:

- **Google Forms** for data collection (Intent and Application Forms).
- **Google Sheets** for data storage and tracking (Intent, Control, Archive tabs).
- **Google Apps Script** for automation, email notifications, and custom menu actions.

## Features

- **Finance Officer Assignment**: Assign a unique Finance Officer (Name, ID, Email, Phone) per loan group in the Control sheet, allowing applications to pend until funds are available.
- **7-Day Guarantor Countdown**: Reminders for guarantors start only after a Finance Officer is assigned, and are sent daily for 7 days.
- **Manual Notification Trigger**: Custom menu option ("Notify New Assignments") sends emails only to groups with newly assigned Finance Officers, skipping active application flows (e.g., ApplicantSubmitted, FinanceReviewed, Expired).
- **Intent Form**: Applicants sign up with two guarantors, creating a unique GroupID (e.g., LC0001). One email is sent to the applicant with the GroupID.
- **Application Form**: Supports submissions by Applicants, Guarantors, and Finance Officers, with prefilled links for ease.
- **Control Sheet**: Tracks loan details, including Finance Officer info, with statuses (PendingFinanceOfficer, Notified, ApplicantSubmitted, FinanceReviewed, Expired).
- **Intent Sheet**: Stores one row per participant (applicant + guarantors), unique by GroupID + CooperatorID.
- **Archive Sheet**: Stores completed or expired applications.
- **Robustness**: Prevents duplicate submissions, validates emails, locks completed applications, and supports multiple pending groups.
- **Custom Menu**: Admins can reset, archive, or notify groups via a "CSULMCS Loan System" menu in the Sheet.

## Prerequisites

- **Google Account** with access to Google Sheets, Forms, and Apps Script.
- **Basic Spreadsheet Knowledge** to set up tabs and columns.
- **Admin Access** to edit the Google Sheet and run scripts.
- **Form IDs** for the Intent and Application Forms (obtained from form URLs).

## Setup Instructions

Follow these steps to set up the system from scratch, like building a toy from a kit:

### 1. Create Google Forms

Create two Google Forms and link them to a single Google Sheet.

#### Intent Form

- **Purpose**: Collects applicant and guarantor sign-up details.
- **Fields**:
  - Your Cooperator ID (Text)
  - Your Name (Text)
  - Your Phone (Text)
  - Your Email (Text)
  - Guarantor 1 Cooperator ID (Text)
  - Guarantor 1 Name (Text)
  - Guarantor 1 Phone (Text)
  - Guarantor 1 Email (Text)
  - Guarantor 2 Cooperator ID (Text)
  - Guarantor 2 Name (Text)
  - Guarantor 2 Phone (Text)
  - Guarantor 2 Email (Text)
- **Steps**:
  1. Create a new Google Form (`forms.google.com`).
  2. Add the fields above.
  3. Go to Responses &gt; Select destination &gt; Create a new spreadsheet (name it, e.g., "CSULMCS Loan System").
  4. Note the Form ID from the URL: `https://docs.google.com/forms/d/[FORM_ID]/edit`.

#### Application Form

- **Purpose**: Collects detailed loan application data from Applicants, Guarantors, and Finance Officers.
- **Fields**:
  - Loan ID (Text)
  - Role (Dropdown: Applicant, Guarantor1, Guarantor2, Finance Officer)
  - Cooperator ID (Dropdown, populated by script)
  - Name (Dropdown)
  - Phone (Dropdown)
  - Email (Dropdown)
  - Home Address (Text)
  - Loan Amount (Figures) (Text)
  - Loan Amount (Words) (Text)
  - Repayment Period (Text)
  - Bank Name (Text)
  - Account Name (Text)
  - Account Number (Text)
  - Guarantor 1 Name (Dropdown)
  - Guarantor 1 Cooperator ID (Dropdown)
  - Guarantor 1 Email (Dropdown)
  - Guarantor 1 Phone (Dropdown)
  - Guarantor 2 Name (Dropdown)
  - Guarantor 2 Cooperator ID (Dropdown)
  - Guarantor 2 Email (Dropdown)
  - Guarantor 2 Phone (Dropdown)
  - Approver Name (Dropdown)
  - Approver ID (Dropdown)
  - Approver Email (Dropdown)
  - Approver Phone (Dropdown)
  - Status (Dropdown: Approved, Denied)
  - Applicant Balance (Text)
  - Applicant Rating (Text)
  - Guarantor 1 Balance (Text)
  - Guarantor 1 Rating (Text)
  - Guarantor 2 Balance (Text)
  - Guarantor 2 Rating (Text)
  - Comments (Paragraph)
- **Steps**:
  1. Create a new Google Form.
  2. Add the fields above, ensuring dropdowns for Role and Status.
  3. Go to Responses &gt; Select destination &gt; Select existing spreadsheet (choose the same Sheet as Intent Form).
  4. Note the Form ID from the URL.

### 2. Set Up Google Sheet

Configure the Google Sheet to store and track loan data.

- **Tabs**:
  - **Intent**: Auto-created by Intent Form. Columns (A:G):
    - Timestamp, GroupID, CooperatorID, Name, Phone, Email, Role
  - **Control**: Create manually. Columns (A:AP):
    - GroupID, CooperatorID, Name, Email, Phone, HomeAddress, LoanAmountFigures, LoanAmountWords, RepaymentPeriod, Guarantor1Name, Guarantor1ID, Guarantor1Email, Guarantor1Phone, Guarantor2Name, Guarantor2ID, Guarantor2Email, Guarantor2Phone, ApproverName, ApproverID, ApproverEmail, ApproverPhone, Status, ApplicantLink, FinanceLink, ApplicationStatus, Locked, Timestamp, Comments, BankName, AccountName, AccountNumber, ApplicantBalance, ApplicantRating, Guarantor1Balance, Guarantor1Rating, Guarantor2Balance, Guarantor2Rating, FinanceOfficerName, FinanceOfficerID, FinanceOfficerEmail, FinanceOfficerPhone
  - **Archive**: Auto-created by script, same columns as Control.
- **Steps**:
  1. Open the Google Sheet linked to the forms.
  2. Rename the Intent Form response tab to "Intent" if needed.
  3. Create a "Control" tab and add the 40 column headers above (A1:AP1).
  4. The Archive tab will be created by the script when needed.

### 3. Install the Script

Add the automation script to the Google Sheet.

- **Steps**:
  1. Clone or download this repository: `https://github.com/mycrochip/CSULMCS-Loan-System`.
  2. Open the Google Sheet, go to `Extensions > Apps Script`.
  3. Delete any default code in the editor.
  4. Copy-paste the script from `CSULMCS_Loan_System_Script.js` in this repository.
  5. Replace the placeholder Form IDs:
     - Update `LOAN_APPLICATION_FORM_ID` with the Application Form ID.
     - Update `LOAN_INTENT_FORM_ID` with the Intent Form ID.
  6. Save the script (File &gt; Save, name it "CSULMCS Loan System").

### 4. Run Initial Setup

Set up triggers and the custom menu.

- **Steps**:
  1. In Apps Script, select the `setupTriggers` function from the dropdown.
  2. Click `Run` and authorize permissions (Google will prompt for spreadsheet, form, and email access).
  3. This creates triggers for:
     - Intent Form submissions (`onIntentFormSubmit`).
     - Application Form submissions (`onApplicationFormSubmit`).
     - Daily reminders at 8 AM (`sendDailyReminders`).
     - Dropdown updates (`syncParticipantDetails`).
  4. Select and run the `createMenu` function to add the "CSULMCS Loan System" menu to the Sheet.
  5. Refresh the Sheet to see the menu.

### 5. Verify Setup

- Submit a test Intent Form with your email for the applicant and guarantors.
- Check the Intent tab for three rows (Applicant, Guarantor1, Guarantor2) with a GroupID (e.g., LC0001).
- Check the Control tab for a placeholder row (GroupID, ApplicationStatus: PendingFinanceOfficer).
- Ensure no errors appear in Apps Script’s Execution Log (View &gt; Logs).

## Usage

Here’s how admins and users interact with the system, like playing with a well-built toy.

### Admin Tasks

1. **Assign Finance Officers**:
   - After Intent Form submissions, find GroupIDs in the Control sheet (Column A).
   - Enter Finance Officer details in columns AM:AP (FinanceOfficerName, FinanceOfficerID, FinanceOfficerEmail, FinanceOfficerPhone).
   - Leave blank to pend applications (e.g., until funds are available).
2. **Send Notifications**:
   - Go to Sheet &gt; `CSULMCS Loan System` &gt; `Notify New Assignments`.
   - Confirm the prompt to submit emails to Applicants, Guarantors, and Finance Officers for groups with newly assigned Finance Officers (status not ApplicantSubmitted, Notified, FinanceReviewed, or Expired).
3. **Manage Applications**:
   - Monitor Control sheet’s ApplicationStatus (Column Y):
     - **PendingFinanceOfficer**: Awaiting Finance Officer assignment.
     - **Notified**: Finance Officer assigned, notifications sent.
     - **ApplicantSubmitted**: Applicant submitted Application Form.
     - **FinanceReviewed**: Finance Officer reviewed (Approved/Denied).
     - **Expired**: Guarantors didn’t respond within 7 days.
   - Use `Reset Application` to unlock a GroupID for resubmission (clears status, sends emails).
   - Use `Archive Application` to move completed/expired groups to the Archive tab.
4. **Pending Applications**:
   - Applications remain in “PendingFinanceOfficer” until a Finance Officer is assigned, preventing the 7-day countdown.

### User Flow

1. **Applicant**:
   - Submits Intent Form with two guarantors.
   - Receives an email with GroupID and instructions to wait for the Finance Officer assignment.
   - After notification, submits the Application Form using a prefilled link.
2. **Guarantors**:
   - Receive notification emails with prefilled links to submit details after the Finance Officer assignment.
   - Submit Application Form within 7 days, receiving daily reminders.
3. **Finance Officer**:
   - Receives notification of assignment and application submissions.
   - Reviews Application Form, sets status (Approved/Denied), and submits.

### Testing

- **Test Intent Form**:
  - Submit with test data (use your email for all roles).
  - Verify Intent and Control sheet updates.
- **Test Notifications**:
  - Assign a Finance Officer in Control (columns AM:AP).
  - Run `Notify New Assignments` and check emails.
- **Test Application Form**:
  - Submit as Applicant, Guarantors, and Finance Officer using prefilled links.
  - Verify sheet updates and email notifications.
- **Test Reminders**:
  - Wait for 8 AM daily trigger or run `sendDailyReminders` manually.
  - Check reminder emails for guarantors and the Finance Officer.
- **Test Reset/Archive**:
  - Use menu options to reset or archive a GroupID and verify results.

## File Structure

```
CSULMCS-Loan-System/
├── CSULMCS_Loan_System_Script.js
  # Main Google Apps Script
├── LICENSE
│
  # MIT License
├── README.md
# Project documentation
└── [Google Sheet]
    ├── Intent
    │
    # Auto-created by Intent Form
    ├── Control
    │
    # Manually created
    └── Archive
        └── Auto-created by script
```

## Troubleshooting

- **Script Errors**:
  - Check Apps Script’s Execution Log (View &gt; Logs) for details.
  - Ensure Form IDs in the script match your forms.
  - Verify all Control sheet columns (A:AP) are present.
- **No Emails Sent**:
  - Confirm email addresses are valid (script skips invalid ones).
  - Check Google’s email quota (typically 100/day for free accounts).
- **Triggers Not Working**:
  - Run `setupTriggers` again to recreate triggers.
  - Ensure permissions are authorized.
- **Dropdowns Not Populating**:
  - Submit an Intent Form to populate the Intent sheet data.
  - Run `syncParticipantDetails` manually if needed.
- **Locked Application**:
  - Use `Reset Application` menu option to unlock a GroupID.
- **Report Issues**:
  - Open an issue at https://github.com/mycrochip/CSULMCS-Loan-System/issues.

## Contributing

We welcome contributions to make this toy box even more fun! To contribute:

1. Fork or clone the repository: `https://github.com/mycrochip/CSULMCS-Loan-System`.
2. Create a feature branch (`git checkout -b feature/YourFeature`).
3. Commit changes (`git commit -m 'Add YourFeature'`).
4. Push to the branch (`git push origin feature/YourFeature`).
5. Submit a pull request with a clear description.

Please ensure the code follows the existing style, includes comments using the toy-box analogy, and tests all changes. See CONTRIBUTING.md for details (create one if needed).

## License

This project is licensed under the MIT License. See the LICENSE file for details.

## Contact

For support or questions, contact the maintainer:

- **GitHub**: @mycrochip
- **Issues**: https://github.com/mycrochip/CSULMCS-Loan-System/issues

---

*Built with ❤️* by the CSULMCS team. Last updated: June 3, 2025.\*
