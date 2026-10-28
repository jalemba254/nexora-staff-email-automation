/****************************************************
 * STAFF EMAIL AUTOMATION SYSTEM
 * Google Sheets + Gmail + Apps Script
 * Version 2.0
 ****************************************************/

const SHEETS = {
  STAFF: 'STAFF',
  TEMPLATES: 'EMAIL_TEMPLATES',
  LOG: 'EMAIL_LOG'
};


/****************************************************
 * MENU
 ****************************************************/

/****************************************************
 * MAIN MENU
 ****************************************************/

function onOpen() {

  SpreadsheetApp
    .getUi()
    .createMenu('📧 Email Automation')
    
.addItem(
  '📨 Email Center',
  'showEmailCenter'
)
    .addItem(
      '📊 Automation Dashboard',
      'showAutomationDashboard'
    )

.addItem(
  '⚙️ Automation Rules',
  'showAutomationRules'
)

    .addSeparator()

    .addItem(
      '▶ Run Automation Now',
      'runAutomationNow'
    )

    .addItem(
      '⚙️ Setup / Repair System',
      'setupSystem'
    )

    .addItem(
      '🕐 Create Hourly Automation',
      'createHourlyTrigger'
    )

    .addToUi();

}


/****************************************************
 * OPEN AUTOMATION DASHBOARD
 ****************************************************/

function showAutomationDashboard() {

  const html =
    HtmlService
      .createHtmlOutputFromFile(
        'AutomationDashboard'
      )
      .setWidth(1200)
      .setHeight(750);

  SpreadsheetApp
    .getUi()
    .showModalDialog(
      html,
      'Automation Dashboard'
    );

}

/****************************************************
 * OPEN SIDEBAR
 ****************************************************/

function showEmailSidebar() {
  const html = HtmlService
    .createHtmlOutputFromFile('Sidebar')
    .setTitle('Email Automation Center');

  SpreadsheetApp.getUi().showSidebar(html);
}


/****************************************************
 * SETUP SYSTEM
 ****************************************************/

function setupSystem() {
  setupHeaders_();

  SpreadsheetApp.getUi().alert(
    'System Setup Complete',
    'STAFF, EMAIL_TEMPLATES and EMAIL_LOG sheets are ready.',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}


/****************************************************
 * SETUP HEADERS
 ****************************************************/

/****************************************************
 * SETUP HEADERS
 ****************************************************/

function setupHeaders_() {

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const staffHeaders = [
    'Staff ID',
    'Full Name',
    'Email',
    'Department',
    'Status',
    'Status Updated At'
  ];

  const templateHeaders = [
    'Template ID',
    'Template Name',
    'Subject',
    'Message',
    'Status'
  ];

  const logHeaders = [
    'Log ID',
    'Date',
    'Staff ID',
    'Staff Name',
    'Email',
    'Template',
    'Subject',
    'Status',
    'Sent By',
    'Message ID',
    'Automation Key'
  ];

  createSheetWithHeaders_(
    ss,
    SHEETS.STAFF,
    staffHeaders
  );

  createSheetWithHeaders_(
    ss,
    SHEETS.TEMPLATES,
    templateHeaders
  );

  createSheetWithHeaders_(
    ss,
    SHEETS.LOG,
    logHeaders
  );

  // Ensure existing STAFF sheet has Status Updated At column
  const staffSheet = ss.getSheetByName(SHEETS.STAFF);

  if (staffSheet) {
    const lastColumn = staffSheet.getLastColumn();

    if (lastColumn < 6) {
      staffSheet.insertColumnAfter(lastColumn);
    }

    if (
      String(staffSheet.getRange(1, 6).getValue()).trim() === ''
    ) {
      staffSheet
        .getRange(1, 6)
        .setValue('Status Updated At');
    }
  }

  // Ensure existing EMAIL_LOG sheet has Automation Key column
  const logSheet = ss.getSheetByName(SHEETS.LOG);

  if (logSheet) {
    const lastColumn = logSheet.getLastColumn();

    if (lastColumn < 11) {
      logSheet.insertColumnAfter(lastColumn);
    }

    if (
      String(logSheet.getRange(1, 11).getValue()).trim() === ''
    ) {
      logSheet
        .getRange(1, 11)
        .setValue('Automation Key');
    }
  }
}


/****************************************************
 * CREATE SHEET + HEADERS
 ****************************************************/

function createSheetWithHeaders_(ss, sheetName, headers) {

  let sheet = ss.getSheetByName(sheetName);

  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
  }

  if (sheet.getLastRow() === 0) {
    sheet.appendRow(headers);
  }
}


/****************************************************
 * GET STAFF
 ****************************************************/

function getStaff() {

  const sheet =
    SpreadsheetApp
      .getActiveSpreadsheet()
      .getSheetByName(SHEETS.STAFF);

  if (!sheet || sheet.getLastRow() < 2) {
    return [];
  }

  const values = sheet
    .getRange(
      2,
      1,
      sheet.getLastRow() - 1,
      5
    )
    .getValues();

  return values
    .filter(row => row[0])
    .map(row => ({
      id: String(row[0]),
      name: String(row[1] || ''),
      email: String(row[2] || ''),
      department: String(row[3] || ''),
      status: String(row[4] || '')
    }));
}


/****************************************************
 * GET TEMPLATES
 ****************************************************/

function getTemplates() {

  const sheet =
    SpreadsheetApp
      .getActiveSpreadsheet()
      .getSheetByName(SHEETS.TEMPLATES);

  if (!sheet || sheet.getLastRow() < 2) {
    return [];
  }

  const values = sheet
    .getRange(
      2,
      1,
      sheet.getLastRow() - 1,
      5
    )
    .getValues();

  return values
    .filter(row => row[0])
    .map(row => ({
      id: String(row[0]),
      name: String(row[1] || ''),
      subject: String(row[2] || ''),
      message: String(row[3] || ''),
      status: String(row[4] || '')
    }));
}


/****************************************************
 * GET STAFF BY ID
 ****************************************************/

function getStaffById(staffId) {

  const staff = getStaff();

  return staff.find(
    person => person.id === String(staffId)
  ) || null;
}


/****************************************************
 * GET TEMPLATE BY ID
 ****************************************************/

function getTemplateById(templateId) {

  const templates = getTemplates();

  return templates.find(
    template => template.id === String(templateId)
  ) || null;
}


/****************************************************
 * REPLACE VARIABLES
 ****************************************************/

function replaceVariables_(text, staff) {

  if (!text) {
    return '';
  }

  return String(text)
    .replace(/\{\{name\}\}/gi, staff.name)
    .replace(/\{\{email\}\}/gi, staff.email)
    .replace(/\{\{department\}\}/gi, staff.department);
}


/****************************************************
 * DUPLICATE PROTECTION
 ****************************************************/

function hasDuplicateEmail_(staffId, templateId) {

  const sheet =
    SpreadsheetApp
      .getActiveSpreadsheet()
      .getSheetByName(SHEETS.LOG);

  if (!sheet || sheet.getLastRow() < 2) {
    return false;
  }

  const values = sheet
    .getRange(
      2,
      1,
      sheet.getLastRow() - 1,
      10
    )
    .getValues();

  return values.some(row => {

    const loggedStaffId = String(row[2] || '');
    const loggedTemplate = String(row[5] || '');
    const status = String(row[7] || '');

    return (
      loggedStaffId === String(staffId) &&
      loggedTemplate === String(templateId) &&
      status === 'Sent'
    );

  });
}


/****************************************************
 * SEND EMAIL
 ****************************************************/

function sendEmail(data) {

  try {

    if (!data) {
      throw new Error('No email data was provided.');
    }

    if (!data.staffId) {
      throw new Error('Please select a staff member.');
    }

    if (!data.subject) {
      throw new Error('Email subject is required.');
    }

    if (!data.message) {
      throw new Error('Email message is required.');
    }

    const staff = getStaffById(data.staffId);

    if (!staff) {
      throw new Error('Staff member not found.');
    }

    if (!staff.email) {
      throw new Error(
        'The selected staff member does not have an email address.'
      );
    }

    const templateId =
      data.templateId || 'MANUAL';

    if (
      templateId !== 'MANUAL' &&
      hasDuplicateEmail_(
        data.staffId,
        templateId
      )
    ) {

      return {
        success: false,
        duplicate: true,
        message:
          'This email has already been sent to this staff member using this template.'
      };
    }

    const subject =
      replaceVariables_(
        data.subject,
        staff
      );

    const message =
      replaceVariables_(
        data.message,
        staff
      );

    const options = {};

    if (data.attachmentUrl) {

      const fileId =
        extractDriveFileId_(
          data.attachmentUrl
        );

      if (!fileId) {
        throw new Error(
          'Invalid Google Drive file URL.'
        );
      }

      const file =
        DriveApp.getFileById(fileId);

      options.attachments = [
        file.getBlob()
      ];
    }

    const result =
      GmailApp.sendEmail(
        staff.email,
        subject,
        message,
        options
      );

    logEmail_({
      staffId: staff.id,
      staffName: staff.name,
      email: staff.email,
      template:
        data.templateName ||
        templateId,
      subject: subject,
      status: 'Sent',
      sentBy:
        Session.getActiveUser().getEmail(),
      messageId: result || ''
    });

    return {
      success: true,
      message:
        'Email sent successfully to ' +
        staff.name
    };

  } catch (error) {

    try {

      const staff =
        data &&
        data.staffId
          ? getStaffById(data.staffId)
          : null;

      logEmail_({
        staffId:
          staff ? staff.id : '',
        staffName:
          staff ? staff.name : '',
        email:
          staff ? staff.email : '',
        template:
          data
            ? data.templateName || data.templateId || 'Manual'
            : 'Manual',
        subject:
          data
            ? data.subject || ''
            : '',
        status: 'Failed',
        sentBy:
          Session.getActiveUser().getEmail(),
        messageId:
          error.message
      });

    } catch (logError) {
      console.log(logError);
    }

    return {
      success: false,
      message: error.message
    };
  }
}


/****************************************************
 * EXTRACT DRIVE FILE ID
 ****************************************************/

function extractDriveFileId_(url) {

  if (!url) {
    return null;
  }

  const match =
    String(url).match(
      /[-\w]{25,}/
    );

  return match
    ? match[0]
    : null;
}


/****************************************************
 * LOG EMAIL
 ****************************************************/

function logEmail_(data) {

  const sheet =
    SpreadsheetApp
      .getActiveSpreadsheet()
      .getSheetByName(SHEETS.LOG);

  if (!sheet) {
    throw new Error(
      'EMAIL_LOG sheet not found.'
    );
  }

  const logId =
    'LOG-' +
    Utilities.getUuid()
      .substring(0, 8)
      .toUpperCase();

  sheet.appendRow([
    logId,
    new Date(),
    data.staffId || '',
    data.staffName || '',
    data.email || '',
    data.template || '',
    data.subject || '',
    data.status || '',
    data.sentBy || '',
    data.messageId || ''
  ]);
}


/****************************************************
 * GET EMAIL LOG
 ****************************************************/

function getEmailLog() {

  const sheet =
    SpreadsheetApp
      .getActiveSpreadsheet()
      .getSheetByName(SHEETS.LOG);

  if (!sheet || sheet.getLastRow() < 2) {
    return [];
  }

  const values =
    sheet
      .getRange(
        2,
        1,
        sheet.getLastRow() - 1,
        10
      )
      .getValues();

  return values
    .reverse()
    .map(row => ({
      id: String(row[0] || ''),
      date: row[1],
      staffId: String(row[2] || ''),
      staffName: String(row[3] || ''),
      email: String(row[4] || ''),
      template: String(row[5] || ''),
      subject: String(row[6] || ''),
      status: String(row[7] || ''),
      sentBy: String(row[8] || ''),
      messageId: String(row[9] || '')
    }));
}


/****************************************************
 * EMAIL STATISTICS
 ****************************************************/

function getEmailStats() {

  const log =
    getEmailLog();

  const today =
    new Date();

  const todayString =
    Utilities.formatDate(
      today,
      Session.getScriptTimeZone(),
      'yyyy-MM-dd'
    );

  let sent = 0;
  let todaySent = 0;
  let failed = 0;

  log.forEach(item => {

    if (item.status === 'Sent') {
      sent++;
    }

    if (item.status === 'Failed') {
      failed++;
    }

    if (
      item.date &&
      Utilities.formatDate(
        new Date(item.date),
        Session.getScriptTimeZone(),
        'yyyy-MM-dd'
      ) === todayString &&
      item.status === 'Sent'
    ) {
      todaySent++;
    }

  });

  return {
    sent: sent,
    today: todaySent,
    failed: failed
  };
}


/****************************************************
 * CREATE TEMPLATE
 ****************************************************/

function createTemplate(data) {

  if (!data) {
    throw new Error(
      'Template data is missing.'
    );
  }

  if (!data.name) {
    throw new Error(
      'Template name is required.'
    );
  }

  if (!data.subject) {
    throw new Error(
      'Template subject is required.'
    );
  }

  if (!data.message) {
    throw new Error(
      'Template message is required.'
    );
  }

  const sheet =
    SpreadsheetApp
      .getActiveSpreadsheet()
      .getSheetByName(
        SHEETS.TEMPLATES
      );

  if (!sheet) {
    throw new Error(
      'EMAIL_TEMPLATES sheet not found.'
    );
  }

  const templateId =
    'TMP-' +
    Utilities.getUuid()
      .substring(0, 8)
      .toUpperCase();

  sheet.appendRow([
    templateId,
    data.name,
    data.subject,
    data.message,
    data.status || 'Active'
  ]);

  return {
    success: true,
    id: templateId
  };
}


/****************************************************
 * UPDATE TEMPLATE
 ****************************************************/

function updateTemplate(data) {

  if (!data || !data.id) {
    throw new Error(
      'Template ID is required.'
    );
  }

  const sheet =
    SpreadsheetApp
      .getActiveSpreadsheet()
      .getSheetByName(
        SHEETS.TEMPLATES
      );

  if (!sheet) {
    throw new Error(
      'EMAIL_TEMPLATES sheet not found.'
    );
  }

  const values =
    sheet.getRange(
      2,
      1,
      Math.max(
        sheet.getLastRow() - 1,
        1
      ),
      5
    ).getValues();

  for (
    let i = 0;
    i < values.length;
    i++
  ) {

    if (
      String(values[i][0]) ===
      String(data.id)
    ) {

      sheet
        .getRange(i + 2, 1, 1, 5)
        .setValues([[
          data.id,
          data.name,
          data.subject,
          data.message,
          data.status || 'Active'
        ]]);

      return {
        success: true
      };
    }
  }

  throw new Error(
    'Template not found.'
  );
}


/****************************************************
 * ARCHIVE TEMPLATE
 ****************************************************/

function archiveTemplate(templateId) {

  const sheet =
    SpreadsheetApp
      .getActiveSpreadsheet()
      .getSheetByName(
        SHEETS.TEMPLATES
      );

  if (!sheet) {
    throw new Error(
      'EMAIL_TEMPLATES sheet not found.'
    );
  }

  const values =
    sheet.getRange(
      2,
      1,
      Math.max(
        sheet.getLastRow() - 1,
        1
      ),
      5
    ).getValues();

  for (
    let i = 0;
    i < values.length;
    i++
  ) {

    if (
      String(values[i][0]) ===
      String(templateId)
    ) {

      sheet
        .getRange(i + 2, 5)
        .setValue('Archived');

      return {
        success: true
      };
    }
  }

  throw new Error(
    'Template not found.'
  );
}


/****************************************************
 * CREATE HOURLY TRIGGER
 ****************************************************/

function createHourlyTrigger() {

  const triggers =
    ScriptApp.getProjectTriggers();

  triggers.forEach(trigger => {

    if (
      trigger
        .getHandlerFunction() ===
      'processAutomaticEmails'
    ) {
      ScriptApp.deleteTrigger(
        trigger
      );
    }

  });

  ScriptApp.newTrigger(
    'processAutomaticEmails'
  )
    .timeBased()
    .everyHours(1)
    .create();

  SpreadsheetApp.getUi().alert(
    'Automation Created',
    'The system will check for automatic emails every hour.',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}


/****************************************************
 * AUTOMATIC EMAIL PROCESSOR
 ****************************************************/

/****************************************************
 * PROCESS AUTOMATIC EMAILS
 * Version 2.1 - Delay Enforcement
 ****************************************************/

function processAutomaticEmails() {

  const ss =
    SpreadsheetApp.getActiveSpreadsheet();

  const staffSheet =
    ss.getSheetByName(SHEETS.STAFF);

  const rulesSheet =
    ss.getSheetByName(AUTOMATION_SHEET);

  const logSheet =
    ss.getSheetByName(SHEETS.LOG);


  if (!staffSheet) {
    console.log('STAFF sheet not found.');
    return;
  }

  if (!rulesSheet) {
    console.log('AUTOMATION_RULES sheet not found.');
    return;
  }

  if (!logSheet) {
    console.log('EMAIL_LOG sheet not found.');
    return;
  }


  const staff =
    getStaff();

  const templates =
    getTemplates();

  const rules =
    getAutomationRules();


  if (!staff.length) {
    console.log('No staff records found.');
    return;
  }

  if (!rules.length) {
    console.log('No automation rules found.');
    return;
  }


  const now =
    new Date();


  rules
    .filter(rule => rule.active)
    .forEach(rule => {

      const matchingStaff =
        staff.filter(person => {

          return (
            String(person.status)
              .trim()
              .toLowerCase() ===

            String(rule.staffStatus)
              .trim()
              .toLowerCase()
          );

        });


      matchingStaff.forEach(person => {

        const template =
          templates.find(
            item =>
              item.id ===
              rule.templateId
          );


        if (!template) {

          console.log(
            'Template not found: ' +
            rule.templateId
          );

          return;
        }


        if (
          template.status ===
          'Archived'
        ) {

          return;
        }


        /*
         * Automation key uniquely identifies
         * this staff member + rule.
         */

        const automationKey =
          person.id +
          '|' +
          rule.id;


        /*
         * Check whether this exact
         * automation has already been sent.
         */

        if (
          hasAutomationBeenSent_(
            automationKey
          )
        ) {

          console.log(
            'Automation already sent: ' +
            automationKey
          );

          return;
        }


        /*
         * Determine when this email
         * became eligible.
         *
         * For the first implementation,
         * we use the current time as the
         * starting point unless the staff
         * status has already been logged.
         */

        const eligibleTime =
          getAutomationEligibleTime_(
            person.id,
            rule.id,
            rule.delayHours
          );


        if (
          now.getTime() <
          eligibleTime.getTime()
        ) {

          console.log(
            'Waiting for delay: ' +
            person.name +
            ' / ' +
            rule.name
          );

          return;
        }


        const subject =
          replaceVariables_(
            template.subject,
            person
          );


        const message =
          replaceVariables_(
            template.message,
            person
          );


        try {

          GmailApp.sendEmail(
            person.email,
            subject,
            message
          );


          logAutomationEmail_({

            automationKey:
              automationKey,

            staffId:
              person.id,

            staffName:
              person.name,

            email:
              person.email,

            template:
              template.name,

            subject:
              subject,

            status:
              'Sent',

            sentBy:
              'AUTOMATION',

            messageId:
              ''

          });


          console.log(
            'Automatic email sent to ' +
            person.email
          );


        } catch (error) {

          logAutomationEmail_({

            automationKey:
              automationKey,

            staffId:
              person.id,

            staffName:
              person.name,

            email:
              person.email,

            template:
              template.name,

            subject:
              subject,

            status:
              'Failed',

            sentBy:
              'AUTOMATION',

            messageId:
              error.message

          });


          console.log(
            'Automatic email failed: ' +
            error.message
          );

        }

      });

    });

}


/****************************************************
 * CHECK AUTOMATION SEND HISTORY
 ****************************************************/

/****************************************************
 * CHECK AUTOMATION DUPLICATES
 * Uses Automation Key AND legacy log data
 ****************************************************/

function hasAutomationBeenSent_(automationKey) {

  const sheet =
    SpreadsheetApp
      .getActiveSpreadsheet()
      .getSheetByName(SHEETS.LOG);

  if (!sheet || sheet.getLastRow() < 2) {
    return false;
  }

  const values =
    sheet
      .getRange(
        2,
        1,
        sheet.getLastRow() - 1,
        Math.max(sheet.getLastColumn(), 11)
      )
      .getValues();

  /*
   * automationKey format:
   * STF-003|RULE-001
   */

  const parts =
    String(automationKey).split('|');

  const staffId =
    parts[0] || '';

  const ruleId =
    parts[1] || '';

  /*
   * Find the rule so we can identify
   * its template.
   */

  const rules =
    getAutomationRules();

  const rule =
    rules.find(
      r => String(r.id) === String(ruleId)
    );

  const templateId =
    rule ? String(rule.templateId) : '';


  return values.some(row => {

    const status =
      String(row[7] || '').trim();

    if (status.toLowerCase() !== 'sent') {
      return false;
    }


    /*
     * 1. New Automation Key check
     */

    const storedKey =
      String(row[10] || '').trim();

    if (
      storedKey ===
      String(automationKey)
    ) {
      return true;
    }


    /*
     * 2. Legacy log protection
     *
     * This catches automated emails that
     * were sent before Automation Key existed.
     */

    const loggedStaffId =
      String(row[2] || '').trim();

    const loggedTemplate =
      String(row[5] || '').trim();


    if (
      loggedStaffId === staffId &&
      loggedTemplate === templateId &&
      String(row[8] || '').trim() ===
        'AUTOMATION'
    ) {
      return true;
    }


    return false;

  });

}

/****************************************************
 * GET AUTOMATION ELIGIBLE TIME
 ****************************************************/

/****************************************************
 * GET AUTOMATION ELIGIBLE TIME
 * Uses STAFF > Status Updated At
 ****************************************************/

function getAutomationEligibleTime_(
  staffId,
  ruleId,
  delayHours
) {

  const sheet =
    SpreadsheetApp
      .getActiveSpreadsheet()
      .getSheetByName(SHEETS.STAFF);

  if (!sheet) {
    return new Date();
  }

  const data =
    sheet.getDataRange().getValues();

  if (data.length < 2) {
    return new Date();
  }

  /*
   * STAFF columns:
   * A = Staff ID
   * B = Full Name
   * C = Email
   * D = Department
   * E = Status
   * F = Status Updated At
   */

  for (let i = 1; i < data.length; i++) {

    const currentStaffId =
      String(data[i][0] || '').trim();

    if (
      currentStaffId !==
      String(staffId).trim()
    ) {
      continue;
    }

    const statusUpdatedAt =
      data[i][5];

    /*
     * If there is no timestamp,
     * treat the employee as eligible
     * immediately.
     */

    if (!statusUpdatedAt) {
      return new Date();
    }

    const statusDate =
      statusUpdatedAt instanceof Date
        ? statusUpdatedAt
        : new Date(statusUpdatedAt);

    if (isNaN(statusDate.getTime())) {
      return new Date();
    }

    const delay =
      Number(delayHours) || 0;

    return new Date(
      statusDate.getTime() +
      delay *
      60 *
      60 *
      1000
    );
  }

  return new Date();
}

/****************************************************
 * LOG AUTOMATION EMAIL
 ****************************************************/

function logAutomationEmail_(data) {

  const sheet =
    SpreadsheetApp
      .getActiveSpreadsheet()
      .getSheetByName(SHEETS.LOG);


  if (!sheet) {

    throw new Error(
      'EMAIL_LOG sheet not found.'
    );

  }


  /*
   * Make sure column K exists
   * for Automation Key.
   */

  if (
    sheet.getMaxColumns() < 11
  ) {

    sheet.insertColumnAfter(
      sheet.getMaxColumns()
    );

  }


  /*
   * Add header if needed.
   */

  if (
    sheet
      .getRange(1, 11)
      .getValue() === ''
  ) {

    sheet
      .getRange(1, 11)
      .setValue(
        'Automation Key'
      );

  }


  const logId =
    'LOG-' +
    Utilities.getUuid()
      .substring(0, 8)
      .toUpperCase();


  sheet.appendRow([

    logId,

    new Date(),

    data.staffId || '',

    data.staffName || '',

    data.email || '',

    data.template || '',

    data.subject || '',

    data.status || '',

    data.sentBy || '',

    data.messageId || '',

    data.automationKey || ''

  ]);

}

/****************************************************
 * SYSTEM STATUS
 ****************************************************/

function getSystemStatus() {

  const ss =
    SpreadsheetApp
      .getActiveSpreadsheet();

  const logSheet =
    ss.getSheetByName(
      SHEETS.LOG
    );

  return {
    spreadsheet:
      ss.getName(),

    staffCount:
      getStaff().length,

    templateCount:
      getTemplates().length,

    logCount:
      logSheet
        ? Math.max(
            0,
            logSheet.getLastRow() - 1
          )
        : 0
  };
}
/****************************************************
 * AUTOMATION RULES
 ****************************************************/

const AUTOMATION_SHEET = 'AUTOMATION_RULES';


/****************************************************
 * GET AUTOMATION RULES
 ****************************************************/

function getAutomationRules() {

  const sheet =
    SpreadsheetApp
      .getActiveSpreadsheet()
      .getSheetByName(AUTOMATION_SHEET);

  if (!sheet || sheet.getLastRow() < 2) {
    return [];
  }

  const values =
    sheet
      .getRange(
        2,
        1,
        sheet.getLastRow() - 1,
        6
      )
      .getValues();

  return values
    .filter(row => row[0])
    .map(row => ({

      id: String(row[0] || ''),

      name: String(row[1] || ''),

      staffStatus: String(row[2] || ''),

      templateId: String(row[3] || ''),

      delayHours:
        Number(row[4]) || 0,

      active:
        row[5] === true ||
        String(row[5]).toUpperCase() === 'TRUE'

    }));

}
function testAutomationRulesBackend() {

  console.log('====================================');
  console.log('AUTOMATION RULES BACKEND TEST');
  console.log('====================================');

  // TEST 1
  console.log('TEST 1: Starting');

  // TEST 2
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  console.log(
    'TEST 2: Spreadsheet = ' + ss.getName()
  );

  // TEST 3
  const sheet = ss.getSheetByName(AUTOMATION_SHEET);

  if (!sheet) {
    throw new Error(
      'AUTOMATION_RULES sheet not found.'
    );
  }

  console.log(
    'TEST 3: AUTOMATION_RULES sheet found'
  );

  // TEST 4
  const lastRow = sheet.getLastRow();
  const lastColumn = sheet.getLastColumn();

  console.log(
    'TEST 4: Last row = ' + lastRow
  );

  console.log(
    'TEST 4B: Last column = ' + lastColumn
  );

  // TEST 5
  if (lastRow < 1) {

    console.log(
      'TEST 5: Sheet is completely empty.'
    );

    console.log('TEST COMPLETE');
    return;
  }

  // TEST 6
  const data = sheet
    .getRange(
      1,
      1,
      lastRow,
      lastColumn
    )
    .getValues();

  console.log(
    'TEST 6: Data successfully read.'
  );

  // TEST 7
  console.log(
    'TEST 7: Number of rows = ' +
    data.length
  );

  // TEST 8
  console.log(
    'TEST 8: Headers = ' +
    JSON.stringify(data[0])
  );

  // TEST 9
  if (data.length <= 1) {

    console.log(
      'TEST 9: NO AUTOMATION RULES FOUND.'
    );

    console.log(
      'The sheet contains headers only.'
    );

    console.log('TEST COMPLETE');
    return;
  }

  // TEST 10
  console.log(
    'TEST 10: Automation rule rows found = ' +
    (data.length - 1)
  );

  // TEST 11
  for (let i = 1; i < data.length; i++) {

    console.log(
      'RULE ROW ' + (i + 1) + ': ' +
      JSON.stringify(data[i])
    );
  }

  console.log('====================================');
  console.log('TEST COMPLETE');
  console.log('====================================');
}
function testGetAutomationRulesOnly() {

  console.log('====================================');
  console.log('GET AUTOMATION RULES TEST');
  console.log('====================================');

  console.log('TEST 1: Calling getAutomationRules()');

  const rules = getAutomationRules();

  console.log('TEST 2: Function completed');

  console.log(
    'TEST 3: Result type = ' +
    typeof rules
  );

  console.log(
    'TEST 4: Is array = ' +
    Array.isArray(rules)
  );

  console.log(
    'TEST 5: Rule count = ' +
    (rules ? rules.length : 'NULL')
  );

  console.log(
    'TEST 6: Rules = ' +
    JSON.stringify(rules)
  );

  console.log('====================================');
  console.log('TEST COMPLETE');
  console.log('====================================');
}


/****************************************************
 * STAFF STATUS TIMESTAMP
 ****************************************************/

function onEdit(e) {

  if (!e || !e.range) {
    return;
  }


  const sheet =
    e.range.getSheet();


  if (
    sheet.getName() !==
    SHEETS.STAFF
  ) {
    return;
  }


  /*
   * Status = Column E
   * Status Updated At = Column F
   */

  if (
    e.range.getColumn() !== 5
  ) {
    return;
  }


  /*
   * Ignore header.
   */

  if (
    e.range.getRow() === 1
  ) {
    return;
  }


  const newStatus =
    String(
      e.range.getValue() || ''
    ).trim();


  const timestampCell =
    sheet.getRange(
      e.range.getRow(),
      6
    );


  /*
   * If status was cleared,
   * clear the timestamp too.
   */

  if (!newStatus) {

    timestampCell.clearContent();

    return;

  }


  /*
   * Record exactly when
   * the status was changed.
   */

  timestampCell.setValue(
    new Date()
  );

}
/****************************************************
 * AUTOMATION RULE MANAGER
 ****************************************************/

function createAutomationRule(data) {

  if (!data) {
    throw new Error('No rule data provided.');
  }

  const sheet =
    SpreadsheetApp
      .getActiveSpreadsheet()
      .getSheetByName(AUTOMATION_SHEET);

  if (!sheet) {
    throw new Error(
      'AUTOMATION_RULES sheet not found.'
    );
  }

  const ruleName =
    String(data.ruleName || '').trim();

  const staffStatus =
    String(data.staffStatus || '').trim();

  const templateId =
    String(data.templateId || '').trim();

  const delayHours =
    Number(data.delayHours || 0);

  const active =
    data.active === false
      ? false
      : true;


  if (!ruleName) {
    throw new Error(
      'Rule name is required.'
    );
  }

  if (!staffStatus) {
    throw new Error(
      'Staff status is required.'
    );
  }

  if (!templateId) {
    throw new Error(
      'Email template is required.'
    );
  }


  if (delayHours < 0) {
    throw new Error(
      'Delay hours cannot be negative.'
    );
  }


  const templates =
    getTemplates();

  const template =
    templates.find(
      item =>
        String(item.id) ===
        templateId
    );


  if (!template) {
    throw new Error(
      'Selected email template was not found.'
    );
  }


  const ruleId =
    'RULE-' +
    Utilities.getUuid()
      .substring(0, 8)
      .toUpperCase();


  sheet.appendRow([

    ruleId,

    ruleName,

    staffStatus,

    templateId,

    delayHours,

    active

  ]);


  return {
    success: true,
    message:
      'Automation rule created successfully.',
    ruleId: ruleId
  };

}


/****************************************************
 * UPDATE AUTOMATION RULE
 ****************************************************/

function updateAutomationRule(data) {

  if (!data || !data.ruleId) {
    throw new Error(
      'Rule ID is required.'
    );
  }


  const sheet =
    SpreadsheetApp
      .getActiveSpreadsheet()
      .getSheetByName(AUTOMATION_SHEET);

  if (!sheet) {
    throw new Error(
      'AUTOMATION_RULES sheet not found.'
    );
  }


  const values =
    sheet
      .getDataRange()
      .getValues();


  let rowNumber = -1;


  for (
    let i = 1;
    i < values.length;
    i++
  ) {

    if (
      String(values[i][0]) ===
      String(data.ruleId)
    ) {

      rowNumber =
        i + 1;

      break;

    }

  }


  if (rowNumber === -1) {
    throw new Error(
      'Automation rule not found.'
    );
  }


  const ruleName =
    String(data.ruleName || '').trim();

  const staffStatus =
    String(data.staffStatus || '').trim();

  const templateId =
    String(data.templateId || '').trim();

  const delayHours =
    Number(data.delayHours || 0);

  const active =
    data.active === false
      ? false
      : true;


  if (!ruleName) {
    throw new Error(
      'Rule name is required.'
    );
  }

  if (!staffStatus) {
    throw new Error(
      'Staff status is required.'
    );
  }

  if (!templateId) {
    throw new Error(
      'Email template is required.'
    );

  }

  if (delayHours < 0) {
    throw new Error(
      'Delay hours cannot be negative.'
    );
  }


  const templates =
    getTemplates();

  const template =
    templates.find(
      item =>
        String(item.id) ===
        templateId
    );


  if (!template) {
    throw new Error(
      'Selected email template was not found.'
    );
  }


  sheet
    .getRange(
      rowNumber,
      1,
      1,
      6
    )
    .setValues([

      [

        data.ruleId,

        ruleName,

        staffStatus,

        templateId,

        delayHours,

        active

      ]

    ]);


  return {
    success: true,
    message:
      'Automation rule updated successfully.'
  };

}


/****************************************************
 * ARCHIVE / DEACTIVATE AUTOMATION RULE
 ****************************************************/

function archiveAutomationRule(ruleId) {

  if (!ruleId) {
    throw new Error(
      'Rule ID is required.'
    );
  }


  const sheet =
    SpreadsheetApp
      .getActiveSpreadsheet()
      .getSheetByName(AUTOMATION_SHEET);

  if (!sheet) {
    throw new Error(
      'AUTOMATION_RULES sheet not found.'
    );
  }


  const values =
    sheet
      .getDataRange()
      .getValues();


  for (
    let i = 1;
    i < values.length;
    i++
  ) {

    if (
      String(values[i][0]) ===
      String(ruleId)
    ) {

      sheet
        .getRange(
          i + 1,
          6
        )
        .setValue(false);


      return {
        success: true,
        message:
          'Automation rule deactivated.'
      };

    }

  }


  throw new Error(
    'Automation rule not found.'
  );

}


/****************************************************
 * REACTIVATE AUTOMATION RULE
 ****************************************************/

function activateAutomationRule(ruleId) {

  if (!ruleId) {
    throw new Error(
      'Rule ID is required.'
    );
  }


  const sheet =
    SpreadsheetApp
      .getActiveSpreadsheet()
      .getSheetByName(AUTOMATION_SHEET);

  if (!sheet) {
    throw new Error(
      'AUTOMATION_RULES sheet not found.'
    );
  }


  const values =
    sheet
      .getDataRange()
      .getValues();


  for (
    let i = 1;
    i < values.length;
    i++
  ) {

    if (
      String(values[i][0]) ===
      String(ruleId)
    ) {

      sheet
        .getRange(
          i + 1,
          6
        )
        .setValue(true);


      return {
        success: true,
        message:
          'Automation rule activated.'
      };

    }

  }


  throw new Error(
    'Automation rule not found.'
  );

}
/****************************************************
 * DEMO - RUN AUTOMATION NOW
 * Manually runs the automatic email processor
 ****************************************************/

function runAutomationNow() {

  try {

    // Run the existing automation engine
    processAutomaticEmails();

    const result = {
      success: true,
      message: 'Automation run completed successfully.'
    };

    // If called from the spreadsheet menu, show the normal alert
    try {
      SpreadsheetApp.getUi().alert(
        'Automation Test Complete',
        'The automatic email processor has finished running.\n\n' +
        'Check EMAIL_LOG and the recipient inbox to confirm the result.',
        SpreadsheetApp.getUi().ButtonSet.OK
      );
    } catch (uiError) {
      // Ignore UI errors when called from the HTML dashboard
    }

    // Return result to AutomationDashboard.html
    return result;

  } catch (error) {

    const message = error && error.message
      ? error.message
      : String(error);

    // If called from the spreadsheet menu, show the normal alert
    try {
      SpreadsheetApp.getUi().alert(
        'Automation Test Failed',
        message,
        SpreadsheetApp.getUi().ButtonSet.OK
      );
    } catch (uiError) {
      // Ignore UI errors when called from the HTML dashboard
    }

    // Return a structured error to the dashboard
    return {
      success: false,
      message: message
    };
  }
}
/****************************************************
 * AUTOMATION DASHBOARD
 ****************************************************/

function getAutomationDashboard() {

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const staffSheet =
    ss.getSheetByName(SHEETS.STAFF);

  const logSheet =
    ss.getSheetByName(SHEETS.LOG);

  const rulesSheet =
    ss.getSheetByName(AUTOMATION_SHEET);

  const staff = getStaff();

  const rules =
    getAutomationRules();

  let logs = [];

  if (logSheet && logSheet.getLastRow() >= 2) {

    const lastColumn =
      Math.max(logSheet.getLastColumn(), 11);

    logs = logSheet
      .getRange(
        2,
        1,
        logSheet.getLastRow() - 1,
        lastColumn
      )
      .getValues()
      .filter(row => row[0])
      .map(row => {

        return {
          logId: String(row[0] || ''),
          date: row[1] || '',
          staffId: String(row[2] || ''),
          staffName: String(row[3] || ''),
          email: String(row[4] || ''),
          template: String(row[5] || ''),
          subject: String(row[6] || ''),
          status: String(row[7] || ''),
          sentBy: String(row[8] || ''),
          messageId: String(row[9] || ''),
          automationKey: String(row[10] || '')
        };

      });

  }

  /*
   * AUTOMATED LOGS ONLY
   */
  const automationLogs =
    logs.filter(log =>
      String(log.sentBy)
        .trim()
        .toUpperCase() === 'AUTOMATION'
    );

  /*
   * SENT
   */
  const sentCount =
    automationLogs.filter(log =>
      String(log.status)
        .trim()
        .toLowerCase() === 'sent'
    ).length;

  /*
   * FAILED
   */
  const failedCount =
    automationLogs.filter(log =>
      String(log.status)
        .trim()
        .toLowerCase() === 'failed'
    ).length;

  /*
   * ACTIVE RULES
   */
  const activeRules =
    rules.filter(rule => rule.active).length;

  /*
   * STAFF CURRENTLY MATCHING ACTIVE RULES
   */
  let waitingCount = 0;
  let eligibleCount = 0;
  let duplicateCount = 0;

  const now = new Date();

  rules
    .filter(rule => rule.active)
    .forEach(rule => {

      const matchingStaff =
        staff.filter(person => {

          return String(person.status)
            .trim()
            .toLowerCase() ===
            String(rule.staffStatus)
              .trim()
              .toLowerCase();

        });

      matchingStaff.forEach(person => {

        const automationKey =
          person.id + '|' + rule.id;

        /*
         * Already sent
         */
        if (
          hasAutomationBeenSent_(
            automationKey
          )
        ) {

          duplicateCount++;

          return;
        }

        /*
         * Determine eligibility
         */
        const eligibleTime =
          getAutomationEligibleTime_(
            person.id,
            rule.id,
            rule.delayHours
          );

        if (
          now.getTime() <
          eligibleTime.getTime()
        ) {

          waitingCount++;

        } else {

          eligibleCount++;

        }

      });

    });

  /*
   * LAST AUTOMATION ACTIVITY
   */
  let lastActivity = null;

  if (automationLogs.length) {

    automationLogs.sort((a, b) => {

      const dateA =
        new Date(a.date).getTime() || 0;

      const dateB =
        new Date(b.date).getTime() || 0;

      return dateB - dateA;

    });

    lastActivity =
      automationLogs[0];

  }

  /*
   * RECENT AUTOMATION LOGS
   */
  const recentLogs =
    automationLogs
      .sort((a, b) => {

        const dateA =
          new Date(a.date).getTime() || 0;

        const dateB =
          new Date(b.date).getTime() || 0;

        return dateB - dateA;

      })
      .slice(0, 10)
      .map(log => {

        return {
          logId: log.logId,
          date: formatDashboardDate_(log.date),
          staffId: log.staffId,
          staffName: log.staffName,
          email: log.email,
          template: log.template,
          subject: log.subject,
          status: log.status,
          sentBy: log.sentBy,
          automationKey: log.automationKey
        };

      });

  return {

    success: true,

    stats: {

      sent: sentCount,

      failed: failedCount,

      waiting: waitingCount,

      eligible: eligibleCount,

      duplicates: duplicateCount,

      staff: staff.length,

      activeRules: activeRules

    },

    lastActivity:
      lastActivity
        ? {
            date:
              formatDashboardDate_(
                lastActivity.date
              ),
            staffName:
              lastActivity.staffName,
            subject:
              lastActivity.subject,
            status:
              lastActivity.status
          }
        : null,

    recentLogs: recentLogs

  };

}


/****************************************************
 * FORMAT DASHBOARD DATE
 ****************************************************/

function formatDashboardDate_(dateValue) {

  if (!dateValue) {
    return '';
  }

  const date =
    new Date(dateValue);

  if (isNaN(date.getTime())) {
    return String(dateValue);
  }

  return Utilities.formatDate(
    date,
    Session.getScriptTimeZone(),
    'dd MMM yyyy, HH:mm'
  );

}
/****************************************************
 * STAFF MANAGEMENT MODULE
 * Version 1.0
 ****************************************************/


/****************************************************
 * GET STAFF MANAGEMENT DATA
 ****************************************************/

function getStaffManagementData() {

  const sheet =
    SpreadsheetApp
      .getActiveSpreadsheet()
      .getSheetByName(SHEETS.STAFF);

  if (!sheet || sheet.getLastRow() < 2) {
    return [];
  }

  const lastRow = sheet.getLastRow();

  const values =
    sheet
      .getRange(
        2,
        1,
        lastRow - 1,
        6
      )
      .getValues();

  return values
    .filter(row => row[0])
    .map(row => {

      let statusUpdatedAt = row[5] || '';

      if (
        statusUpdatedAt &&
        statusUpdatedAt instanceof Date
      ) {

        statusUpdatedAt =
          Utilities.formatDate(
            statusUpdatedAt,
            Session.getScriptTimeZone(),
            'dd MMM yyyy, HH:mm'
          );

      }

      return {

        id:
          String(row[0] || ''),

        name:
          String(row[1] || ''),

        email:
          String(row[2] || ''),

        department:
          String(row[3] || ''),

        status:
          String(row[4] || ''),

        statusUpdatedAt:
          statusUpdatedAt

      };

    });

}


/****************************************************
 * GENERATE STAFF ID
 ****************************************************/

function generateStaffId_() {

  const sheet =
    SpreadsheetApp
      .getActiveSpreadsheet()
      .getSheetByName(SHEETS.STAFF);

  if (!sheet) {
    throw new Error(
      'STAFF sheet not found.'
    );
  }

  const lastRow =
    sheet.getLastRow();

  if (lastRow < 2) {
    return 'STF-001';
  }

  const values =
    sheet
      .getRange(
        2,
        1,
        lastRow - 1,
        1
      )
      .getValues();

  let highestNumber = 0;

  values.forEach(row => {

    const id =
      String(row[0] || '').trim();

    const match =
      id.match(/^STF-(\d+)$/i);

    if (match) {

      const number =
        parseInt(
          match[1],
          10
        );

      if (number > highestNumber) {
        highestNumber = number;
      }

    }

  });

  return (
    'STF-' +
    String(highestNumber + 1)
      .padStart(3, '0')
  );

}


/****************************************************
 * CHECK STAFF EMAIL DUPLICATE
 ****************************************************/

function staffEmailExists_(
  email,
  excludeStaffId
) {

  const staff =
    getStaff();

  const normalizedEmail =
    String(email || '')
      .trim()
      .toLowerCase();

  const excludedId =
    String(excludeStaffId || '')
      .trim();

  return staff.some(person => {

    const existingEmail =
      String(person.email || '')
        .trim()
        .toLowerCase();

    const existingId =
      String(person.id || '')
        .trim();

    return (
      existingEmail === normalizedEmail &&
      existingEmail !== '' &&
      existingId !== excludedId
    );

  });

}


/****************************************************
 * ADD STAFF
 ****************************************************/

function addStaff(data) {

  if (!data) {
    throw new Error(
      'Staff data is required.'
    );
  }

  const name =
    String(data.name || '').trim();

  const email =
    String(data.email || '').trim();

  const department =
    String(data.department || '').trim();

  const status =
    String(data.status || 'Active').trim();


  if (!name) {
    throw new Error(
      'Full name is required.'
    );
  }

  if (!email) {
    throw new Error(
      'Email address is required.'
    );
  }

  if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) {
    throw new Error(
      'Please enter a valid email address.'
    );
  }

  if (!department) {
    throw new Error(
      'Department is required.'
    );
  }


  if (
    staffEmailExists_(email)
  ) {

    throw new Error(
      'A staff member with this email address already exists.'
    );

  }


  const sheet =
    SpreadsheetApp
      .getActiveSpreadsheet()
      .getSheetByName(SHEETS.STAFF);

  if (!sheet) {
    throw new Error(
      'STAFF sheet not found.'
    );
  }


  const staffId =
    generateStaffId_();

  const statusUpdatedAt =
    new Date();


  sheet.appendRow([

    staffId,

    name,

    email,

    department,

    status,

    statusUpdatedAt

  ]);


  return {

    success: true,

    message:
      'Staff member added successfully.',

    staff: {

      id: staffId,

      name: name,

      email: email,

      department: department,

      status: status,

      statusUpdatedAt:
        Utilities.formatDate(
          statusUpdatedAt,
          Session.getScriptTimeZone(),
          'dd MMM yyyy, HH:mm'
        )

    }

  };

}


/****************************************************
 * UPDATE STAFF
 ****************************************************/

function updateStaff(data) {

  if (!data || !data.staffId) {
    throw new Error(
      'Staff ID is required.'
    );
  }


  const sheet =
    SpreadsheetApp
      .getActiveSpreadsheet()
      .getSheetByName(SHEETS.STAFF);

  if (!sheet) {
    throw new Error(
      'STAFF sheet not found.'
    );
  }


  const staffId =
    String(data.staffId).trim();

  const name =
    String(data.name || '').trim();

  const email =
    String(data.email || '').trim();

  const department =
    String(data.department || '').trim();

  const status =
    String(data.status || '').trim();


  if (!name) {
    throw new Error(
      'Full name is required.'
    );
  }

  if (!email) {
    throw new Error(
      'Email address is required.'
    );
  }

  if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) {
    throw new Error(
      'Please enter a valid email address.'
    );
  }

  if (!department) {
    throw new Error(
      'Department is required.'
    );
  }

  if (!status) {
    throw new Error(
      'Staff status is required.'
    );
  }


  if (
    staffEmailExists_(
      email,
      staffId
    )
  ) {

    throw new Error(
      'Another staff member is already using this email address.'
    );

  }


  const lastRow =
    sheet.getLastRow();

  if (lastRow < 2) {
    throw new Error(
      'No staff records found.'
    );
  }


  const values =
    sheet
      .getRange(
        2,
        1,
        lastRow - 1,
        6
      )
      .getValues();


  let rowNumber = -1;


  for (
    let i = 0;
    i < values.length;
    i++
  ) {

    if (
      String(values[i][0]).trim() ===
      staffId
    ) {

      rowNumber =
        i + 2;

      break;

    }

  }


  if (rowNumber === -1) {
    throw new Error(
      'Staff member not found.'
    );
  }


  const existingStatus =
    String(
      values[rowNumber - 2][4] || ''
    ).trim();


  let statusUpdatedAt =
    values[rowNumber - 2][5] || '';


  /*
   * Update timestamp only when
   * the staff status changes.
   */

  if (
    existingStatus.toLowerCase() !==
    status.toLowerCase()
  ) {

    statusUpdatedAt =
      new Date();

  }


  if (!statusUpdatedAt) {
    statusUpdatedAt =
      new Date();
  }


  sheet
    .getRange(
      rowNumber,
      1,
      1,
      6
    )
    .setValues([[
      staffId,
      name,
      email,
      department,
      status,
      statusUpdatedAt
    ]]);


  return {

    success: true,

    message:
      'Staff member updated successfully.'

  };

}


/****************************************************
 * UPDATE STAFF STATUS
 ****************************************************/

function updateStaffStatus(
  staffId,
  newStatus
) {

  if (!staffId) {
    throw new Error(
      'Staff ID is required.'
    );
  }

  if (!newStatus) {
    throw new Error(
      'New staff status is required.'
    );
  }


  const sheet =
    SpreadsheetApp
      .getActiveSpreadsheet()
      .getSheetByName(SHEETS.STAFF);

  if (!sheet) {
    throw new Error(
      'STAFF sheet not found.'
    );
  }


  const lastRow =
    sheet.getLastRow();

  if (lastRow < 2) {
    throw new Error(
      'No staff records found.'
    );
  }


  const values =
    sheet
      .getRange(
        2,
        1,
        lastRow - 1,
        6
      )
      .getValues();


  let rowNumber = -1;


  for (
    let i = 0;
    i < values.length;
    i++
  ) {

    if (
      String(values[i][0]).trim() ===
      String(staffId).trim()
    ) {

      rowNumber =
        i + 2;

      break;

    }

  }


  if (rowNumber === -1) {
    throw new Error(
      'Staff member not found.'
    );
  }


  const currentStatus =
    String(
      values[rowNumber - 2][4] || ''
    ).trim();


  const status =
    String(newStatus).trim();


  /*
   * Update status.
   */

  sheet
    .getRange(
      rowNumber,
      5
    )
    .setValue(status);


  /*
   * IMPORTANT:
   * Update the timestamp whenever
   * the status changes.
   */

  if (
    currentStatus.toLowerCase() !==
    status.toLowerCase()
  ) {

    sheet
      .getRange(
        rowNumber,
        6
      )
      .setValue(new Date());

  }


  return {

    success: true,

    message:
      'Staff status updated successfully.'

  };

}


/****************************************************
 * ACTIVATE STAFF
 ****************************************************/

function activateStaff(staffId) {

  return updateStaffStatus(
    staffId,
    'Active'
  );

}


/****************************************************
 * DEACTIVATE STAFF
 ****************************************************/

function deactivateStaff(staffId) {

  return updateStaffStatus(
    staffId,
    'Inactive'
  );

}


/****************************************************
 * GET STAFF MANAGEMENT SUMMARY
 ****************************************************/

function getStaffManagementSummary() {

  const staff =
    getStaffManagementData();


  const active =
    staff.filter(person =>
      String(person.status)
        .trim()
        .toLowerCase() ===
      'active'
    ).length;


  const inactive =
    staff.filter(person =>
      String(person.status)
        .trim()
        .toLowerCase() ===
      'inactive'
    ).length;


  const departments =
    new Set(
      staff
        .map(person =>
          String(person.department || '').trim()
        )
        .filter(Boolean)
    );


  return {

    total:
      staff.length,

    active:
      active,

    inactive:
      inactive,

    departments:
      departments.size

  };

}
function testStaffManagementBackend() {

  console.log('TEST 1: Starting');

  const ss =
    SpreadsheetApp.getActiveSpreadsheet();

  console.log(
    'TEST 2: Spreadsheet = ' +
    ss.getName()
  );

  const sheet =
    ss.getSheetByName(SHEETS.STAFF);

  if (!sheet) {
    throw new Error('STAFF sheet not found.');
  }

  console.log(
    'TEST 3: STAFF sheet found'
  );

  console.log(
    'TEST 4: Last row = ' +
    sheet.getLastRow()
  );

  const staff =
    getStaff();

  console.log(
    'TEST 5: getStaff completed'
  );

  console.log(
    'TEST 6: Staff count = ' +
    staff.length
  );

  console.log(
    'TEST COMPLETE'
  );

}
/****************************************************
 * ADD AUTOMATION RULE
 ****************************************************/
function addAutomationRule(data) {

  if (!data) {
    throw new Error('Rule data is required.');
  }

  const sheet =
    SpreadsheetApp
      .getActiveSpreadsheet()
      .getSheetByName(AUTOMATION_SHEET);

  if (!sheet) {
    throw new Error(
      'AUTOMATION_RULES sheet not found.'
    );
  }

  const ruleName =
    String(data.ruleName || '').trim();

  const staffStatus =
    String(data.staffStatus || '').trim();

  const templateId =
    String(data.templateId || '').trim();

  const delayHours =
    Number(data.delayHours || 0);

  const active =
    data.active === false
      ? false
      : true;


  // Validation
  if (!ruleName) {
    throw new Error(
      'Rule name is required.'
    );
  }

  if (!staffStatus) {
    throw new Error(
      'Staff status is required.'
    );
  }

  if (!templateId) {
    throw new Error(
      'Email template is required.'
    );
  }

  if (delayHours < 0) {
    throw new Error(
      'Delay hours cannot be negative.'
    );
  }


  // Confirm template exists
  const templates =
    getTemplates();

  const template =
    templates.find(
      item =>
        String(item.id) ===
        templateId
    );

  if (!template) {
    throw new Error(
      'Selected email template was not found.'
    );
  }


  // Generate new rule ID
  const ruleId =
    generateAutomationRuleId_();


  // Add rule
  sheet.appendRow([

    ruleId,

    ruleName,

    staffStatus,

    templateId,

    delayHours,

    active

  ]);


  return {

    success: true,

    message:
      'Automation rule created successfully.',

    ruleId:
      ruleId

  };

}


/****************************************************
 * GENERATE AUTOMATION RULE ID
 ****************************************************/
function generateAutomationRuleId_() {

  const sheet =
    SpreadsheetApp
      .getActiveSpreadsheet()
      .getSheetByName(AUTOMATION_SHEET);

  if (!sheet) {
    throw new Error(
      'AUTOMATION_RULES sheet not found.'
    );
  }


  const lastRow =
    sheet.getLastRow();


  if (lastRow < 2) {
    return 'Rule-001';
  }


  const ids =
    sheet
      .getRange(
        2,
        1,
        lastRow - 1,
        1
      )
      .getValues()
      .flat();


  let maxNumber = 0;


  ids.forEach(id => {

    const match =
      String(id).match(
        /Rule-(\d+)/i
      );

    if (match) {

      const number =
        Number(match[1]);

      if (number > maxNumber) {
        maxNumber = number;
      }

    }

  });


  return (
    'Rule-' +
    String(maxNumber + 1)
      .padStart(3, '0')
  );

}


/****************************************************
 * DELETE / DEACTIVATE AUTOMATION RULE
 ****************************************************/
function deleteAutomationRule(ruleId) {

  if (!ruleId) {
    throw new Error(
      'Rule ID is required.'
    );
  }


  const sheet =
    SpreadsheetApp
      .getActiveSpreadsheet()
      .getSheetByName(AUTOMATION_SHEET);

  if (!sheet) {
    throw new Error(
      'AUTOMATION_RULES sheet not found.'
    );
  }


  const values =
    sheet
      .getDataRange()
      .getValues();


  let rowNumber = -1;


  for (
    let i = 1;
    i < values.length;
    i++
  ) {

    if (
      String(values[i][0]) ===
      String(ruleId)
    ) {

      rowNumber =
        i + 1;

      break;

    }

  }


  if (rowNumber === -1) {
    throw new Error(
      'Automation rule not found.'
    );
  }


  /*
   * We deactivate rather than physically
   * delete the rule.
   *
   * This protects the automation history
   * and prevents accidental data loss.
   */

  sheet
    .getRange(
      rowNumber,
      6
    )
    .setValue(false);


  return {

    success: true,

    message:
      'Automation rule deactivated successfully.'

  };

}
function testAutomationRulesPage() {

  const html =
    HtmlService
      .createHtmlOutputFromFile('AutomationRules')
      .setTitle('NEXORA Automation Rules');

  SpreadsheetApp
    .getUi()
    .showSidebar(html);
}
/****************************************************
 * OPEN AUTOMATION RULES SIDEBAR
 ****************************************************/
function showAutomationRules() {

  const html =
    HtmlService
      .createHtmlOutputFromFile('AutomationRules')
      .setTitle('NEXORA Automation Rules');

  SpreadsheetApp
    .getUi()
    .showSidebar(html);
}
/****************************************************
 * OPEN EMAIL CENTER
 ****************************************************/
function showEmailCenter() {

  const html =
    HtmlService
      .createHtmlOutputFromFile('EmailCenter')
      .setTitle('NEXORA Email Center');

  SpreadsheetApp
    .getUi()
    .showSidebar(html);
}
function showTemplatesSidebar() {
  const html = HtmlService
    .createHtmlOutputFromFile('Templates')
    .setTitle('NEXORA Email Templates');

  SpreadsheetApp.getUi().showSidebar(html);
}
function getEmailTemplates() {

const ss = SpreadsheetApp.getActiveSpreadsheet();
const sheet = ss.getSheetByName(SHEETS.TEMPLATES);

if (!sheet) {
throw new Error(
'Template sheet "' + SHEETS.TEMPLATES + '" was not found.'
);
}

const data = sheet.getDataRange().getValues();

if (!data || data.length <= 1) {
return [];
}

const headers = data[0];
const results = [];

for (let r = 1; r < data.length; r++) {


const row = data[r];

let hasData = false;

for (let c = 0; c < row.length; c++) {
  if (String(row[c]).trim() !== '') {
    hasData = true;
    break;
  }
}

if (!hasData) {
  continue;
}

const template = {};

for (let c = 0; c < headers.length; c++) {

  const key = String(headers[c])
    .trim()
    .toLowerCase();

  template[key] = row[c] !== undefined && row[c] !== null
    ? row[c]
    : '';
}

results.push({

  id:
    template.id ||
    template['template id'] ||
    template.templateid ||
    '',

  name:
    template.name ||
    template['template name'] ||
    template.templatename ||
    template.title ||
    template.template ||
    '',

  subject:
    template.subject ||
    '',

  body:
    template.body ||
    template.message ||
    template.content ||
    '',

  status:
    template.status ||
    'Active'
});


}

return results;
}
function saveEmailTemplate(template) {

if (!template) {
throw new Error('Template data was not provided.');
}

const name = String(template.name || '').trim();
const subject = String(template.subject || '').trim();
const message = String(template.message || '').trim();
const status = String(template.status || 'Active').trim() || 'Active';
const templateId = String(template.id || '').trim();

if (!name) {
throw new Error('Template Name is required.');
}

if (!subject) {
throw new Error('Subject is required.');
}

if (!message) {
throw new Error('Message is required.');
}

const ss = SpreadsheetApp.getActiveSpreadsheet();
const sheet = ss.getSheetByName(SHEETS.TEMPLATES);

if (!sheet) {
throw new Error(
'Template sheet "' + SHEETS.TEMPLATES + '" was not found.'
);
}

const data = sheet.getDataRange().getValues();

if (data.length === 0) {
throw new Error('The EMAIL_TEMPLATES sheet is empty.');
}

const headers = data[0];
const columnMap = {};

for (let c = 0; c < headers.length; c++) {
const key = String(headers[c]).trim().toLowerCase();
columnMap[key] = c;
}

const idCol = columnMap['template id'];
const nameCol = columnMap['template name'];
const subjectCol = columnMap['subject'];
const messageCol = columnMap['message'];
const statusCol = columnMap['status'];

if (
idCol === undefined ||
nameCol === undefined ||
subjectCol === undefined ||
messageCol === undefined ||
statusCol === undefined
) {
throw new Error(
'EMAIL_TEMPLATES must contain: Template ID, Template Name, Subject, Message, Status.'
);
}

let targetRow = -1;

if (templateId) {

for (let r = 1; r < data.length; r++) {

  const existingId = String(data[r][idCol] || '').trim();

  if (existingId.toLowerCase() === templateId.toLowerCase()) {
    targetRow = r + 1;
    break;
  }
}

}

let finalId = templateId;

if (targetRow === -1) {

let maxNumber = 0;

for (let r = 1; r < data.length; r++) {

  const existingId = String(data[r][idCol] || '').trim();
  const match = existingId.match(/^TMP-(\d+)$/i);

  if (match) {

    const number = Number(match[1]);

    if (number > maxNumber) {
      maxNumber = number;
    }
  }
}

finalId = 'TMP-' + String(maxNumber + 1).padStart(3, '0');

targetRow = sheet.getLastRow() + 1;

}

sheet.getRange(targetRow, idCol + 1).setValue(finalId);
sheet.getRange(targetRow, nameCol + 1).setValue(name);
sheet.getRange(targetRow, subjectCol + 1).setValue(subject);
sheet.getRange(targetRow, messageCol + 1).setValue(message);
sheet.getRange(targetRow, statusCol + 1).setValue(status);

SpreadsheetApp.flush();

return {
success: true,
action: templateId ? 'updated' : 'created',
templateId: finalId,
name: name
};
}