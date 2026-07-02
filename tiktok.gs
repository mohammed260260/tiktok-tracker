/******************************************************
 * TikTok Tracker + Payment System
 *
 * CLEAN DASHBOARD VERSION
 *
 * Main Sheet:
 * All Cases
 *
 * Tracker Sheet:
 * TikTok Tracker
 *
 * Accountant Paid Sheet:
 * TikTok Paid
 *
 * TikTok marker:
 * Column E only
 *
 * Yusuf:
 * #38761d
 *
 * Suleman:
 * #cc4125
 *
 * Payment:
 * £25 per completed TikTok
 *
 * Payment cycle:
 * 7 days before end of month
 *
 * Main functions:
 * updateTikTokTracker()
 * sendMonthlyTikTokPaymentEmails()
 * sendTikTokPaymentSummaryTestToMe()
 * markSelectedTikTokRowsAsPaid()
 * createTikTokNightlyAndPaymentTriggers()
 ******************************************************/

const TIKTOK_CONFIG = {
  MAIN_SHEET_NAME: 'All Cases',
  TRACKER_SHEET_NAME: 'TikTok Tracker',
  PAID_SHEET_NAME: 'TikTok Paid',

  START_DATA_ROW: 1060,

  COL_SOLICITOR: 1,
  COL_ACCIDENT_DATE: 2,
  COL_CASE_REF: 3,
  COL_SOLICITOR_REF: 4,
  COL_TYPE: 5,
  COL_CLIENT_NAME: 6,

  ROW_SCAN_LAST_COLUMN: 14,

  YUSUF_COLOUR: '#38761d',
  SULEMAN_COLOUR: '#cc4125',

  COLOUR_TOLERANCE: 35,
  ROW_WIDE_RED_CELL_LIMIT: 3,

  RATE_PER_TIKTOK: 25,

  PERSON_EMAILS: {
    Yusuf: 'yusuf@speedyclaim.co.uk',
    Suleman: 'suleman.siddiq@speedyclaim.co.uk'
  },

  ROW_WIDE_RED_COLOURS: [
    '#ff0000',
    '#cc0000',
    '#cc4125',
    '#c00000',
    '#980000',
    '#d93025',
    '#e06666',
    '#ea4335'
  ],

  TRACKER_HEADER_ROW: 22,
  TRACKER_DATA_START_ROW: 23,
  DUE_THIS_RUN_START_ROW: 10,
  DUE_THIS_RUN_MAX_ROWS: 10,

  EXCLUDE_KEYWORDS: [
    'cancel',
    'cancelled',
    'canceled',
    'cancelled hire',
    'cancelled claim',
    'claim cancelled',
    'claim canceled'
  ]
};


/**
 * Main tracker update.
 */
function updateTikTokTracker() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const mainSheet = ss.getSheetByName(TIKTOK_CONFIG.MAIN_SHEET_NAME);

  if (!mainSheet) {
    throw new Error('Main sheet not found: ' + TIKTOK_CONFIG.MAIN_SHEET_NAME);
  }

  let trackerSheet = ss.getSheetByName(TIKTOK_CONFIG.TRACKER_SHEET_NAME);

  if (!trackerSheet) {
    trackerSheet = ss.insertSheet(TIKTOK_CONFIG.TRACKER_SHEET_NAME);
  }

  const existingPaymentMap = getExistingTikTokPaymentMap_(trackerSheet);

  try {
    trackerSheet.getDataRange().breakApart();
  } catch (err) {}

  try {
    trackerSheet.showColumns(1, 24);
  } catch (err) {}

  trackerSheet.clear();

  const lastRow = mainSheet.getLastRow();

  if (lastRow < TIKTOK_CONFIG.START_DATA_ROW) {
    buildEmptyTikTokTracker_(trackerSheet);
    return;
  }

  const numRows = lastRow - TIKTOK_CONFIG.START_DATA_ROW + 1;

  const rowScanValues = mainSheet
    .getRange(TIKTOK_CONFIG.START_DATA_ROW, 1, numRows, TIKTOK_CONFIG.ROW_SCAN_LAST_COLUMN)
    .getDisplayValues();

  const rowScanBackgrounds = mainSheet
    .getRange(TIKTOK_CONFIG.START_DATA_ROW, 1, numRows, TIKTOK_CONFIG.ROW_SCAN_LAST_COLUMN)
    .getBackgrounds();

  const values = mainSheet
    .getRange(TIKTOK_CONFIG.START_DATA_ROW, 1, numRows, TIKTOK_CONFIG.COL_CLIENT_NAME)
    .getDisplayValues();

  const typeBackgrounds = mainSheet
    .getRange(TIKTOK_CONFIG.START_DATA_ROW, TIKTOK_CONFIG.COL_TYPE, numRows, 1)
    .getBackgrounds();

  const rows = [];
  const nextPaymentWeek = getNextTikTokPaymentWeekText_();

  for (let i = 0; i < numRows; i++) {
    const rowNumber = TIKTOK_CONFIG.START_DATA_ROW + i;

    const solicitor = cleanTikTokValue_(values[i][TIKTOK_CONFIG.COL_SOLICITOR - 1]);
    const accidentDate = cleanTikTokValue_(values[i][TIKTOK_CONFIG.COL_ACCIDENT_DATE - 1]);
    const caseRef = cleanTikTokValue_(values[i][TIKTOK_CONFIG.COL_CASE_REF - 1]);
    const solicitorRef = cleanTikTokValue_(values[i][TIKTOK_CONFIG.COL_SOLICITOR_REF - 1]);
    const type = cleanTikTokValue_(values[i][TIKTOK_CONFIG.COL_TYPE - 1]);
    const clientName = cleanTikTokValue_(values[i][TIKTOK_CONFIG.COL_CLIENT_NAME - 1]);

    if (!caseRef && !clientName) {
      continue;
    }

    const typeColour = normaliseTikTokHex_(typeBackgrounds[i][0]);

    const textExcluded = isExcludedTikTokRowByText_(rowScanValues[i]);
    const redRowInfo = getRowWideRedInfo_(rowScanBackgrounds[i]);

    if (textExcluded || redRowInfo.isRedRow) {
      rows.push(buildTrackerRow_({
        person: '',
        caseRef,
        clientName,
        accidentDate,
        solicitor,
        type,
        solicitorRef,
        rowNumber,
        typeColour,
        matchType: '',
        distance: '',
        status: 'Excluded',
        old: {},
        paymentWeek: '',
        amount: 0,
        reason: textExcluded ? 'Cancelled text found' : 'Row-wide red colour found'
      }));
      continue;
    }

    const colourResult = getTikTokPersonFromColour_(typeColour);

    if (colourResult.person) {
      const paymentKey = makeTikTokPaymentKey_(colourResult.person, caseRef, clientName);
      const old = existingPaymentMap[paymentKey] || {};

      const paid = toBoolean_(old.paid);
      const emailSentDate = cleanTikTokValue_(old.emailSentDate);

      let status = 'Unpaid';

      if (paid) {
        status = 'Paid';
      } else if (emailSentDate) {
        status = 'Email Sent';
      }

      rows.push(buildTrackerRow_({
        person: colourResult.person,
        caseRef,
        clientName,
        accidentDate,
        solicitor,
        type,
        solicitorRef,
        rowNumber,
        typeColour,
        matchType: colourResult.matchType,
        distance: colourResult.distance,
        status,
        old,
        paymentWeek: cleanTikTokValue_(old.paymentWeek) || nextPaymentWeek,
        amount: TIKTOK_CONFIG.RATE_PER_TIKTOK,
        reason: ''
      }));

      continue;
    }

    if (isPossibleManualTikTokColour_(typeColour)) {
      rows.push(buildTrackerRow_({
        person: '',
        caseRef,
        clientName,
        accidentDate,
        solicitor,
        type,
        solicitorRef,
        rowNumber,
        typeColour,
        matchType: '',
        distance: '',
        status: 'Needs Review',
        old: {},
        paymentWeek: '',
        amount: 0,
        reason: 'Wrong / unknown colour'
      }));
    }
  }

  applyDuplicatePaymentKeySafety_(rows);
  sortTikTokRows_(rows);
  buildTikTokTracker_(trackerSheet, rows);

  const summary = getTikTokPaymentSummary_(rows);

  SpreadsheetApp.getUi().alert(
    'TikTok Tracker updated.\n\n' +
    'Current payment week: ' + summary.currentPaymentWeek + '\n\n' +
    'Yusuf completed this run: ' + summary.Yusuf.completed + '\n' +
    'Yusuf paid this run: £' + summary.Yusuf.paidAmount + '\n' +
    'Yusuf balance due this run: £' + summary.Yusuf.balance + '\n\n' +
    'Suleman completed this run: ' + summary.Suleman.completed + '\n' +
    'Suleman paid this run: £' + summary.Suleman.paidAmount + '\n' +
    'Suleman balance due this run: £' + summary.Suleman.balance + '\n\n' +
    'Lifetime paid: £' + summary.History.totalPaid + '\n' +
    'Paid records: ' + summary.History.totalPaidCount + '\n\n' +
    'Needs Review: ' + summary.NeedsReview.count + '\n' +
    'Excluded: ' + summary.Excluded.count
  );
}


/**
 * Creates one row for TikTok Tracker.
 */
function buildTrackerRow_(data) {
  const old = data.old || {};
  const paid = toBoolean_(old.paid);
  const sendEmail = toBoolean_(old.sendEmail);
  const emailSentDate = cleanTikTokValue_(old.emailSentDate);
  const paymentWeek = cleanTikTokValue_(data.paymentWeek);
  const status = cleanTikTokValue_(data.status);
  const amount = Number(data.amount) || 0;
  const outstandingThisRun = getTikTokOutstandingThisRun_(amount, paymentWeek, status, paid);
  const actionNeeded = getTikTokActionNeeded_(status, outstandingThisRun, sendEmail, emailSentDate, paid);

  const paymentKey = data.person
    ? makeTikTokPaymentKey_(data.person, data.caseRef, data.clientName)
    : '';

  return [
    data.person,
    data.caseRef,
    data.clientName,
    data.accidentDate,
    data.solicitor,
    data.type,
    data.solicitorRef,
    data.rowNumber,
    data.typeColour,
    data.matchType,
    data.distance,
    amount,
    amount,
    outstandingThisRun,
    actionNeeded,
    sendEmail,
    emailSentDate,
    paid,
    cleanTikTokValue_(old.paidDate),
    paymentWeek,
    status,
    cleanTikTokValue_(old.notes) || data.reason,
    paymentKey,
    Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm')
  ];
}


/**
 * Builds the TikTok Tracker sheet.
 */
function buildTikTokTracker_(sheet, rows) {
  const summary = getTikTokPaymentSummary_(rows);
  const nearMatchCount = rows.filter(row => row[9] === 'Near Match').length;
  const nextPaymentWeek = getNextTikTokPaymentWeekText_();
  const dueThisRunRows = rows
    .filter(row => (Number(row[13]) || 0) > 0)
    .slice(0, TIKTOK_CONFIG.DUE_THIS_RUN_MAX_ROWS);

  const now = Utilities.formatDate(
    new Date(),
    Session.getScriptTimeZone(),
    'dd/MM/yyyy HH:mm'
  );

  const topRows = [
    ['TikTok Tracker', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', ''],
    ['Last Updated', now, '', 'Next Payment Week', nextPaymentWeek, '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', ''],
    ['', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', ''],

    ['YUSUF', '', '', '', 'SULEMAN', '', '', '', 'CURRENT RUN TOTAL', '', '', '', 'OUTSTANDING', '', '', '', 'ISSUES', '', '', '', 'PAID HISTORY', '', '', ''],
    ['Completed This Run', summary.Yusuf.completed, '', '', 'Completed This Run', summary.Suleman.completed, '', '', 'Completed This Run', summary.Total.completed, '', '', 'Yusuf Outstanding This Run', summary.Yusuf.balance, '', '', 'Needs Review', summary.NeedsReview.count, '', '', 'Lifetime Paid', summary.History.totalPaid, '', ''],
    ['Total Owed This Run', summary.Yusuf.totalOwed, '', '', 'Total Owed This Run', summary.Suleman.totalOwed, '', '', 'Total Owed This Run', summary.Total.totalOwed, '', '', 'Suleman Outstanding This Run', summary.Suleman.balance, '', '', 'Excluded', summary.Excluded.count, '', '', 'Paid Records', summary.History.totalPaidCount, '', ''],
    ['Paid This Run', summary.Yusuf.paidAmount, '', '', 'Paid This Run', summary.Suleman.paidAmount, '', '', 'Paid This Run', summary.Total.paidAmount, '', '', 'Total Outstanding This Run', summary.Total.balance, '', '', 'Near Matches', nearMatchCount, '', '', 'Yusuf Lifetime', summary.History.Yusuf.totalPaid, '', ''],
    ['Balance Due This Run', summary.Yusuf.balance, '', '', 'Balance Due This Run', summary.Suleman.balance, '', '', 'Balance Due This Run', summary.Total.balance, '', '', 'Rows Awaiting Payment', summary.Total.awaitingPaymentCount, 'Rows Ready To Email', summary.Total.readyToEmailCount, 'Rows Payment Due', summary.Total.paymentDueCount, '', '', 'Suleman Lifetime', summary.History.Suleman.totalPaid, '', ''],
    ['', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', ''],
    ['DUE THIS RUN', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', ''],
    ['TikTok By', 'Case Ref', 'Client Name', 'Amount Owed / Outstanding This Run', 'Action Needed', 'Send Email?', 'Email Sent Date', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '']
  ];

  if (dueThisRunRows.length > 0) {
    dueThisRunRows.forEach(row => {
      topRows.push([
        row[0],
        row[1],
        row[2],
        row[13],
        row[14],
        row[15],
        row[16],
        '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', ''
      ]);
    });
  } else {
    topRows.push(['No outstanding rows this run', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '']);
  }

  while (topRows.length < TIKTOK_CONFIG.TRACKER_HEADER_ROW - 1) {
    topRows.push(['', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '']);
  }

  topRows.push(
    [
      'TikTok By',
      'Case Ref',
      'Client Name',
      'Accident Date',
      'Solicitor',
      'Type',
      'Solicitor Ref',
      'Main Sheet Row',
      'Detected Colour',
      'Match Type',
      'Colour Distance',
      'Rate',
      'Amount Owed',
      'Outstanding This Run',
      'Action Needed',
      'Send Email?',
      'Email Sent Date',
      'Paid?',
      'Paid Date',
      'Payment Week',
      'Status',
      'Notes',
      'Payment Key',
      'Last Seen'
    ]
  );

  sheet.getRange(1, 1, topRows.length, topRows[0].length).setValues(topRows);

  if (rows.length > 0) {
    sheet.getRange(TIKTOK_CONFIG.TRACKER_DATA_START_ROW, 1, rows.length, rows[0].length).setValues(rows);

    sheet.getRange(TIKTOK_CONFIG.TRACKER_DATA_START_ROW, 16, rows.length, 1).insertCheckboxes();
    sheet.getRange(TIKTOK_CONFIG.TRACKER_DATA_START_ROW, 18, rows.length, 1).insertCheckboxes();
  }

  formatTikTokTracker_(sheet, rows.length);
}


/**
 * Sends real monthly payment emails to Yusuf and Suleman.
 * This updates Email Sent status.
 */
function sendMonthlyTikTokPaymentEmails() {
  updateTikTokTracker();

  const grouped = getUnpaidTikTokGroupedRows_();

  const paymentRunDate = getNextTikTokPaymentRunDate_();
  const paymentWeek = getPaymentWeekCommencingText_(paymentRunDate);
  const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy');

  let emailsSent = 0;

  Object.keys(grouped).forEach(person => {
    const items = grouped[person];

    if (items.length === 0) {
      return;
    }

    const email = TIKTOK_CONFIG.PERSON_EMAILS[person];

    if (!email) {
      return;
    }

    const emailParts = buildTikTokPaymentEmailBodies_(person, items, paymentRunDate, paymentWeek, false);

    MailApp.sendEmail({
      to: email,
      subject: emailParts.subject,
      body: emailParts.plainBody,
      htmlBody: emailParts.htmlBody
    });

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(TIKTOK_CONFIG.TRACKER_SHEET_NAME);

    items.forEach(item => {
      sheet.getRange(item.rowNumber, 16).setValue(true);
      sheet.getRange(item.rowNumber, 17).setValue(today);
      sheet.getRange(item.rowNumber, 20).setValue(paymentWeek);
      sheet.getRange(item.rowNumber, 21).setValue('Email Sent');
      sheet.getRange(item.rowNumber, 15).setValue('EMAILED - AWAITING PAYMENT');
    });

    emailsSent++;
  });

  SpreadsheetApp.getUi().alert('Payment summary emails sent: ' + emailsSent);
}


/**
 * TEST ONLY.
 * Sends one Yusuf test summary and one Suleman test summary to you/admin.
 * Does NOT tick Email Sent.
 * Does NOT mark anything paid.
 * Does NOT update tracker rows.
 */
function sendTikTokPaymentSummaryTestToMe() {
  updateTikTokTracker();

  const TEST_EMAIL = Session.getActiveUser().getEmail() || 'mohammed@speedyclaim.co.uk';
  const grouped = getUnpaidTikTokGroupedRows_();

  const paymentRunDate = getNextTikTokPaymentRunDate_();
  const paymentWeek = getPaymentWeekCommencingText_(paymentRunDate);

  let emailsSent = 0;

  ['Yusuf', 'Suleman'].forEach(person => {
    let items = grouped[person] || [];

    if (items.length === 0) {
      items = [{
        rowNumber: '',
        caseRef: 'TEST ONLY',
        clientName: 'No unpaid TikTok rows currently found for ' + person,
        amount: 0
      }];
    }

    const emailParts = buildTikTokPaymentEmailBodies_(person, items, paymentRunDate, paymentWeek, true);

    MailApp.sendEmail({
      to: TEST_EMAIL,
      subject: '[TEST] ' + emailParts.subject,
      body: emailParts.plainBody,
      htmlBody: emailParts.htmlBody
    });

    emailsSent++;
  });

  SpreadsheetApp.getUi().alert(
    'Test payment summary emails sent to: ' + TEST_EMAIL + '\n\n' +
    'Emails sent: ' + emailsSent + '\n\n' +
    'No rows were marked as emailed or paid.'
  );
}


/**
 * Gets unpaid rows grouped by Yusuf/Suleman.
 */
function getUnpaidTikTokGroupedRows_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(TIKTOK_CONFIG.TRACKER_SHEET_NAME);

  if (!sheet) {
    throw new Error('TikTok Tracker not found.');
  }

  const lastRow = sheet.getLastRow();

  const grouped = {
    Yusuf: [],
    Suleman: []
  };

  if (lastRow < TIKTOK_CONFIG.TRACKER_DATA_START_ROW) {
    return grouped;
  }

  const values = sheet.getRange(TIKTOK_CONFIG.TRACKER_DATA_START_ROW, 1, lastRow - TIKTOK_CONFIG.TRACKER_HEADER_ROW, 24).getValues();

  for (let i = 0; i < values.length; i++) {
    const rowNumber = TIKTOK_CONFIG.TRACKER_DATA_START_ROW + i;
    const row = values[i];

    const person = cleanTikTokValue_(row[0]);
    const caseRef = cleanTikTokValue_(row[1]);
    const clientName = cleanTikTokValue_(row[2]);
    const amount = Number(row[12]) || TIKTOK_CONFIG.RATE_PER_TIKTOK;
    const outstandingThisRun = Number(row[13]) || 0;
    const sendEmail = toBoolean_(row[15]);
    const emailSentDate = cleanTikTokValue_(row[16]);
    const paid = toBoolean_(row[17]);
    const actionNeeded = cleanTikTokValue_(row[14]);

    if (!grouped[person]) {
      continue;
    }

    if (!sendEmail || emailSentDate || paid) {
      continue;
    }

    if (actionNeeded === 'EXCLUDED' || actionNeeded === 'NEEDS REVIEW' || outstandingThisRun <= 0) {
      continue;
    }

    grouped[person].push({
      rowNumber,
      caseRef,
      clientName,
      amount
    });
  }

  return grouped;
}


/**
 * Builds plain and HTML email bodies.
 */
function buildTikTokPaymentEmailBodies_(person, items, paymentRunDate, paymentWeek, isTest) {
  const total = items.reduce((sum, item) => sum + item.amount, 0);

  const caseLines = items
    .map(item => '- ' + item.caseRef + ' | ' + item.clientName + ' | £' + item.amount)
    .join('\n');

  const subject = 'TikTok payment confirmation - ' + paymentWeek;

  const plainBody =
    (isTest ? 'TEST EMAIL ONLY - No tracker rows have been updated.\n\n' : '') +
    'Hi ' + person + ',\n\n' +
    'This is to confirm your TikTok payment summary.\n\n' +
    'Payment week: ' + paymentWeek + '\n' +
    'Payment run date: ' + Utilities.formatDate(paymentRunDate, Session.getScriptTimeZone(), 'dd/MM/yyyy') + '\n\n' +
    'Completed TikToks: ' + items.filter(item => item.amount > 0).length + '\n' +
    'Rate: £' + TIKTOK_CONFIG.RATE_PER_TIKTOK + ' per claim\n' +
    'Amount due: £' + total + '\n\n' +
    'Cases included:\n' +
    caseLines + '\n\n' +
    'Good work. Keep the momentum going.\n\n' +
    'This payment will be paid during the week shown above.\n\n' +
    'Regards,\n' +
    'Speedy Claim';

  const caseRowsHtml = items
    .map(item => {
      return (
        '<tr>' +
          '<td style="padding:10px;border-bottom:1px solid #e5e7eb;">' + escapeHtml_(item.caseRef) + '</td>' +
          '<td style="padding:10px;border-bottom:1px solid #e5e7eb;">' + escapeHtml_(item.clientName) + '</td>' +
          '<td style="padding:10px;border-bottom:1px solid #e5e7eb;text-align:right;font-weight:600;">£' + item.amount + '</td>' +
        '</tr>'
      );
    })
    .join('');

  const htmlBody =
    '<div style="font-family:Arial,Helvetica,sans-serif;background:#f3f4f6;padding:24px;color:#111827;">' +
      '<div style="max-width:680px;margin:0 auto;background:#ffffff;border-radius:14px;overflow:hidden;border:1px solid #e5e7eb;">' +

        (isTest
          ? '<div style="background:#991b1b;color:#ffffff;padding:12px 28px;font-size:13px;font-weight:700;">TEST EMAIL ONLY - No tracker rows have been updated</div>'
          : '') +

        '<div style="background:#0f172a;color:#ffffff;padding:24px 28px;">' +
          '<h2 style="margin:0;font-size:22px;">TikTok Payment Summary</h2>' +
          '<p style="margin:8px 0 0;font-size:14px;color:#cbd5e1;">Speedy Claim TikTok work payment confirmation</p>' +
        '</div>' +

        '<div style="padding:26px 28px;">' +
          '<p style="font-size:15px;margin:0 0 14px;">Hi <strong>' + escapeHtml_(person) + '</strong>,</p>' +

          '<p style="font-size:15px;line-height:1.6;margin:0 0 20px;">' +
            'This is to confirm your TikTok payment summary. The amount below is due for the completed TikToks included in this payment run.' +
          '</p>' +

          '<div style="background:#ecfdf5;border:1px solid #bbf7d0;border-radius:12px;padding:18px 20px;margin:20px 0;">' +
            '<div style="font-size:13px;color:#166534;font-weight:700;text-transform:uppercase;letter-spacing:0.04em;">Amount Due</div>' +
            '<div style="font-size:34px;font-weight:800;color:#166534;margin-top:6px;">£' + total + '</div>' +
            '<div style="font-size:14px;color:#166534;margin-top:6px;">' + items.filter(item => item.amount > 0).length + ' TikTok(s) × £' + TIKTOK_CONFIG.RATE_PER_TIKTOK + ' per claim</div>' +
          '</div>' +

          '<table style="width:100%;border-collapse:collapse;margin:18px 0;border:1px solid #e5e7eb;border-radius:10px;overflow:hidden;">' +
            '<tr style="background:#f9fafb;">' +
              '<td style="padding:10px;font-weight:700;border-bottom:1px solid #e5e7eb;">Payment Week</td>' +
              '<td style="padding:10px;border-bottom:1px solid #e5e7eb;">' + escapeHtml_(paymentWeek) + '</td>' +
            '</tr>' +
            '<tr>' +
              '<td style="padding:10px;font-weight:700;border-bottom:1px solid #e5e7eb;">Payment Run Date</td>' +
              '<td style="padding:10px;border-bottom:1px solid #e5e7eb;">' + Utilities.formatDate(paymentRunDate, Session.getScriptTimeZone(), 'dd/MM/yyyy') + '</td>' +
            '</tr>' +
            '<tr style="background:#f9fafb;">' +
              '<td style="padding:10px;font-weight:700;">Status</td>' +
              '<td style="padding:10px;">This payment will be paid during the week shown above.</td>' +
            '</tr>' +
          '</table>' +

          '<h3 style="font-size:16px;margin:26px 0 10px;">Cases included</h3>' +

          '<table style="width:100%;border-collapse:collapse;border:1px solid #e5e7eb;border-radius:10px;overflow:hidden;">' +
            '<thead>' +
              '<tr style="background:#0f172a;color:#ffffff;">' +
                '<th style="padding:10px;text-align:left;font-size:13px;">Case Ref</th>' +
                '<th style="padding:10px;text-align:left;font-size:13px;">Client Name</th>' +
                '<th style="padding:10px;text-align:right;font-size:13px;">Amount</th>' +
              '</tr>' +
            '</thead>' +
            '<tbody>' +
              caseRowsHtml +
            '</tbody>' +
          '</table>' +

          '<div style="background:#fff7ed;border:1px solid #fed7aa;border-radius:12px;padding:16px 18px;margin:24px 0;">' +
            '<p style="margin:0;font-size:15px;line-height:1.6;color:#7c2d12;">' +
              '<strong>Good work — keep going.</strong><br>' +
              'You have built up <strong>£' + total + '</strong> from the TikTok work for this payment run. Every completed TikTok helps push Speedy Claim online, brings more attention to the claims we handle, and increases what you earn from this side of the work. Keep the momentum going and aim to beat this number on the next payment run.' +
            '</p>' +
          '</div>' +

          '<p style="font-size:15px;line-height:1.6;margin:22px 0 0;">Regards,<br><strong>Speedy Claim</strong></p>' +
        '</div>' +

        '<div style="background:#f9fafb;padding:14px 28px;font-size:12px;color:#6b7280;border-top:1px solid #e5e7eb;">' +
          (isTest ? 'This is an automated test payment summary from Speedy Claim.' : 'This is an automated payment summary from Speedy Claim.') +
        '</div>' +

      '</div>' +
    '</div>';

  return {
    subject,
    plainBody,
    htmlBody
  };
}


/**
 * Mark selected tracker rows as paid.
 */
function markSelectedTikTokRowsAsPaid() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();

  if (sheet.getName() !== TIKTOK_CONFIG.TRACKER_SHEET_NAME) {
    SpreadsheetApp.getUi().alert('Please select rows inside the TikTok Tracker sheet.');
    return;
  }

  const range = sheet.getActiveRange();

  if (!range) {
    SpreadsheetApp.getUi().alert('Select one or more TikTok rows first.');
    return;
  }

  const startRow = range.getRow();
  const numRows = range.getNumRows();

  if (startRow < TIKTOK_CONFIG.TRACKER_DATA_START_ROW) {
    SpreadsheetApp.getUi().alert('Please select actual TikTok rows, not the dashboard/header rows.');
    return;
  }

  const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy');
  const now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm');
  const activeUser = Session.getActiveUser().getEmail() || '';

  let marked = 0;
  let skippedAlreadyPaid = 0;
  let addedToPaidTab = 0;
  let skippedPaidTabDuplicate = 0;

  const paidRowsForAccountant = [];

  for (let i = 0; i < numRows; i++) {
    const row = startRow + i;

    const rowValues = sheet.getRange(row, 1, 1, 24).getValues()[0];

    const person = cleanTikTokValue_(rowValues[0]);
    const caseRef = cleanTikTokValue_(rowValues[1]);
    const clientName = cleanTikTokValue_(rowValues[2]);
    const accidentDate = cleanTikTokValue_(rowValues[3]);
    const solicitor = cleanTikTokValue_(rowValues[4]);
    const type = cleanTikTokValue_(rowValues[5]);
    const solicitorRef = cleanTikTokValue_(rowValues[6]);
    const amount = Number(rowValues[12]) || TIKTOK_CONFIG.RATE_PER_TIKTOK;
    const emailSentDate = cleanTikTokValue_(rowValues[16]);
    const alreadyPaid = toBoolean_(rowValues[17]);
    const paymentWeek = cleanTikTokValue_(rowValues[19]);
    const status = cleanTikTokValue_(rowValues[20]);
    const paymentKey = cleanTikTokValue_(rowValues[22]);

    if (!person || !caseRef || status === 'Excluded' || status === 'Needs Review') {
      continue;
    }

    if (alreadyPaid) {
      skippedAlreadyPaid++;
      continue;
    }

    sheet.getRange(row, 14).setValue(0);
    sheet.getRange(row, 15).setValue('PAID');
    sheet.getRange(row, 18).setValue(true);
    sheet.getRange(row, 19).setValue(today);
    sheet.getRange(row, 21).setValue('Paid');

    marked++;

    paidRowsForAccountant.push([
      today,
      person,
      caseRef,
      clientName,
      accidentDate,
      solicitor,
      solicitorRef,
      type,
      amount,
      paymentWeek,
      emailSentDate,
      paymentKey,
      activeUser,
      now
    ]);
  }

  if (paidRowsForAccountant.length > 0) {
    const result = appendRowsToTikTokPaid_(paidRowsForAccountant);
    addedToPaidTab = result.added;
    skippedPaidTabDuplicate = result.duplicates;
  }

  SpreadsheetApp.getUi().alert(
    'Payment update complete.\n\n' +
    'Marked paid: ' + marked + '\n' +
    'Already paid skipped: ' + skippedAlreadyPaid + '\n' +
    'Added to TikTok Paid: ' + addedToPaidTab + '\n' +
    'Duplicate paid records skipped: ' + skippedPaidTabDuplicate
  );
}


/**
 * Adds paid rows into TikTok Paid tab for accountant.
 */
function appendRowsToTikTokPaid_(rows) {
  const paidSheet = getOrCreateTikTokPaidSheet_();
  const existingKeys = getExistingTikTokPaidKeys_(paidSheet);

  const rowsToAdd = [];
  let duplicateCount = 0;

  rows.forEach(row => {
    const paymentKey = cleanTikTokValue_(row[11]);

    if (!paymentKey) {
      return;
    }

    if (existingKeys[paymentKey]) {
      duplicateCount++;
      return;
    }

    rowsToAdd.push(row);
    existingKeys[paymentKey] = true;
  });

  if (rowsToAdd.length > 0) {
    paidSheet.getRange(paidSheet.getLastRow() + 1, 1, rowsToAdd.length, rowsToAdd[0].length)
      .setValues(rowsToAdd);

    formatTikTokPaidSheet_(paidSheet);
  }

  return {
    added: rowsToAdd.length,
    duplicates: duplicateCount
  };
}


/**
 * Creates TikTok Paid tab if missing.
 */
function getOrCreateTikTokPaidSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(TIKTOK_CONFIG.PAID_SHEET_NAME);

  if (!sheet) {
    sheet = ss.insertSheet(TIKTOK_CONFIG.PAID_SHEET_NAME);
  }

  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, 14).setValues([[
      'Paid Date',
      'TikTok By',
      'Case Ref',
      'Client Name',
      'Accident Date',
      'Solicitor',
      'Solicitor Ref',
      'Type',
      'Amount Paid',
      'Payment Week',
      'Email Sent Date',
      'Payment Key',
      'Marked Paid By',
      'Added To Paid Tab At'
    ]]);
  }

  formatTikTokPaidSheet_(sheet);

  return sheet;
}


/**
 * Formats TikTok Paid tab.
 */
function formatTikTokPaidSheet_(sheet) {
  const lastRow = Math.max(sheet.getLastRow(), 1);

  sheet.setFrozenRows(1);

  sheet.getRange(1, 1, 1, 14)
    .setFontWeight('bold')
    .setBackground('#1f4e79')
    .setFontColor('#ffffff')
    .setBorder(true, true, true, true, true, true);

  if (lastRow > 1) {
    sheet.getRange(2, 1, lastRow - 1, 14)
      .setBorder(true, true, true, true, true, true);

    sheet.getRange(2, 9, lastRow - 1, 1)
      .setNumberFormat('£#,##0.00');
  }

  sheet.autoResizeColumns(1, 14);

  if (sheet.getFilter()) {
    sheet.getFilter().remove();
  }

  sheet.getRange(1, 1, lastRow, 14).createFilter();
}


/**
 * Reads existing payment keys in TikTok Paid tab.
 */
function getExistingTikTokPaidKeys_(sheet) {
  const keys = {};

  if (!sheet || sheet.getLastRow() < 2) {
    return keys;
  }

  const values = sheet.getRange(2, 12, sheet.getLastRow() - 1, 1).getDisplayValues();

  values.forEach(row => {
    const key = cleanTikTokValue_(row[0]);

    if (key) {
      keys[key] = true;
    }
  });

  return keys;
}


/**
 * Nightly tracker update.
 */
function nightlyTikTokTrackerUpdate() {
  updateTikTokTracker();
}


/**
 * Payment email check.
 */
function runTikTokPaymentEmailsIfDueToday() {
  // Safety guard: scheduled triggers must never send real payment emails.
  // Real payment emails are only sent by manually running sendMonthlyTikTokPaymentEmails().
  nightlyTikTokTrackerUpdate();
}


/**
 * Create nightly update and payment check triggers.
 */
function createTikTokNightlyAndPaymentTriggers() {
  const triggers = ScriptApp.getProjectTriggers();

  triggers.forEach(trigger => {
    const fn = trigger.getHandlerFunction();

    if (
      fn === 'nightlyTikTokTrackerUpdate' ||
      fn === 'runTikTokPaymentEmailsIfDueToday'
    ) {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  ScriptApp.newTrigger('nightlyTikTokTrackerUpdate')
    .timeBased()
    .everyDays(1)
    .atHour(23)
    .create();

  SpreadsheetApp.getUi().alert(
    'TikTok triggers created.\n\nTracker update: every night around 11pm.\nNo scheduled trigger sends real payment emails. Use sendMonthlyTikTokPaymentEmails() manually for month-end emails.'
  );
}


/**
 * Remove TikTok triggers.
 */
function removeTikTokTriggers() {
  const triggers = ScriptApp.getProjectTriggers();
  let removed = 0;

  triggers.forEach(trigger => {
    const fn = trigger.getHandlerFunction();

    if (
      fn === 'nightlyTikTokTrackerUpdate' ||
      fn === 'runTikTokPaymentEmailsIfDueToday'
    ) {
      ScriptApp.deleteTrigger(trigger);
      removed++;
    }
  });

  SpreadsheetApp.getUi().alert('Removed TikTok triggers: ' + removed);
}


/**
 * Adds TikTok helpers to the spreadsheet menu.
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('TikTok')
    .addItem('Update TikTok Tracker', 'updateTikTokTracker')
    .addSeparator()
    .addItem('Show Outstanding This Run', 'filterTikTokOutstandingThisRun')
    .addItem('Clear Tracker Filters', 'clearTikTokTrackerFilters')
    .addSeparator()
    .addItem('Send Monthly Payment Emails', 'sendMonthlyTikTokPaymentEmails')
    .addItem('Send Test Payment Summary To Me', 'sendTikTokPaymentSummaryTestToMe')
    .addSeparator()
    .addItem('Mark Selected Rows As Paid', 'markSelectedTikTokRowsAsPaid')
    .addToUi();
}


/**
 * Filters the main tracker to rows requiring attention this run.
 */
function filterTikTokOutstandingThisRun() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(TIKTOK_CONFIG.TRACKER_SHEET_NAME);

  if (!sheet) {
    SpreadsheetApp.getUi().alert('TikTok Tracker not found.');
    return;
  }

  clearTikTokTrackerFilters();

  const lastRow = Math.max(sheet.getLastRow(), TIKTOK_CONFIG.TRACKER_DATA_START_ROW);
  const filter = sheet
    .getRange(TIKTOK_CONFIG.TRACKER_HEADER_ROW, 1, lastRow - TIKTOK_CONFIG.TRACKER_HEADER_ROW + 1, 24)
    .createFilter();

  const actionNeededCriteria = SpreadsheetApp.newFilterCriteria()
    .setVisibleValues([
      'READY TO EMAIL',
      'PAYMENT DUE',
      'EMAILED - AWAITING PAYMENT'
    ])
    .build();

  filter.setColumnFilterCriteria(15, actionNeededCriteria);
}


/**
 * Clears tracker filters and shows all tracker rows.
 */
function clearTikTokTrackerFilters() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(TIKTOK_CONFIG.TRACKER_SHEET_NAME);

  if (!sheet) {
    SpreadsheetApp.getUi().alert('TikTok Tracker not found.');
    return;
  }

  const filter = sheet.getFilter();

  if (filter) {
    filter.remove();
  }
}


/**
 * Reads existing tracker payment/email state before rebuild.
 */
function getExistingTikTokPaymentMap_(trackerSheet) {
  const map = {};

  if (trackerSheet && trackerSheet.getLastRow() >= 12) {
    const lastRow = trackerSheet.getLastRow();
    let headerRow = TIKTOK_CONFIG.TRACKER_HEADER_ROW;
    let dataStartRow = TIKTOK_CONFIG.TRACKER_DATA_START_ROW;
    let headers = trackerSheet.getRange(headerRow, 1, 1, 24).getDisplayValues()[0];

    if (cleanTikTokValue_(headers[0]) !== 'TikTok By') {
      headerRow = 11;
      dataStartRow = 12;
      headers = trackerSheet.getRange(headerRow, 1, 1, 24).getDisplayValues()[0];
    }

    const hasOutstandingColumns = cleanTikTokValue_(headers[13]) === 'Outstanding This Run';
    const values = trackerSheet.getRange(dataStartRow, 1, lastRow - headerRow, 24).getValues();

    values.forEach(row => {
      const person = cleanTikTokValue_(row[0]);
      const caseRef = cleanTikTokValue_(row[1]);
      const clientName = cleanTikTokValue_(row[2]);
      const paymentKey = cleanTikTokValue_(hasOutstandingColumns ? row[22] : row[20]) || makeTikTokPaymentKey_(person, caseRef, clientName);

      if (!person || !caseRef || !paymentKey) {
        return;
      }

      map[paymentKey] = {
        sendEmail: toBoolean_(hasOutstandingColumns ? row[15] : row[13]),
        emailSentDate: hasOutstandingColumns ? row[16] : row[14],
        paid: toBoolean_(hasOutstandingColumns ? row[17] : row[15]),
        paidDate: hasOutstandingColumns ? row[18] : row[16],
        paymentWeek: hasOutstandingColumns ? row[19] : row[17],
        status: hasOutstandingColumns ? row[20] : row[18],
        notes: hasOutstandingColumns ? row[21] : row[19]
      };
    });
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const paidSheet = ss.getSheetByName(TIKTOK_CONFIG.PAID_SHEET_NAME);

  if (paidSheet && paidSheet.getLastRow() >= 2) {
    const paidValues = paidSheet.getRange(2, 1, paidSheet.getLastRow() - 1, 14).getValues();

    paidValues.forEach(row => {
      const paidDate = row[0];
      const person = cleanTikTokValue_(row[1]);
      const caseRef = cleanTikTokValue_(row[2]);
      const clientName = cleanTikTokValue_(row[3]);
      const emailSentDate = row[10];
      const paymentKey = cleanTikTokValue_(row[11]) || makeTikTokPaymentKey_(person, caseRef, clientName);

      if (!paymentKey) {
        return;
      }

      map[paymentKey] = {
        sendEmail: false,
        emailSentDate: emailSentDate,
        paid: true,
        paidDate: paidDate,
        paymentWeek: row[9],
        status: 'Paid',
        notes: 'Paid record exists in TikTok Paid'
      };
    });
  }

  return map;
}



function getTikTokOutstandingThisRun_(amount, paymentWeek, status, paid) {
  const currentPaymentWeek = getNextTikTokPaymentWeekText_();
  const cleanStatus = cleanTikTokValue_(status);

  if (paid || cleanStatus === 'Paid' || cleanStatus === 'Excluded' || cleanStatus === 'Needs Review') {
    return 0;
  }

  if (cleanTikTokValue_(paymentWeek) !== currentPaymentWeek) {
    return 0;
  }

  return Number(amount) || 0;
}


function getTikTokActionNeeded_(status, outstandingThisRun, sendEmail, emailSentDate, paid) {
  const cleanStatus = cleanTikTokValue_(status);

  if (cleanStatus === 'Excluded') {
    return 'EXCLUDED';
  }

  if (cleanStatus === 'Needs Review') {
    return 'NEEDS REVIEW';
  }

  if (paid || cleanStatus === 'Paid') {
    return 'PAID';
  }

  if (emailSentDate) {
    return 'EMAILED - AWAITING PAYMENT';
  }

  if (outstandingThisRun > 0 && !sendEmail && !emailSentDate) {
    return 'PAYMENT DUE';
  }

  if (outstandingThisRun > 0 && sendEmail && !emailSentDate) {
    return 'READY TO EMAIL';
  }

  return '';
}


/**
 * Summary calculation.
 *
 * IMPORTANT:
 * The top dashboard now shows the CURRENT PAYMENT RUN only.
 * Old paid rows remain in the tracker and in TikTok Paid, but they do not keep
 * increasing the active dashboard after the payment week moves on.
 */
function getTikTokPaymentSummary_(rows) {
  const currentPaymentWeek = getNextTikTokPaymentWeekText_();
  const history = getTikTokLifetimePaidSummary_();

  const summary = {
    currentPaymentWeek: currentPaymentWeek,
    Yusuf: createEmptyPaymentSummary_(),
    Suleman: createEmptyPaymentSummary_(),
    Total: createEmptyPaymentSummary_(),
    NeedsReview: { count: 0 },
    Excluded: { count: 0 },
    History: history
  };

  rows.forEach(row => {
    const person = cleanTikTokValue_(row[0]);
    const amount = Number(row[12]) || 0;
    const outstandingThisRun = Number(row[13]) || 0;
    const actionNeeded = cleanTikTokValue_(row[14]);
    const paid = toBoolean_(row[17]);
    const status = cleanTikTokValue_(row[20]);
    const paymentWeek = cleanTikTokValue_(row[19]);

    if (status === 'Needs Review') {
      summary.NeedsReview.count++;
      return;
    }

    if (status === 'Excluded') {
      summary.Excluded.count++;
      return;
    }

    if (!summary[person]) {
      return;
    }

    if (paymentWeek !== currentPaymentWeek) {
      return;
    }

    summary[person].completed++;
    summary[person].totalOwed += amount;

    summary.Total.completed++;
    summary.Total.totalOwed += amount;

    if (paid) {
      summary[person].paidCount++;
      summary[person].paidAmount += amount;

      summary.Total.paidCount++;
      summary.Total.paidAmount += amount;
    } else {
      summary[person].balance += outstandingThisRun;

      summary.Total.balance += outstandingThisRun;

      if (actionNeeded === 'READY TO EMAIL') {
        summary[person].readyToEmailCount++;
        summary.Total.readyToEmailCount++;
      } else if (actionNeeded === 'PAYMENT DUE') {
        summary[person].paymentDueCount++;
        summary.Total.paymentDueCount++;
      } else if (actionNeeded === 'EMAILED - AWAITING PAYMENT') {
        summary[person].awaitingPaymentCount++;
        summary.Total.awaitingPaymentCount++;
      }
    }
  });

  return summary;
}


/**
 * Reads lifetime paid totals from TikTok Paid.
 */
function getTikTokLifetimePaidSummary_() {
  const history = {
    Yusuf: { totalPaid: 0, totalPaidCount: 0 },
    Suleman: { totalPaid: 0, totalPaidCount: 0 },
    totalPaid: 0,
    totalPaidCount: 0
  };

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const paidSheet = ss.getSheetByName(TIKTOK_CONFIG.PAID_SHEET_NAME);

  if (!paidSheet || paidSheet.getLastRow() < 2) {
    return history;
  }

  const values = paidSheet.getRange(2, 1, paidSheet.getLastRow() - 1, 14).getValues();

  values.forEach(row => {
    const person = cleanTikTokValue_(row[1]);
    const amount = Number(row[8]) || 0;

    if (!history[person]) {
      return;
    }

    history[person].totalPaid += amount;
    history[person].totalPaidCount++;

    history.totalPaid += amount;
    history.totalPaidCount++;
  });

  return history;
}


function createEmptyPaymentSummary_() {
  return {
    completed: 0,
    totalOwed: 0,
    paidAmount: 0,
    balance: 0,
    paidCount: 0,
    readyToEmailCount: 0,
    paymentDueCount: 0,
    awaitingPaymentCount: 0
  };
}


/**
 * Sort rows by status, then accident date, then person/case ref.
 */
function sortTikTokRows_(rows) {
  const statusOrder = {
    'Unpaid': 1,
    'Email Sent': 2,
    'Paid': 3,
    'Needs Review': 4,
    'Excluded': 5
  };

  rows.sort((a, b) => {
    const statusA = statusOrder[cleanTikTokValue_(a[20])] || 99;
    const statusB = statusOrder[cleanTikTokValue_(b[20])] || 99;

    if (statusA !== statusB) {
      return statusA - statusB;
    }

    const dateA = parseTikTokDate_(a[3]);
    const dateB = parseTikTokDate_(b[3]);

    if (dateA && dateB && dateA.getTime() !== dateB.getTime()) {
      return dateA.getTime() - dateB.getTime();
    }

    if (dateA && !dateB) return -1;
    if (!dateA && dateB) return 1;

    const personA = cleanTikTokValue_(a[0]);
    const personB = cleanTikTokValue_(b[0]);

    if (personA !== personB) {
      return personA.localeCompare(personB);
    }

    const caseA = cleanTikTokValue_(a[1]);
    const caseB = cleanTikTokValue_(b[1]);

    return caseA.localeCompare(caseB);
  });
}


/**
 * Parses UK dates like 07/04/2026 or 7/4/2026.
 */
function parseTikTokDate_(value) {
  const text = cleanTikTokValue_(value);

  if (!text) {
    return null;
  }

  const parts = text.split('/');

  if (parts.length !== 3) {
    return null;
  }

  const day = Number(parts[0]);
  const month = Number(parts[1]) - 1;
  const year = Number(parts[2]);

  if (!day || month < 0 || !year) {
    return null;
  }

  return new Date(year, month, day);
}


/**
 * Prevents duplicate payment keys from being paid accidentally.
 */
function applyDuplicatePaymentKeySafety_(rows) {
  const counts = {};

  rows.forEach(row => {
    const paymentKey = cleanTikTokValue_(row[22]);
    const status = cleanTikTokValue_(row[20]);

    if (!paymentKey || status === 'Excluded') {
      return;
    }

    counts[paymentKey] = (counts[paymentKey] || 0) + 1;
  });

  rows.forEach(row => {
    const paymentKey = cleanTikTokValue_(row[22]);
    const status = cleanTikTokValue_(row[20]);
    const paid = toBoolean_(row[17]);

    if (!paymentKey || status === 'Excluded' || paid) {
      return;
    }

    if ((counts[paymentKey] || 0) > 1) {
      row[13] = 0;
      row[14] = 'NEEDS REVIEW';
      row[20] = 'Needs Review';
      row[21] = 'Duplicate payment key found - check before paying';
      row[11] = 0;
      row[12] = 0;
    }
  });
}


/**
 * Formatting.
 */
function formatTikTokTracker_(sheet, rowCount) {
  sheet.setFrozenRows(TIKTOK_CONFIG.TRACKER_HEADER_ROW);

  sheet.getRange('A1:X1')
    .merge()
    .setFontWeight('bold')
    .setFontSize(18)
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle')
    .setBackground('#0f172a')
    .setFontColor('#ffffff');

  sheet.setRowHeight(1, 36);

  sheet.getRange('A2:E2')
    .setFontWeight('bold')
    .setBackground('#f8fafc')
    .setBorder(true, true, true, true, true, true);

  sheet.getRange('A4:D4')
    .merge()
    .setValue('YUSUF')
    .setFontWeight('bold')
    .setFontSize(14)
    .setHorizontalAlignment('center')
    .setBackground('#38761d')
    .setFontColor('#ffffff');

  sheet.getRange('E4:H4')
    .merge()
    .setValue('SULEMAN')
    .setFontWeight('bold')
    .setFontSize(14)
    .setHorizontalAlignment('center')
    .setBackground('#cc4125')
    .setFontColor('#ffffff');

  sheet.getRange('I4:L4')
    .merge()
    .setValue('TOTAL')
    .setFontWeight('bold')
    .setFontSize(14)
    .setHorizontalAlignment('center')
    .setBackground('#1f4e79')
    .setFontColor('#ffffff');

  sheet.getRange('M4:P4')
    .merge()
    .setValue('OUTSTANDING')
    .setFontWeight('bold')
    .setFontSize(14)
    .setHorizontalAlignment('center')
    .setBackground('#b45f06')
    .setFontColor('#ffffff');

  sheet.getRange('Q4:T4')
    .merge()
    .setValue('ISSUES')
    .setFontWeight('bold')
    .setFontSize(14)
    .setHorizontalAlignment('center')
    .setBackground('#666666')
    .setFontColor('#ffffff');

  sheet.getRange('U4:X4')
    .merge()
    .setValue('PAID HISTORY')
    .setFontWeight('bold')
    .setFontSize(14)
    .setHorizontalAlignment('center')
    .setBackground('#783f04')
    .setFontColor('#ffffff');


  sheet.getRange('A5:D8')
    .setBackground('#eaf7e8')
    .setBorder(true, true, true, true, true, true);

  sheet.getRange('E5:H8')
    .setBackground('#fff0df')
    .setBorder(true, true, true, true, true, true);

  sheet.getRange('I5:L8')
    .setBackground('#e7f0fa')
    .setBorder(true, true, true, true, true, true);

  sheet.getRange('M5:P8')
    .setBackground('#fff2cc')
    .setBorder(true, true, true, true, true, true);

  sheet.getRange('Q5:T8')
    .setBackground('#eeeeee')
    .setBorder(true, true, true, true, true, true);

  sheet.getRange('U5:X8')
    .setBackground('#fff2cc')
    .setBorder(true, true, true, true, true, true);

  sheet.getRange('A5:A8').setFontWeight('bold');
  sheet.getRange('E5:E8').setFontWeight('bold');
  sheet.getRange('I5:I8').setFontWeight('bold');
  sheet.getRange('M5:M8').setFontWeight('bold');
  sheet.getRange('Q5:Q8').setFontWeight('bold');
  sheet.getRange('U5:U8').setFontWeight('bold');
  sheet.getRange('Q8:Q8').setFontWeight('bold');
  sheet.getRange('S8:S8').setFontWeight('bold');

  sheet.getRange('B6:B8').setNumberFormat('£#,##0.00');
  sheet.getRange('F6:F8').setNumberFormat('£#,##0.00');
  sheet.getRange('J6:J8').setNumberFormat('£#,##0.00');
  sheet.getRange('N5:N7').setNumberFormat('£#,##0.00');
  sheet.getRange('V5:V8').setNumberFormat('£#,##0.00');

  sheet.getRange('A10:X10')
    .merge()
    .setFontWeight('bold')
    .setFontSize(14)
    .setHorizontalAlignment('center')
    .setBackground('#7f1d1d')
    .setFontColor('#ffffff')
    .setBorder(true, true, true, true, true, true);

  sheet.getRange('A11:G11')
    .setFontWeight('bold')
    .setBackground('#f4cccc')
    .setHorizontalAlignment('center')
    .setBorder(true, true, true, true, true, true);

  sheet.getRange(
    TIKTOK_CONFIG.DUE_THIS_RUN_START_ROW + 2,
    1,
    TIKTOK_CONFIG.DUE_THIS_RUN_MAX_ROWS,
    7
  )
    .setBackground('#fff2cc')
    .setBorder(true, true, true, true, true, true);

  sheet.getRange(
    TIKTOK_CONFIG.DUE_THIS_RUN_START_ROW + 2,
    4,
    TIKTOK_CONFIG.DUE_THIS_RUN_MAX_ROWS,
    1
  ).setNumberFormat('£#,##0.00');

  sheet.getRange(TIKTOK_CONFIG.TRACKER_HEADER_ROW, 1, 1, 24)
    .setFontWeight('bold')
    .setBackground('#d9ead3')
    .setBorder(true, true, true, true, true, true)
    .setHorizontalAlignment('center');

  if (rowCount > 0) {
    sheet.getRange(TIKTOK_CONFIG.TRACKER_DATA_START_ROW, 1, rowCount, 24)
      .setBorder(true, true, true, true, true, true)
      .setVerticalAlignment('middle');

    sheet.getRange(TIKTOK_CONFIG.TRACKER_DATA_START_ROW, 12, rowCount, 3)
      .setNumberFormat('£#,##0.00');

    const values = sheet.getRange(TIKTOK_CONFIG.TRACKER_DATA_START_ROW, 1, rowCount, 24).getValues();

    for (let i = 0; i < values.length; i++) {
      const rowNumber = TIKTOK_CONFIG.TRACKER_DATA_START_ROW + i;

      const person = values[i][0];
      const matchType = values[i][9];
      const outstandingThisRun = Number(values[i][13]) || 0;
      const actionNeeded = cleanTikTokValue_(values[i][14]);
      const sendEmail = toBoolean_(values[i][15]);
      const paid = toBoolean_(values[i][17]);
      const status = cleanTikTokValue_(values[i][20]);

      if (actionNeeded === 'NEEDS REVIEW') {
        sheet.getRange(rowNumber, 1, 1, 24).setBackground('#f4cccc');
        sheet.getRange(rowNumber, 15, 1, 8).setFontWeight('bold');
        continue;
      }

      if (actionNeeded === 'EXCLUDED') {
        sheet.getRange(rowNumber, 1, 1, 24).setBackground('#d9d9d9');
        sheet.getRange(rowNumber, 15, 1, 8).setFontWeight('bold');
        continue;
      }

      if (actionNeeded === 'READY TO EMAIL') {
        sheet.getRange(rowNumber, 1, 1, 24).setBackground('#cfe2f3');
      } else if (outstandingThisRun > 0 || actionNeeded === 'PAYMENT DUE' || actionNeeded === 'EMAILED - AWAITING PAYMENT') {
        sheet.getRange(rowNumber, 1, 1, 24).setBackground('#fce8b2');
      } else if (paid || actionNeeded === 'PAID') {
        sheet.getRange(rowNumber, 1, 1, 24).setBackground('#d9ead3');
      } else if (person === 'Yusuf') {
        sheet.getRange(rowNumber, 1, 1, 24).setBackground('#eaf7e8');
      } else if (person === 'Suleman') {
        sheet.getRange(rowNumber, 1, 1, 24).setBackground('#fff0df');
      }

      sheet.getRange(rowNumber, 14, 1, 2).setFontWeight('bold');
      sheet.getRange(rowNumber, 21).setFontWeight('bold');

      if (matchType === 'Near Match') {
        sheet.getRange(rowNumber, 10, 1, 2)
          .setBackground('#fff2cc')
          .setFontWeight('bold');
      }

      if (sendEmail) {
        sheet.getRange(rowNumber, 16, 1, 2)
          .setBackground('#cfe2f3')
          .setFontWeight('bold');
      }

      if (paid) {
        sheet.getRange(rowNumber, 18, 1, 2)
          .setBackground('#b6d7a8')
          .setFontWeight('bold');
      }
    }
  }

  sheet.autoResizeColumns(1, 24);

  sheet.setColumnWidth(1, 110);
  sheet.setColumnWidth(2, 90);
  sheet.setColumnWidth(3, 210);
  sheet.setColumnWidth(4, 105);
  sheet.setColumnWidth(5, 130);
  sheet.setColumnWidth(6, 140);
  sheet.setColumnWidth(7, 140);
  sheet.setColumnWidth(12, 70);
  sheet.setColumnWidth(13, 105);
  sheet.setColumnWidth(14, 165);
  sheet.setColumnWidth(15, 210);
  sheet.setColumnWidth(16, 95);
  sheet.setColumnWidth(17, 115);
  sheet.setColumnWidth(18, 75);
  sheet.setColumnWidth(19, 95);
  sheet.setColumnWidth(20, 230);
  sheet.setColumnWidth(21, 145);
  sheet.setColumnWidth(22, 260);

  try {
    sheet.hideColumns(8, 4);
    sheet.hideColumns(23, 2);
  } catch (err) {}

  if (sheet.getFilter()) {
    sheet.getFilter().remove();
  }

  const filterLastRow = Math.max(TIKTOK_CONFIG.TRACKER_DATA_START_ROW, TIKTOK_CONFIG.TRACKER_HEADER_ROW + rowCount);
  sheet.getRange(TIKTOK_CONFIG.TRACKER_HEADER_ROW, 1, filterLastRow - TIKTOK_CONFIG.TRACKER_HEADER_ROW + 1, 24).createFilter();
}


/**
 * Exclusions.
 */
function isExcludedTikTokRowByText_(rowValues) {
  const rowText = rowValues
    .map(value => cleanTikTokValue_(value))
    .join(' ')
    .toLowerCase();

  if (!rowText) {
    return false;
  }

  return TIKTOK_CONFIG.EXCLUDE_KEYWORDS.some(keyword => {
    return rowText.includes(keyword.toLowerCase());
  });
}


function getRowWideRedInfo_(rowBackgrounds) {
  let redCount = 0;

  rowBackgrounds.forEach(colour => {
    if (isRowWideRedColour_(colour)) {
      redCount++;
    }
  });

  return {
    redCount: redCount,
    isRedRow: redCount >= TIKTOK_CONFIG.ROW_WIDE_RED_CELL_LIMIT
  };
}


function isRowWideRedColour_(hexColour) {
  const colour = normaliseTikTokHex_(hexColour);

  if (!colour) {
    return false;
  }

  const exactRedColours = TIKTOK_CONFIG.ROW_WIDE_RED_COLOURS.map(normaliseTikTokHex_);

  if (exactRedColours.includes(colour)) {
    return true;
  }

  for (let i = 0; i < exactRedColours.length; i++) {
    const distance = getColourDistance_(colour, exactRedColours[i]);

    if (distance !== null && distance <= TIKTOK_CONFIG.COLOUR_TOLERANCE) {
      return true;
    }
  }

  return false;
}


/**
 * Colour detection.
 */
function getTikTokPersonFromColour_(hexColour) {
  const colour = normaliseTikTokHex_(hexColour);

  if (!colour) {
    return {
      person: '',
      matchType: '',
      distance: ''
    };
  }

  const yusufColour = normaliseTikTokHex_(TIKTOK_CONFIG.YUSUF_COLOUR);
  const sulemanColour = normaliseTikTokHex_(TIKTOK_CONFIG.SULEMAN_COLOUR);

  if (colour === yusufColour) {
    return {
      person: 'Yusuf',
      matchType: 'Exact',
      distance: 0
    };
  }

  if (colour === sulemanColour) {
    return {
      person: 'Suleman',
      matchType: 'Exact',
      distance: 0
    };
  }

  const yusufDistance = getColourDistance_(colour, yusufColour);
  const sulemanDistance = getColourDistance_(colour, sulemanColour);

  const tolerance = TIKTOK_CONFIG.COLOUR_TOLERANCE;

  if (yusufDistance !== null && yusufDistance <= tolerance) {
    return {
      person: 'Yusuf',
      matchType: 'Near Match',
      distance: Math.round(yusufDistance)
    };
  }

  if (sulemanDistance !== null && sulemanDistance <= tolerance) {
    return {
      person: 'Suleman',
      matchType: 'Near Match',
      distance: Math.round(sulemanDistance)
    };
  }

  return {
    person: '',
    matchType: '',
    distance: ''
  };
}


function isPossibleManualTikTokColour_(hexColour) {
  const colour = normaliseTikTokHex_(hexColour);

  if (!colour) {
    return false;
  }

  if (colour === '#ffffff' || colour === 'white') {
    return false;
  }

  const rgb = tikTokHexToRgb_(colour);

  if (!rgb) {
    return false;
  }

  if (rgb.r > 245 && rgb.g > 245 && rgb.b > 245) {
    return false;
  }

  return true;
}


function getColourDistance_(colourA, colourB) {
  const rgbA = tikTokHexToRgb_(colourA);
  const rgbB = tikTokHexToRgb_(colourB);

  if (!rgbA || !rgbB) {
    return null;
  }

  return Math.sqrt(
    Math.pow(rgbA.r - rgbB.r, 2) +
    Math.pow(rgbA.g - rgbB.g, 2) +
    Math.pow(rgbA.b - rgbB.b, 2)
  );
}


function tikTokHexToRgb_(hex) {
  if (!hex) {
    return null;
  }

  let cleanHex = String(hex)
    .replace('#', '')
    .trim();

  if (cleanHex.length === 3) {
    cleanHex = cleanHex
      .split('')
      .map(ch => ch + ch)
      .join('');
  }

  if (cleanHex.length !== 6) {
    return null;
  }

  const num = parseInt(cleanHex, 16);

  if (isNaN(num)) {
    return null;
  }

  return {
    r: (num >> 16) & 255,
    g: (num >> 8) & 255,
    b: num & 255
  };
}


/**
 * Payment date helpers.
 */
function getCurrentMonthTikTokPaymentRunDate_() {
  const now = new Date();
  const lastDay = new Date(now.getFullYear(), now.getMonth() + 1, 0);
  const paymentDate = new Date(lastDay);
  paymentDate.setDate(lastDay.getDate() - 7);
  return dateOnly_(paymentDate);
}


function getNextTikTokPaymentRunDate_() {
  const today = dateOnly_(new Date());
  let paymentDate = getCurrentMonthTikTokPaymentRunDate_();

  if (today.getTime() > paymentDate.getTime()) {
    const nextMonth = new Date(today.getFullYear(), today.getMonth() + 1, 1);
    const lastDay = new Date(nextMonth.getFullYear(), nextMonth.getMonth() + 1, 0);
    paymentDate = new Date(lastDay);
    paymentDate.setDate(lastDay.getDate() - 7);
  }

  return dateOnly_(paymentDate);
}


function getNextTikTokPaymentWeekText_() {
  return getPaymentWeekCommencingText_(getNextTikTokPaymentRunDate_());
}


function getPaymentWeekCommencingText_(date) {
  const weekStart = getMondayOfWeek_(date);

  return 'Week commencing ' + Utilities.formatDate(
    weekStart,
    Session.getScriptTimeZone(),
    'dd/MM/yyyy'
  );
}


function getMondayOfWeek_(date) {
  const d = dateOnly_(date);
  const day = d.getDay();
  const diff = d.getDate() - day + (day === 0 ? -6 : 1);
  return dateOnly_(new Date(d.setDate(diff)));
}


function dateOnly_(date) {
  return new Date(date.getFullYear(), date.getMonth(), date.getDate());
}


/**
 * Helpers.
 */
function buildEmptyTikTokTracker_(sheet) {
  sheet.getRange(1, 1, 1, 2).setValues([
    ['TikTok Tracker', 'No TikTok rows found']
  ]);
}


function makeTikTokPaymentKey_(person, caseRef, clientName) {
  return [
    cleanTikTokValue_(person).toLowerCase(),
    cleanTikTokValue_(caseRef).toLowerCase()
  ].join('|');
}


function showSelectedTikTokColourCodes() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const range = sheet.getActiveRange();

  if (!range) {
    SpreadsheetApp.getUi().alert('Select a coloured cell first.');
    return;
  }

  const backgrounds = range.getBackgrounds();
  const values = range.getDisplayValues();

  let message = '';

  for (let r = 0; r < backgrounds.length; r++) {
    for (let c = 0; c < backgrounds[r].length; c++) {
      const cell = range.getCell(r + 1, c + 1);
      const a1 = cell.getA1Notation();
      const value = values[r][c];
      const colour = backgrounds[r][c];

      message += a1 + ' | ' + value + ' | ' + colour + '\n';
    }
  }

  SpreadsheetApp.getUi().alert(message || 'No colour data found.');
}


function toBoolean_(value) {
  return value === true || String(value).toLowerCase() === 'true';
}


function cleanTikTokValue_(value) {
  if (value === null || value === undefined) {
    return '';
  }

  return String(value)
    .replace(/\s+/g, ' ')
    .trim();
}


function normaliseTikTokHex_(hex) {
  if (!hex) {
    return '';
  }

  return String(hex)
    .trim()
    .toLowerCase();
}


function escapeHtml_(value) {
  if (value === null || value === undefined) {
    return '';
  }

  return String(value)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#039;');
}
