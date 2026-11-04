/**
 * FLOW REALTY x Mask360 | Codename Beverly Hills | Live Campaign Dashboard
 * Google Apps Script web app, bound to the Campaign Master Sheet.
 * Reads the sheet live and writes client feedback and approvals back to it.
 */

var CONFIG = {
  SHEET_ID: '1TUeUwXPxSGD3Gh9ZoHyC08Pn_jIspX3HojbvvzOj4lI',
  AGENCY_LOGO: 'https://lh3.googleusercontent.com/d/1WXc1kEE1js2iL5Q19JQe6UbFuk3iVI1r',
  // Paste a public link to the FLOW REALTY logo here. Leave blank to show the name as text.
  CLIENT_LOGO: '',
  CLIENT_NAME: 'FLOW REALTY',
  PROJECT_NAME: 'Codename Beverly Hills',
  TABS: {
    overview: '1. Overview',
    roadmap: '2. Roadmap',
    deliverables: '3. Deliverables',
    inputs: '4. Client Inputs',
    updates: '5. Updates'
  }
};

function doGet() {
  var t = HtmlService.createTemplateFromFile('Index');
  t.config = JSON.stringify({
    agencyLogo: CONFIG.AGENCY_LOGO,
    clientLogo: CONFIG.CLIENT_LOGO,
    clientName: CONFIG.CLIENT_NAME,
    projectName: CONFIG.PROJECT_NAME
  });
  return t.evaluate()
    .setTitle(CONFIG.CLIENT_NAME + ' x Mask360 | Campaign Dashboard')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function ss_() {
  return SpreadsheetApp.openById(CONFIG.SHEET_ID);
}

function vals_(tab, range) {
  return ss_().getSheetByName(tab).getRange(range).getDisplayValues();
}

/** Everything the dashboard shows, read live from the sheet. */
function getData() {
  var o = vals_(CONFIG.TABS.overview, 'A1:D31');
  var status = {
    phase: o[5][1],
    progress: o[6][1],
    deliverables: o[7][1],
    nextMilestone: o[8][1],
    waitingOnFlow: o[9][1],
    launch: o[10][1],
    lastUpdated: o[11][1]
  };
  var keyDates = o.slice(15, 21)
    .filter(function (r) { return r[0] || r[1]; })
    .map(function (r) { return { date: r[0], milestone: r[1], status: r[2] }; });
  var team = o.slice(24, 31)
    .filter(function (r) { return r[0]; })
    .map(function (r) { return { name: r[0], role: r[1], company: r[2], email: r[3] }; });

  var roadmap = vals_(CONFIG.TABS.roadmap, 'A6:F12')
    .filter(function (r) { return r[0]; })
    .map(function (r) {
      return { phase: r[0], dates: r[1], happens: r[2], receives: r[3], needs: r[4], status: r[5] };
    });

  var deliverables = vals_(CONFIG.TABS.deliverables, 'A6:J29')
    .filter(function (r) { return r[1]; })
    .map(function (r) {
      return {
        num: r[0], name: r[1], what: r[2], phase: r[3], due: r[4],
        stage: r[5], link: r[7], comments: r[8], score: parseFloat(r[9]) || 0
      };
    });

  var inputs = [];
  var section = '';
  vals_(CONFIG.TABS.inputs, 'A6:G36').forEach(function (r) {
    if (!r[0] && r[1]) { section = r[1]; return; }
    if (r[1]) {
      inputs.push({ num: r[0], section: section, item: r[1], detail: r[2], usedFor: r[3], due: r[4], status: r[5], notes: r[6] });
    }
  });

  var upSheet = ss_().getSheetByName(CONFIG.TABS.updates);
  var last = Math.max(upSheet.getLastRow(), 6);
  var updates = upSheet.getRange('A6:D' + last).getDisplayValues()
    .filter(function (r) { return r[1]; })
    .map(function (r) { return { date: r[0], update: r[1], next: r[2], owner: r[3] }; })
    .reverse();

  return {
    status: status, keyDates: keyDates, team: team, roadmap: roadmap,
    deliverables: deliverables, inputs: inputs, updates: updates,
    fetchedAt: Utilities.formatDate(new Date(), 'Asia/Kolkata', 'd MMM yyyy, h:mm a')
  };
}

function findDeliverableRow_(num) {
  var sh = ss_().getSheetByName(CONFIG.TABS.deliverables);
  var nums = sh.getRange('A6:A29').getDisplayValues();
  for (var i = 0; i < nums.length; i++) {
    if (String(nums[i][0]) === String(num)) return { sheet: sh, row: i + 6 };
  }
  throw new Error('Deliverable ' + num + ' not found');
}

function clean_(s, max) {
  return String(s || '').replace(/[\r\n]+/g, ' ').trim().slice(0, max);
}

function now_(fmt) {
  return Utilities.formatDate(new Date(), 'Asia/Kolkata', fmt || 'd MMM, h:mm a');
}

function stampLastUpdated_(ss) {
  ss.getSheetByName(CONFIG.TABS.overview).getRange('B12').setValue("'" + now_('d MMM yyyy, h:mm a'));
}

function logUpdate_(text, next, owner, ss) {
  ss = ss || ss_();
  var sh = ss.getSheetByName(CONFIG.TABS.updates);
  var row = Math.max(sh.getLastRow() + 1, 6);
  sh.getRange(row, 1, 1, 4).setValues([["'" + now_(), text, next || '', owner || 'Mask360']]);
  stampLastUpdated_(ss);
}

/**
 * Self-updating log. Runs automatically whenever anyone edits the sheet:
 * stage, status and link changes are written to the Updates tab with the time,
 * and the Overview "Last updated" stamp refreshes. A note typed into a new row
 * of the Updates tab gets its date and owner filled in by itself.
 */
function onEdit(e) {
  if (!e || !e.range) return;
  var ss = e.source;
  var sh = e.range.getSheet();
  var name = sh.getName();
  var row = e.range.getRow();
  var col = e.range.getColumn();
  if (e.range.getNumRows() > 1 || e.range.getNumColumns() > 1) { stampLastUpdated_(ss); return; }
  var val = e.range.getDisplayValue();

  if (name === CONFIG.TABS.updates) {
    if (row >= 6 && col === 2 && val) {
      if (!sh.getRange(row, 1).getDisplayValue()) sh.getRange(row, 1).setValue("'" + now_());
      if (!sh.getRange(row, 4).getDisplayValue()) sh.getRange(row, 4).setValue('Mask360');
    }
    stampLastUpdated_(ss);
    return;
  }

  if (name === CONFIG.TABS.deliverables && row >= 6 && row <= 29) {
    var title = sh.getRange(row, 2).getDisplayValue();
    if (col === 6 && val) logUpdate_(title + ' is now ' + val, '', 'Mask360', ss);
    else if (col === 8 && val) logUpdate_(title + ' shared: link added', 'FLOW to review', 'Mask360', ss);
    else stampLastUpdated_(ss);
    return;
  }

  if (name === CONFIG.TABS.roadmap && row >= 6 && col === 6 && val) {
    logUpdate_('Phase ' + sh.getRange(row, 1).getDisplayValue() + ' is now ' + val, '', 'Mask360', ss);
    return;
  }

  if (name === CONFIG.TABS.inputs && row >= 6 && col === 6 && val) {
    logUpdate_('Client input "' + sh.getRange(row, 2).getDisplayValue() + '" marked ' + val, '', 'FLOW', ss);
    return;
  }

  if (name === CONFIG.TABS.overview && row >= 16 && row <= 21 && col === 3 && val) {
    logUpdate_('Milestone "' + sh.getRange(row, 2).getDisplayValue() + '" marked ' + val, '', 'Mask360', ss);
    return;
  }

  stampLastUpdated_(ss);
}

/** Client leaves feedback on one deliverable. */
function submitFeedback(num, name, text) {
  name = clean_(name, 60) || 'FLOW';
  text = clean_(text, 1000);
  if (!text) throw new Error('Feedback is empty');
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var f = findDeliverableRow_(num);
    var cell = f.sheet.getRange(f.row, 9);
    var stamp = Utilities.formatDate(new Date(), 'Asia/Kolkata', 'd MMM, h:mm a');
    var entry = '[' + stamp + ', ' + name + '] ' + text;
    var prev = cell.getDisplayValue();
    cell.setValue(prev ? prev + '\n' + entry : entry);
    var title = f.sheet.getRange(f.row, 2).getDisplayValue();
    logUpdate_('Feedback from ' + name + ' on ' + title + ': ' + text, 'Mask360 to review and respond', 'Mask360');
  } finally {
    lock.releaseLock();
  }
  return getData();
}

/** Client approves a deliverable that is with them for review. */
function approveDeliverable(num, name) {
  name = clean_(name, 60) || 'FLOW';
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var f = findDeliverableRow_(num);
    var stageCell = f.sheet.getRange(f.row, 6);
    if (stageCell.getDisplayValue() !== 'With FLOW for review') {
      throw new Error('Only items with FLOW for review can be approved');
    }
    stageCell.setValue('Approved');
    var title = f.sheet.getRange(f.row, 2).getDisplayValue();
    logUpdate_(title + ' approved by ' + name, 'Mask360 to finalise and deliver', 'Mask360');
  } finally {
    lock.releaseLock();
  }
  return getData();
}
