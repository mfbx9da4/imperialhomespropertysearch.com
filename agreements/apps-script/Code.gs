/**
 * Imperial Homes client agreements.
 *
 * Lives inside the "Imperial Homes Agreements" Google Sheet (Extensions > Apps Script).
 * - Sophie uses Imperial Homes > New agreement (Form.html), or copies an earlier client's row,
 *   previews the client page, then sends. Every agreement is one row on the Agreements tab.
 * - The client page (imperialhomespropertysearch.co.uk/b/?CODE) calls doGet / doPost.
 * - Emails go through the Cloudflare mail relay, which sends via BT SMTP.
 *
 * Script properties (Project Settings > Script properties), set by "Set up":
 *   RELAY_URL     https://imperial-homes-mailer.<account>.workers.dev
 *   RELAY_KEY     same random string as the relay's RELAY_KEY secret
 *   NOTIFY_EMAIL  where "client has accepted" emails go
 *   SITE_URL      https://imperialhomespropertysearch.co.uk
 */

var TABS = { agreements: 'Agreements', terms: 'Terms', template: 'Template' };

var AGREEMENT_HEADERS = [
  'Client name', 'Email', 'Phone', 'Brief', 'Success fee', 'Deposit',
  'Code', 'Link', 'Status', 'Terms version', 'Sent at', 'Viewed at',
  'Accepted at', 'Accepted name', 'Browser', 'Signed PDF', 'Last error',
];

// Columns Sophie fills in. Everything after these is written by the script.
var REQUIRED = ['Client name', 'Email', 'Brief', 'Success fee', 'Deposit'];

var DEFAULT_TEMPLATE = [
  ['Invite subject', 'Your property search brief from Imperial Homes'],
  ['Invite email',
    'Dear {Client Name},\n\n' +
    'Thank you for instructing Imperial Homes to assist with your property search.\n\n' +
    'Please review your property search brief and our terms, and accept them here:\n{Link}\n\n' +
    'Kind regards,\nImperial Homes Property Search'],
  ['Page letter',
    'Dear {Client Name},\n\n' +
    'Thank you for instructing Imperial Homes to assist with your property search.\n\n' +
    'Please find below your property search brief, which summarises your requirements and ' +
    'preferences based on our initial discussions.\n\n' +
    'Please take a moment to review the details and let us know if anything needs amending. ' +
    'We want to ensure we have a clear understanding of your requirements before commencing the search.\n\n' +
    'We look forward to assisting you in finding your next property.\n\n' +
    'Kind regards,\nImperial Homes Property Search'],
  ['Confirmation subject', 'Your Imperial Homes agreement is confirmed'],
  ['Confirmation email',
    'Dear {Client Name},\n\n' +
    'Thank you for accepting your property search brief and our terms. ' +
    'A copy of what you agreed to is attached for your records.\n\n' +
    'To begin your search, please pay the deposit of {Deposit} by bank transfer:\n\n' +
    '{Bank Details}\n' +
    'Payment reference: {Reference}\n\n' +
    'For your security, we will never change our bank details by email. ' +
    'If you are unsure, please call us to confirm them before you pay.\n\n' +
    'We will be in touch shortly to begin your search.\n\n' +
    'Kind regards,\nImperial Homes Property Search'],
  ['Bank details',
    'ADD BANK DETAILS HERE, for example:\n' +
    'Account name: Imperial Homes Property Search\nSort code: 00-00-00\nAccount number: 00000000'],
  ['Notify subject', '{Client Name} has accepted their agreement'],
  // Starting values for the New agreement form. Sophie changes them for each client as needed.
  ['Brief template',
    'Location:\nProperty type:\nBedrooms:\nBudget:\nMust haves:\nNice to haves:\nTimescale:'],
  ['Default success fee', ''],
  ['Default deposit', ''],
  ['WhatsApp message',
    'Hi {First Name}, this is Imperial Homes. Here is your property search brief and terms. ' +
    'Please review and accept here: {Link}'],
];

var CODE_ALPHABET = '23456789ABCDEFGHJKMNPQRSTUVWXYZ'; // no 0/O, 1/I/L
var CODE_LENGTH = 6;
var TIMEZONE = 'Europe/London';

// ---------------------------------------------------------------------------
// Sheet menu
// ---------------------------------------------------------------------------

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Imperial Homes')
    .addItem('New agreement', 'openNewAgreement')
    .addItem('New agreement, copied from selected row', 'openCopyOfSelected')
    .addSeparator()
    .addItem('Send agreement (selected row)', 'sendSelected')
    .addItem('WhatsApp message (selected row)', 'whatsappSelected')
    .addItem('Cancel agreement (selected row)', 'cancelSelected')
    .addSeparator()
    .addItem('Set up', 'setup')
    .addItem('Send a test email to me', 'testEmail')
    .addToUi();
}

function setup() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();

  var agreements = ss.getSheetByName(TABS.agreements) || ss.insertSheet(TABS.agreements);
  // Plain text everywhere Sophie types, so "£1,500" and phone numbers stay exactly as typed.
  agreements.getRange('A:Q').setNumberFormat('@');
  if (agreements.getLastRow() === 0) {
    agreements.appendRow(AGREEMENT_HEADERS);
    agreements.setFrozenRows(1);
    agreements.getRange(1, 1, 1, AGREEMENT_HEADERS.length).setFontWeight('bold');
  }

  var terms = ss.getSheetByName(TABS.terms) || ss.insertSheet(TABS.terms);
  if (terms.getLastRow() === 0) {
    terms.appendRow(['Version', 'Active', 'Text']);
    terms.appendRow(['v1', true, 'PASTE THE TERMS AND CONDITIONS HERE']);
    terms.setFrozenRows(1);
    terms.getRange(1, 1, 1, 3).setFontWeight('bold');
  }

  var template = ss.getSheetByName(TABS.template) || ss.insertSheet(TABS.template);
  template.getRange('B:B').setNumberFormat('@');
  if (template.getLastRow() === 0) {
    template.appendRow(['Name', 'Text']);
    DEFAULT_TEMPLATE.forEach(function (row) { template.appendRow(row); });
    template.setFrozenRows(1);
    template.getRange(1, 1, 1, 2).setFontWeight('bold');
  }

  var props = PropertiesService.getScriptProperties();
  var defaults = {
    RELAY_URL: '',
    RELAY_KEY: '',
    NOTIFY_EMAIL: 'imperialhomes@btinternet.com',
    SITE_URL: 'https://imperialhomespropertysearch.co.uk',
  };
  Object.keys(defaults).forEach(function (key) {
    var current = props.getProperty(key) || defaults[key];
    var answer = ui.prompt('Set up', key + (current ? ' (leave blank to keep: ' + mask(key, current) + ')' : ''),
      ui.ButtonSet.OK_CANCEL);
    if (answer.getSelectedButton() !== ui.Button.OK) return;
    var value = answer.getResponseText().trim() || current;
    props.setProperty(key, value);
  });
  ui.alert('Set up is done. Now deploy the web app (Deploy > New deployment) if you have not yet.');
}

// --- New agreement form (Form.html) ---------------------------------------

function openNewAgreement() {
  showForm(formDefaults());
}

/** Most clients want something close to an earlier client: start from that row and tweak it. */
function openCopyOfSelected() {
  withSelectedRow(function (row) {
    showForm({
      clientName: '',
      email: '',
      phone: '',
      brief: String(row.get('Brief')),
      successFee: String(row.get('Success fee')),
      deposit: String(row.get('Deposit')),
      copiedFrom: String(row.get('Client name')),
    });
  });
}

function showForm(prefill) {
  var template = HtmlService.createTemplateFromFile('Form');
  template.prefillJson = JSON.stringify(prefill);
  SpreadsheetApp.getUi().showModalDialog(template.evaluate().setWidth(640).setHeight(720), 'New agreement');
}

function formDefaults() {
  return {
    clientName: '',
    email: '',
    phone: '',
    brief: templateTextOr('Brief template', ''),
    successFee: templateTextOr('Default success fee', ''),
    deposit: templateTextOr('Default deposit', ''),
    copiedFrom: '',
  };
}

/**
 * Called by Form.html. Adds the row, then either keeps it as a draft (so Sophie can preview the
 * client page) or sends it. Returns what the form shows next.
 */
function createAgreement(form, send) {
  var values = {
    'Client name': String(form.clientName || '').trim(),
    'Email': String(form.email || '').trim(),
    'Phone': String(form.phone || '').trim(),
    'Brief': String(form.brief || '').trim(),
    'Success fee': String(form.successFee || '').trim(),
    'Deposit': String(form.deposit || '').trim(),
  };
  checkRequired(values);

  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(TABS.agreements);
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(String);
  // Plain text, so Sheets keeps "£1,500" and "07700 900123" exactly as typed.
  var range = sheet.getRange(sheet.getLastRow() + 1, 1, 1, headers.length).setNumberFormat('@');
  range.setValues([headers.map(function (h) { return values[h] === undefined ? '' : values[h]; })]);
  var row = new Row(sheet, range.getRow());
  assignLink(row);
  row.set('Status', 'draft');
  if (send) sendInvite(row);
  return result(row);
}

/** Called by Form.html after a preview: send the draft it just made. */
function sendDraft(code) {
  var row = findRow(code);
  if (!row) throw new Error('Could not find that agreement. Check the Agreements tab.');
  sendInvite(row);
  return result(row);
}

function result(row) {
  var wa = whatsapp(row);
  return {
    code: row.get('Code'),
    link: row.get('Link'),
    status: row.get('Status'),
    email: row.get('Email'),
    whatsappText: wa.text,
    whatsappUrl: wa.url,
  };
}

function sendSelected() {
  withSelectedRow(function (row) {
    var link = sendInvite(row);
    SpreadsheetApp.getUi().alert('Sent to ' + row.get('Email') + '\n\nLink: ' + link);
  });
}

function whatsappSelected() {
  withSelectedRow(function (row) {
    if (!row.get('Code')) throw new Error('Send the agreement first, so it has a link.');
    var wa = whatsapp(row);
    var html = HtmlService.createHtmlOutput(
      '<p style="font-family:sans-serif">Copy this message, or open WhatsApp with it ready to send:</p>' +
      '<textarea style="width:100%;height:110px;font-family:sans-serif">' + escapeHtml(wa.text) + '</textarea>' +
      '<p><a target="_blank" href="' + escapeHtml(wa.url) + '">Open WhatsApp</a></p>'
    ).setWidth(420).setHeight(240);
    SpreadsheetApp.getUi().showModalDialog(html, 'WhatsApp message');
  });
}

function whatsapp(row) {
  var text = fill(templateText('WhatsApp message'), placeholders(row));
  // UK numbers: 07700 900123 becomes 447700900123. No number: WhatsApp asks who to send to.
  var phone = String(row.get('Phone') || '').replace(/[^0-9]/g, '').replace(/^0/, '44');
  var url = (phone ? 'https://wa.me/' + phone : 'https://wa.me/') + '?text=' + encodeURIComponent(text);
  return { text: text, url: url };
}

function cancelSelected() {
  withSelectedRow(function (row) {
    if (row.get('Status') === 'accepted') throw new Error('This agreement is already accepted. It cannot be cancelled here.');
    row.set('Status', 'cancelled');
    SpreadsheetApp.getUi().alert('Cancelled. The link now shows that the agreement is no longer available.');
  });
}

function testEmail() {
  var me = Session.getActiveUser().getEmail() || prop('NOTIFY_EMAIL');
  sendMail({ to: me, subject: 'Imperial Homes test email', text: 'If you can read this, sending works.' });
  SpreadsheetApp.getUi().alert('Test email sent to ' + me);
}

function withSelectedRow(fn) {
  var ui = SpreadsheetApp.getUi();
  var sheet = SpreadsheetApp.getActiveSheet();
  if (sheet.getName() !== TABS.agreements) {
    ui.alert('Select a row on the ' + TABS.agreements + ' tab first.');
    return;
  }
  var index = sheet.getActiveRange().getRow();
  if (index < 2) {
    ui.alert('Select a client row, not the header.');
    return;
  }
  try {
    fn(new Row(sheet, index));
  } catch (err) {
    ui.alert(err.message);
  }
}

// ---------------------------------------------------------------------------
// Core actions (also used by the tests)
// ---------------------------------------------------------------------------

function sendInvite(row) {
  checkRequired({
    'Client name': row.get('Client name'), 'Email': row.get('Email'), 'Brief': row.get('Brief'),
    'Success fee': row.get('Success fee'), 'Deposit': row.get('Deposit'),
  });
  var status = row.get('Status');
  if (status === 'accepted') throw new Error('This agreement is already accepted.');
  if (status === 'cancelled') throw new Error('This agreement is cancelled. Make a new agreement instead.');

  // Check now, so a client never accepts and then gets an email without bank details.
  if (/ADD BANK DETAILS/.test(templateText('Bank details'))) {
    throw new Error('Add the bank details on the ' + TABS.template + ' tab first.');
  }
  // A draft may have waited a while: send it with the terms that are active now.
  if (status === 'draft') row.set('Terms version', activeTerms().version);
  var link = assignLink(row);

  var vars = placeholders(row);
  try {
    sendMail({
      to: row.get('Email'),
      subject: fill(templateText('Invite subject'), vars),
      text: fill(templateText('Invite email'), vars),
    });
  } catch (err) {
    row.set('Last error', 'Invite: ' + err.message);
    throw err;
  }
  row.set('Last error', '');
  if (!status || status === 'draft' || status === 'sent') row.set('Status', 'sent');
  row.set('Sent at', now());
  return link;
}

function checkRequired(values) {
  var missing = REQUIRED.filter(function (h) { return !String(values[h] || '').trim(); });
  if (missing.length) throw new Error('Please fill in: ' + missing.join(', '));
  if (!isEmail(values['Email'])) throw new Error('The email address does not look right: ' + values['Email']);
}

/** Gives the row its code, terms version and link, if it does not have them yet. */
function assignLink(row) {
  if (!row.get('Code')) {
    row.set('Code', newCode(row.sheet));
    row.set('Terms version', activeTerms().version);
  }
  var link = prop('SITE_URL').replace(/\/$/, '') + '/b/?' + row.get('Code');
  row.set('Link', link);
  return link;
}

/** Client page: GET ?code=XXXXXX returns the agreement to show. */
function doGet(e) {
  return respond(function () {
    var row = findRow(e && e.parameter && e.parameter.code);
    if (!row) return { status: 'not_found' };
    var status = row.get('Status');
    if (status === 'cancelled') return { status: 'cancelled' };
    if (status === 'sent') {
      row.set('Status', 'viewed');
      row.set('Viewed at', now());
    }
    return view(row);
  });
}

/** Client page: POST {code, name, agree, browser} accepts the agreement. */
function doPost(e) {
  return respond(function () {
    var body = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    return accept(body);
  });
}

function accept(body) {
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var row = findRow(body.code);
    if (!row) return { status: 'not_found' };
    var status = row.get('Status');
    if (status === 'cancelled') return { status: 'cancelled' };
    if (status === 'accepted') return view(row); // Double click: already done.
    if (status === 'draft') return { status: 'error', error: 'This is a preview. It can be accepted after it is sent.' };
    if (status !== 'sent' && status !== 'viewed') return { status: 'not_found' };

    var name = String(body.name || '').trim().slice(0, 120);
    if (!name) return { status: 'error', error: 'Please type your full name.' };
    if (body.agree !== true) return { status: 'error', error: 'Please tick the box to confirm you agree.' };

    var acceptedAt = now();
    var browser = String(body.browser || '').slice(0, 300);
    var pdf = makePdf(row, name, acceptedAt, browser);

    row.set('Status', 'accepted');
    row.set('Accepted at', acceptedAt);
    row.set('Accepted name', safeCell(name));
    row.set('Browser', safeCell(browser));
    row.set('Signed PDF', pdf.url);
    row.protect();

    // The acceptance is saved above, so an email failure must not undo it.
    var vars = placeholders(row);
    var attachment = { filename: pdf.filename, content: pdf.base64, mimeType: 'application/pdf' };
    var errors = [];
    try {
      sendMail({
        to: row.get('Email'),
        subject: fill(templateText('Confirmation subject'), vars),
        text: fill(templateText('Confirmation email'), vars),
        attachments: [attachment],
      });
    } catch (err) {
      errors.push('Confirmation: ' + err.message);
    }
    try {
      sendMail({
        to: prop('NOTIFY_EMAIL'),
        subject: fill(templateText('Notify subject'), vars),
        text: name + ' accepted the agreement for ' + row.get('Client name') + ' (' + row.get('Email') + ') at ' +
          acceptedAt + '.\n\nSuccess fee: ' + row.get('Success fee') + '\nDeposit: ' + row.get('Deposit') +
          '\nTerms version: ' + row.get('Terms version') + '\nSigned copy: ' + pdf.url,
        attachments: [attachment],
      });
    } catch (err) {
      errors.push('Notify: ' + err.message);
    }
    row.set('Last error', errors.join(' | '));
    return view(row);
  } finally {
    lock.releaseLock();
  }
}

function view(row) {
  var terms = termsByVersion(row.get('Terms version'));
  return {
    status: row.get('Status'),
    clientName: row.get('Client name'),
    letter: fill(templateText('Page letter'), placeholders(row)),
    brief: String(row.get('Brief')),
    successFee: String(row.get('Success fee')),
    deposit: String(row.get('Deposit')),
    termsVersion: terms.version,
    terms: terms.text,
    acceptedAt: row.get('Accepted at') || null,
    acceptedName: row.get('Accepted name') || null,
  };
}

// ---------------------------------------------------------------------------
// Signed copy
// ---------------------------------------------------------------------------

function agreementHtml(row, name, acceptedAt, browser) {
  var p = function (s) { return escapeHtml(String(s)).replace(/\n/g, '<br>'); };
  var terms = termsByVersion(row.get('Terms version'));
  return '<html><head><meta charset="utf-8"><style>' +
    'body{font-family:Helvetica,Arial,sans-serif;font-size:11pt;color:#222;line-height:1.45}' +
    'h1{color:#002366;font-size:18pt}h2{color:#002366;font-size:13pt;margin-top:20px}' +
    'table{border-collapse:collapse}td{padding:4px 12px 4px 0;vertical-align:top}' +
    '.box{border:1px solid #ccc;padding:10px}</style></head><body>' +
    '<h1>Imperial Homes Property Search</h1>' +
    '<p>' + p(fill(templateText('Page letter'), placeholders(row))) + '</p>' +
    '<h2>Property search brief</h2><div class="box">' + p(row.get('Brief')) + '</div>' +
    '<h2>Fees</h2><table><tr><td><b>Success fee</b></td><td>' + p(row.get('Success fee')) + '</td></tr>' +
    '<tr><td><b>Deposit</b></td><td>' + p(row.get('Deposit')) + '</td></tr></table>' +
    '<h2>Terms and conditions (version ' + p(terms.version) + ')</h2><div class="box">' + p(terms.text) + '</div>' +
    '<h2>Acceptance</h2><table>' +
    '<tr><td><b>Accepted by</b></td><td>' + p(name) + '</td></tr>' +
    '<tr><td><b>Client</b></td><td>' + p(row.get('Client name')) + ' (' + p(row.get('Email')) + ')</td></tr>' +
    '<tr><td><b>Accepted at</b></td><td>' + p(acceptedAt) + ' (UK time)</td></tr>' +
    '<tr><td><b>Reference</b></td><td>' + p(row.get('Code')) + '</td></tr>' +
    '<tr><td><b>Browser</b></td><td>' + p(browser) + '</td></tr></table>' +
    '<p>The client typed their name and ticked "I have read and agree to the property search brief, ' +
    'the fees and the terms and conditions above" before clicking Accept.</p>' +
    '</body></html>';
}

function makePdf(row, name, acceptedAt, browser) {
  var filename = 'Imperial Homes agreement ' + row.get('Code') + ' ' + row.get('Client name') + '.pdf';
  var blob = Utilities.newBlob(agreementHtml(row, name, acceptedAt, browser), 'text/html', 'agreement.html')
    .getAs('application/pdf')
    .setName(filename);
  var file = signedFolder().createFile(blob);
  return { url: file.getUrl(), filename: filename, base64: Utilities.base64Encode(blob.getBytes()) };
}

function signedFolder() {
  var props = PropertiesService.getScriptProperties();
  var id = props.getProperty('SIGNED_FOLDER_ID');
  if (id) {
    try {
      return DriveApp.getFolderById(id);
    } catch (err) {
      // Folder was deleted. Make a new one below.
    }
  }
  var folder = DriveApp.createFolder('Imperial Homes signed agreements');
  props.setProperty('SIGNED_FOLDER_ID', folder.getId());
  return folder;
}

// ---------------------------------------------------------------------------
// Email via the BT relay
// ---------------------------------------------------------------------------

function sendMail(msg) {
  var html = msg.html || textToHtml(msg.text);
  var res = UrlFetchApp.fetch(prop('RELAY_URL'), {
    method: 'post',
    contentType: 'application/json',
    headers: { Authorization: 'Bearer ' + prop('RELAY_KEY') },
    payload: JSON.stringify({
      to: msg.to,
      subject: msg.subject,
      text: msg.text,
      html: html,
      attachments: msg.attachments,
    }),
    muteHttpExceptions: true,
  });
  if (res.getResponseCode() !== 200) {
    throw new Error('email failed (' + res.getResponseCode() + '): ' + res.getContentText().slice(0, 200));
  }
}

function textToHtml(text) {
  var body = escapeHtml(String(text || ''))
    .replace(/(https?:\/\/[^\s<]+)/g, '<a href="$1">$1</a>')
    .replace(/\n/g, '<br>');
  return '<div style="font-family:Helvetica,Arial,sans-serif;font-size:15px;line-height:1.5;color:#222">' +
    body + '</div>';
}

// ---------------------------------------------------------------------------
// Sheet helpers
// ---------------------------------------------------------------------------

/** A row on the Agreements tab, read and written by header name. */
function Row(sheet, index) {
  this.sheet = sheet;
  this.index = index;
  var width = sheet.getLastColumn();
  this.headers = sheet.getRange(1, 1, 1, width).getValues()[0].map(String);
  this.values = sheet.getRange(index, 1, 1, width).getValues()[0];
}

Row.prototype.col = function (header) {
  var i = this.headers.indexOf(header);
  if (i < 0) throw new Error('Column "' + header + '" is missing on the ' + TABS.agreements + ' tab.');
  return i;
};

Row.prototype.get = function (header) {
  var v = this.values[this.col(header)];
  return v === null || v === undefined ? '' : v;
};

Row.prototype.set = function (header, value) {
  var i = this.col(header);
  this.values[i] = value;
  // Plain text, or Sheets turns "2026-11-10 13:07:53" into a date and "£1,500" into 1500.
  this.sheet.getRange(this.index, i + 1).setNumberFormat('@').setValue(value);
};

/** Warn before anyone edits an accepted row. The signed PDF is the real record. */
Row.prototype.protect = function () {
  this.sheet.getRange(this.index, 1, 1, this.headers.length).protect()
    .setDescription('Accepted agreement ' + this.get('Code'))
    .setWarningOnly(true);
};

function findRow(code) {
  code = String(code || '').trim().toUpperCase();
  if (!new RegExp('^[' + CODE_ALPHABET + ']{' + CODE_LENGTH + '}$').test(code)) return null;
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(TABS.agreements);
  var data = sheet.getDataRange().getValues();
  var col = data[0].map(String).indexOf('Code');
  for (var r = 1; r < data.length; r++) {
    if (String(data[r][col]).toUpperCase() === code) return new Row(sheet, r + 1);
  }
  return null;
}

function newCode(sheet) {
  var data = sheet.getDataRange().getValues();
  var col = data[0].map(String).indexOf('Code');
  var used = {};
  for (var r = 1; r < data.length; r++) used[String(data[r][col])] = true;
  for (var attempt = 0; attempt < 20; attempt++) {
    var bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, Utilities.getUuid());
    var code = '';
    for (var i = 0; i < CODE_LENGTH; i++) code += CODE_ALPHABET[(bytes[i] + 256) % CODE_ALPHABET.length];
    if (!used[code]) return code;
  }
  throw new Error('Could not make a unique code. Try again.');
}

function activeTerms() {
  var rows = termsRows().filter(function (t) { return t.active; });
  if (rows.length !== 1) throw new Error('Tick exactly one row as Active on the ' + TABS.terms + ' tab.');
  if (/PASTE THE TERMS/.test(rows[0].text)) throw new Error('Add the real terms on the ' + TABS.terms + ' tab first.');
  return rows[0];
}

function termsByVersion(version) {
  var match = termsRows().filter(function (t) { return t.version === String(version); })[0];
  if (!match) throw new Error('Terms version ' + version + ' is missing from the ' + TABS.terms + ' tab.');
  return match;
}

function termsRows() {
  var data = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(TABS.terms).getDataRange().getValues();
  return data.slice(1).map(function (r) {
    return { version: String(r[0]), active: r[1] === true || String(r[1]).toUpperCase() === 'TRUE', text: String(r[2]) };
  });
}

function templateText(name) {
  var data = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(TABS.template).getDataRange().getValues();
  for (var r = 1; r < data.length; r++) {
    if (String(data[r][0]) === name) return String(data[r][1]);
  }
  throw new Error('"' + name + '" is missing from the ' + TABS.template + ' tab.');
}

function templateTextOr(name, fallback) {
  try {
    return templateText(name);
  } catch (err) {
    return fallback;
  }
}

function placeholders(row) {
  var name = String(row.get('Client name')).trim();
  return {
    'Client Name': name,
    'First Name': name.split(/\s+/)[0],
    'Link': row.get('Link'),
    'Deposit': row.get('Deposit'),
    'Success Fee': row.get('Success fee'),
    'Reference': row.get('Code'),
    // A function, so the Template tab is only read when a text uses it.
    'Bank Details': function () { return templateText('Bank details'); },
  };
}

function fill(text, vars) {
  return String(text).replace(/\{([A-Za-z ]+)\}/g, function (all, key) {
    if (!Object.prototype.hasOwnProperty.call(vars, key)) return all;
    return typeof vars[key] === 'function' ? vars[key]() : vars[key];
  });
}

// ---------------------------------------------------------------------------
// Small utilities
// ---------------------------------------------------------------------------

function respond(fn) {
  var out;
  try {
    out = fn();
  } catch (err) {
    console.error(err);
    out = { status: 'error', error: 'Something went wrong. Please try again, or contact Imperial Homes.' };
  }
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}

function prop(key) {
  var value = PropertiesService.getScriptProperties().getProperty(key);
  if (!value) throw new Error('Script property ' + key + ' is not set. Run Imperial Homes > Set up.');
  return value;
}

function now() {
  return Utilities.formatDate(new Date(), TIMEZONE, 'yyyy-MM-dd HH:mm:ss');
}

function isEmail(s) {
  return /^[^\s@<>]+@[^\s@<>]+\.[^\s@<>]+$/.test(String(s).trim());
}

// Stop text typed by the client from running as a spreadsheet formula.
function safeCell(s) {
  return /^[=+\-@]/.test(s) ? "'" + s : s;
}

function escapeHtml(s) {
  return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}

function mask(key, value) {
  return key === 'RELAY_KEY' ? value.slice(0, 4) + '...' : value;
}
