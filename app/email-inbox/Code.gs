/**
 * Farm Fleet Expenses — email inbox
 *
 * Paste this whole file into a new Google Apps Script project (script.google.com)
 * owned by the Gmail / Google Workspace account invoices get emailed to, then follow
 * the steps in app/README.md ("Emailing invoices in"). In short:
 *
 *   1. Run `setup` once (pick it in the function dropdown, click Run, approve access).
 *      The execution log prints the address to email invoices to and the access key.
 *   2. Deploy → New deployment → Web app — Execute as: Me, Who has access: Anyone.
 *   3. Paste the web app URL and the access key into the app: Settings → Email inbox.
 *
 * Any email delivered to you+invoices@your-domain (see ADDRESS_TAG) with a PDF or
 * photo attached shows up in the app's Expenses → Email Inbox. Nothing is copied
 * anywhere — the app asks this script for the attachment, reads it in the browser,
 * and tells the script once it's been saved (or dismissed) so it isn't offered again.
 */

const SETTINGS = {
  // Invoices go to <your address with +ADDRESS_TAG>, e.g. jane+invoices@example.com.
  // Gmail and Google Workspace deliver plus-addresses to the normal inbox.
  ADDRESS_TAG: 'invoices',
  // Set this instead to use a specific address (e.g. a Workspace alias or group
  // that delivers to this account). Leave blank to use the plus-address above.
  ADDRESS_OVERRIDE: '',
  // Only look this far back. Older emails are never offered.
  LOOKBACK_DAYS: 120,
  // Images smaller than this are skipped — they're almost always email-signature
  // logos and social icons, not invoice photos.
  MIN_IMAGE_BYTES: 20000,
  // Attachments larger than this are skipped (too big to send to the browser).
  MAX_BYTES: 15 * 1024 * 1024,
  // Gmail label added to a thread once its invoices are imported (just for you to
  // see in Gmail — the script doesn't depend on it).
  DONE_LABEL: 'Fleet invoices/Imported',
  MAX_THREADS: 50,
};

const DONE_PREFIX = 'done:';

function setup() {
  const props = PropertiesService.getScriptProperties();
  let key = props.getProperty('ACCESS_KEY');
  if (!key) {
    key = (Utilities.getUuid() + Utilities.getUuid()).replace(/-/g, '');
    props.setProperty('ACCESS_KEY', key);
  }
  GmailApp.getUserLabelByName(SETTINGS.DONE_LABEL) || GmailApp.createLabel(SETTINGS.DONE_LABEL);
  Logger.log('Email invoices to:  ' + inboxAddress());
  Logger.log('Access key:         ' + key);
  Logger.log('Next: Deploy → New deployment → Web app (Execute as: Me, Who has access: Anyone).');
}

// Run this to invalidate the old key (e.g. if it leaked) — then paste the new one
// into Settings → Email inbox on every device.
function resetAccessKey() {
  PropertiesService.getScriptProperties().deleteProperty('ACCESS_KEY');
  setup();
}

function doGet(e) {
  const p = (e && e.parameter) || {};
  const key = PropertiesService.getScriptProperties().getProperty('ACCESS_KEY');
  if (!key || p.key !== key) return json({ error: 'unauthorized' });
  try {
    if (p.action === 'list') return json({ address: inboxAddress(), items: listItems() });
    if (p.action === 'get') return json(getItem(p.id));
    if (p.action === 'done') return json(markDone(p.id));
    return json({ error: 'unknown action' });
  } catch (err) {
    return json({ error: String(err && err.message || err) });
  }
}

function json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function inboxAddress() {
  if (SETTINGS.ADDRESS_OVERRIDE) return SETTINGS.ADDRESS_OVERRIDE.trim().toLowerCase();
  const [local, domain] = Session.getEffectiveUser().getEmail().toLowerCase().split('@');
  return local.split('+')[0] + '+' + SETTINGS.ADDRESS_TAG + '@' + domain;
}

function invoiceMimeType(att) {
  const type = (att.getContentType() || '').toLowerCase();
  const name = (att.getName() || '').toLowerCase();
  if (type === 'application/pdf' || name.endsWith('.pdf')) return 'application/pdf';
  if (type === 'image/jpeg' || type === 'image/jpg' || /\.jpe?g$/.test(name)) return 'image/jpeg';
  if (type === 'image/png' || name.endsWith('.png')) return 'image/png';
  if (type === 'image/webp' || name.endsWith('.webp')) return 'image/webp';
  return '';
}

// Every PDF/photo attachment on a message, with a stable id of
// "<message id>:<attachment index>".
function invoiceAttachments(message) {
  const out = [];
  message.getAttachments({ includeInlineImages: true, includeAttachments: true }).forEach((att, i) => {
    const mimeType = invoiceMimeType(att);
    const size = att.getSize();
    if (!mimeType || size > SETTINGS.MAX_BYTES) return;
    if (mimeType !== 'application/pdf' && size < SETTINGS.MIN_IMAGE_BYTES) return;
    out.push({ id: message.getId() + ':' + i, att, mimeType, size });
  });
  return out;
}

function doneKeys() {
  const props = PropertiesService.getScriptProperties();
  const all = props.getProperties();
  const cutoff = Date.now() - (SETTINGS.LOOKBACK_DAYS + 7) * 86400000;
  const done = {};
  Object.keys(all).forEach((k) => {
    if (k.indexOf(DONE_PREFIX) !== 0) return;
    // Emails older than the lookback window are never searched again, so their
    // "done" markers can be forgotten — keeps the property store small.
    if (Number(all[k]) < cutoff) props.deleteProperty(k);
    else done[k.slice(DONE_PREFIX.length)] = true;
  });
  return done;
}

function listItems() {
  const address = inboxAddress();
  const query = `{to:${address} deliveredto:${address} cc:${address}} has:attachment newer_than:${SETTINGS.LOOKBACK_DAYS}d`;
  const done = doneKeys();
  const items = [];
  GmailApp.search(query, 0, SETTINGS.MAX_THREADS).forEach((thread) => {
    thread.getMessages().forEach((message) => {
      invoiceAttachments(message).forEach(({ id, att, mimeType, size }) => {
        if (done[id]) return;
        items.push({
          id,
          fileName: att.getName() || 'invoice',
          mimeType,
          size,
          from: message.getFrom(),
          subject: message.getSubject(),
          receivedAt: message.getDate().toISOString(),
        });
      });
    });
  });
  items.sort((a, b) => (a.receivedAt < b.receivedAt ? 1 : -1));
  return items;
}

function findAttachment(id) {
  const [messageId] = String(id || '').split(':');
  if (!messageId) throw new Error('Missing id');
  const message = GmailApp.getMessageById(messageId);
  if (!message) throw new Error('That email no longer exists');
  const found = invoiceAttachments(message).find((a) => a.id === id);
  if (!found) throw new Error('That attachment no longer exists');
  return { message, ...found };
}

function getItem(id) {
  const { att, mimeType } = findAttachment(id);
  return { id, fileName: att.getName() || 'invoice', mimeType, data: Utilities.base64Encode(att.getBytes()) };
}

function markDone(id) {
  const { message } = findAttachment(id);
  PropertiesService.getScriptProperties().setProperty(DONE_PREFIX + id, String(Date.now()));
  const label = GmailApp.getUserLabelByName(SETTINGS.DONE_LABEL);
  if (label) message.getThread().addLabel(label);
  return { ok: true };
}
