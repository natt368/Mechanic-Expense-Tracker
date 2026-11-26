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
 *   3. Paste the web app URL and the access key into the app: Settings → Email & AI.
 *
 * Optional AI invoice reading: add a script property named ANTHROPIC_API_KEY
 * (Project Settings → Script properties) holding a Claude API key. The app then sends
 * each invoice here, this script asks Claude to read it line by line and suggest the
 * equipment and expense type for each line from your past invoices, and the app shows
 * the suggestions for you to check before saving. The key never leaves this script.
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

  // AI invoice reading (only used once ANTHROPIC_API_KEY is set — see the top of
  // this file). Effort trades speed for care: 'low' is quickest; 'medium' or 'high'
  // think longer about tricky invoices but take longer to answer.
  AI_MODEL: 'claude-opus-5-5',
  AI_EFFORT: 'low',
  // Lines of past invoices sent as examples, most recent first.
  AI_HISTORY_LINES: 300,
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
  Logger.log('AI invoice reading:  ' + (aiApiKey() ? 'on (ANTHROPIC_API_KEY is set)' : 'off — add an ANTHROPIC_API_KEY script property to turn it on'));
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
    if (p.action === 'list') return json({ address: inboxAddress(), items: listItems(), ai: !!aiApiKey() });
    if (p.action === 'get') return json(getItem(p.id));
    if (p.action === 'done') return json(markDone(p.id));
    return json({ error: 'unknown action' });
  } catch (err) {
    return json({ error: String(err && err.message || err) });
  }
}

// The app sends invoices for AI reading as a POST (text/plain, so the browser
// skips the CORS preflight Apps Script can't answer) with the key in the body.
function doPost(e) {
  let body;
  try { body = JSON.parse((e && e.postData && e.postData.contents) || '{}'); } catch (err) { return json({ error: 'bad request' }); }
  const key = PropertiesService.getScriptProperties().getProperty('ACCESS_KEY');
  if (!key || body.key !== key) return json({ error: 'unauthorized' });
  try {
    if (body.action === 'extract') return json(extractInvoice(body));
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

// ------------------------------------------------------------------ AI invoice reading

function aiApiKey() {
  return PropertiesService.getScriptProperties().getProperty('ANTHROPIC_API_KEY') || '';
}

const AI_INSTRUCTIONS = [
  'You read invoices for a farm equipment repair shop and code every line so it can be booked against the right machine.',
  '',
  'For each purchasable line on the invoice (parts, labour, fuel, fluids, tires, freight, shop supplies, environmental or core charges), return:',
  '- part: the description as printed, cleaned up (no leading line numbers).',
  '- partNumber: the part / SKU number if there is one, else "".',
  '- qty, unitCost, totalCost: numbers. totalCost is the line amount after any line discount, before sales tax. If only a total is shown, set unitCost = totalCost / qty.',
  '- liters: litres of fuel for fuel lines (convert gallons to litres), else 0.',
  '- equipment: the machine this line was bought for. It MUST be exactly one of the names in the equipment list, or "" when there is no reasonable basis to choose.',
  '- expenseType: what kind of expense it is. Prefer one of the expense types listed; only invent a new short type when none fits.',
  '- confidence: "high" when the invoice or history makes the equipment clear, "medium" for a sensible inference, "low" for a guess.',
  '- reason: a few words on why you picked that equipment (e.g. "same part # as before", "invoice says unit 4", "only combine in list").',
  '',
  'Do not return subtotal, total, tax (GST/PST/HST/VAT), payment, balance or deposit lines.',
  '',
  'Choosing equipment, strongest evidence first:',
  '1. Anything on the invoice or in the email naming the machine: unit number, serial number, make/model, "for the combine", etc.',
  '2. History: the same part number, or the same kind of part from the same vendor, almost always goes to the same machine as before.',
  '3. Fit: what the part is for and which listed machines it plausibly fits.',
  'If one invoice is clearly all for one machine, use it for every line unless a line obviously belongs elsewhere.',
  'Also return the invoice-level vendor (the business that issued the invoice) and invoiceDate (YYYY-MM-DD, "" if not shown).',
].join('\n');

const AI_SCHEMA = {
  type: 'object',
  properties: {
    vendor: { type: 'string' },
    invoiceDate: { type: 'string' },
    lineItems: {
      type: 'array',
      items: {
        type: 'object',
        properties: {
          part: { type: 'string' },
          partNumber: { type: 'string' },
          qty: { type: 'number' },
          unitCost: { type: 'number' },
          totalCost: { type: 'number' },
          liters: { type: 'number' },
          equipment: { type: 'string' },
          expenseType: { type: 'string' },
          confidence: { type: 'string', enum: ['high', 'medium', 'low'] },
          reason: { type: 'string' },
        },
        required: ['part', 'partNumber', 'qty', 'unitCost', 'totalCost', 'liters', 'equipment', 'expenseType', 'confidence', 'reason'],
        additionalProperties: false,
      },
    },
  },
  required: ['vendor', 'invoiceDate', 'lineItems'],
  additionalProperties: false,
};

function cell(v) { return String(v == null ? '' : v).replace(/[\t\r\n]+/g, ' ').trim(); }

function buildContext(body) {
  const ctx = body.context || {};
  const equipment = (ctx.equipment || []).map((e) => '- ' + cell(e.name) + (e.category ? ' (' + cell(e.category) + ')' : '') + (e.details ? ' — ' + cell(e.details) : ''));
  const types = (ctx.expenseTypes || []).map((t) => '- ' + cell(t));
  const history = (ctx.history || []).slice(0, SETTINGS.AI_HISTORY_LINES)
    .map((h) => [h.part, h.partNumber, h.vendor, h.equipment, h.expenseType].map(cell).join('\t'));
  let text = 'Equipment list:\n' + (equipment.join('\n') || '(none yet)') +
    '\n\nExpense types used so far:\n' + (types.join('\n') || '(none yet)') +
    '\n\nHow past invoice lines were coded (most recent first; part, part #, vendor, equipment, expense type — tab separated):\n' +
    (history.join('\n') || '(no history yet)');
  if (body.inboxId) {
    try {
      const { message } = findAttachment(body.inboxId);
      text += '\n\nThis invoice was emailed in. From: ' + cell(message.getFrom()) + '\nSubject: ' + cell(message.getSubject()) +
        '\nEmail text:\n' + message.getPlainBody().slice(0, 3000);
    } catch (err) { /* email gone — read the file alone */ }
  }
  return text;
}

function extractInvoice(body) {
  const apiKey = aiApiKey();
  if (!apiKey) throw new Error('AI reading is not set up: add an ANTHROPIC_API_KEY script property.');
  const file = body.file || {};
  if (!file.data || !file.mimeType) throw new Error('No invoice file was sent');
  const fileBlock = file.mimeType === 'application/pdf'
    ? { type: 'document', source: { type: 'base64', media_type: 'application/pdf', data: file.data } }
    : { type: 'image', source: { type: 'base64', media_type: file.mimeType, data: file.data } };
  const request = {
    model: SETTINGS.AI_MODEL,
    max_tokens: 16000,
    // If a safety classifier declines the request, Anthropic re-runs it on its
    // recommended fallback model instead of returning nothing.
    fallbacks: 'default',
    output_config: { effort: SETTINGS.AI_EFFORT, format: { type: 'json_schema', schema: AI_SCHEMA } },
    system: AI_INSTRUCTIONS,
    messages: [{ role: 'user', content: [fileBlock, { type: 'text', text: buildContext(body) + '\n\nRead the invoice above and code every line.' }] }],
  };
  const res = UrlFetchApp.fetch('https://api.anthropic.com/v1/messages', {
    method: 'post',
    contentType: 'application/json',
    headers: {
      'x-api-key': apiKey,
      'anthropic-version': '2026-10-21',
      'anthropic-beta': 'server-side-fallback-2026-07-01',
    },
    payload: JSON.stringify(request),
    muteHttpExceptions: true,
  });
  let data;
  try { data = JSON.parse(res.getContentText()); } catch (err) { throw new Error('Claude API returned HTTP ' + res.getResponseCode()); }
  if (res.getResponseCode() !== 200) throw new Error('Claude API: ' + ((data.error && data.error.message) || ('HTTP ' + res.getResponseCode())));
  if (data.stop_reason === 'refusal') throw new Error('Claude declined to read this invoice.');
  if (data.stop_reason === 'max_tokens') throw new Error('The invoice was too long to read in one go.');
  const text = (data.content || []).filter((b) => b.type === 'text').map((b) => b.text).join('');
  const result = JSON.parse(text);
  result.model = data.model;
  result.usage = data.usage;
  return result;
}
