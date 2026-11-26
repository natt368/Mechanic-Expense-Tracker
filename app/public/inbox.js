// Email inbox — talks to the Google Apps Script web app in app/email-inbox/Code.gs,
// which serves the PDF/photo attachments of invoices emailed to the account's
// +invoices address. Plain GET requests (no custom headers), so the browser can
// call the script cross-origin without a preflight.
const EmailInbox = (() => {
  async function call(url, key, params) {
    const qs = new URLSearchParams({ ...params, key });
    let res;
    try {
      res = await fetch(url.trim() + (url.includes('?') ? '&' : '?') + qs.toString());
    } catch (e) {
      // A fetch that never gets a readable response is almost always the script
      // deployment's access not being "Anyone" (Google answers with a sign-in
      // page instead of the script's JSON) or a mistyped URL.
      throw new Error("Couldn't reach the email inbox script. Check the web app URL, and that the deployment's \"Who has access\" is set to Anyone.");
    }
    let data;
    try { data = await res.json(); } catch (e) { throw new Error("The email inbox script didn't send a valid response. Check the web app URL."); }
    if (data.error === 'unauthorized') throw new Error('The email inbox access key is wrong. Copy it again from the script\'s execution log.');
    if (data.error) throw new Error('Email inbox: ' + data.error);
    return data;
  }

  return {
    list(url, key) { return call(url, key, { action: 'list' }); },
    async get(url, key, id) {
      const data = await call(url, key, { action: 'get', id });
      return { fileName: data.fileName, mimeType: data.mimeType, dataUrl: 'data:' + data.mimeType + ';base64,' + data.data };
    },
    done(url, key, id) { return call(url, key, { action: 'done', id }); },
    // AI invoice reading. A POST with a text/plain body is a "simple" request,
    // so it also goes straight through without a CORS preflight.
    async extract(url, key, { dataUrl, mimeType, inboxId, context }) {
      let res;
      try {
        res = await fetch(url.trim(), {
          method: 'POST',
          headers: { 'Content-Type': 'text/plain;charset=utf-8' },
          body: JSON.stringify({ key, action: 'extract', inboxId: inboxId || '', context, file: { mimeType, data: dataUrl.split(',')[1] || '' } }),
        });
      } catch (e) {
        throw new Error("Couldn't reach the Google script for AI reading.");
      }
      let data;
      try { data = await res.json(); } catch (e) { throw new Error("The Google script didn't send a valid response — it may need redeploying with the latest Code.gs."); }
      if (data.error === 'unauthorized') throw new Error('The email inbox access key is wrong.');
      if (data.error) throw new Error(data.error);
      return data;
    },
  };
})();
