/* Nutrition Readiness Report: make the guest's 4-page PDF and email it.

   POST /api/report   { answers: ['A'..'C' x 9], lead: { name, email, role, institution }, variant: 'k12'|'uni', consent: true }
     -> prints report/report.html for these answers with headless Chromium, emails the PDF, returns { ok }
   GET  /api/report?a=CBBBBBBBA&n=Name&r=Role&i=Institution
     -> returns the PDF itself. Off unless ALLOW_PREVIEW=1 (for testing only).

   Email is sent through Microsoft Graph when GRAPH_* is set (Compass's own mailbox), otherwise through Resend.

   Environment variables (set in Vercel, never in code):
     GRAPH_TENANT_ID, GRAPH_CLIENT_ID   Entra ID app registration with the Mail.Send application permission
     GRAPH_CLIENT_SECRET                client secret, or instead a certificate:
     GRAPH_CERT_PRIVATE_KEY, GRAPH_CERT_THUMBPRINT   certificate private key (PEM) and its thumbprint (hex, as shown in Entra ID)
     GRAPH_SENDER        the one mailbox the app may send as, e.g. CompassDiningReadinessVercel@compass-group.co.in
     RESEND_API_KEY      fallback email service key (resend.com)
     REPORT_FROM         Resend sender, e.g. "Chartwells Dining Readiness <report@reports.example.com>"
     REPORT_REPLY_TO     optional reply-to address
     SHEET_WEBHOOK_URL   optional: Google Apps Script web-app URL that adds one row per guest to the team's sheet
     SHEET_SECRET        shared key the sheet script checks (required with SHEET_WEBHOOK_URL)
     REPORT_TEAM_TO      optional: the Chartwells contact (comma separated) who gets one details email per guest,
                         with name, email, phone, role, institution, scores and the PDF (REPORT_BCC is read as a fallback)
     REPORT_BASE_URL     the site's own address, e.g. https://compass-dining-readiness.vercel.app (defaults to the Vercel production URL)
     ALLOWED_ORIGINS     extra sites allowed to send reports, comma separated (the site itself is always allowed)
     EVENT_KEY           optional shared key; when set, requests must send it in the X-Event-Key header
                         (and then the per-network limit is skipped, since venue tablets share one address)
     ALLOW_PREVIEW       1 to turn on the GET preview (testing only)
     REPORT_TO_OVERRIDE  optional: send every report to this address instead of the guest (testing only)
     CHROME_PATH         local testing only: path to a desktop Chrome */
// Vercel runs on AWS Lambda (Amazon Linux 2023); make sure @sparticuz/chromium unpacks the libraries Chromium needs
if (process.env.VERCEL && !process.env.AWS_LAMBDA_JS_RUNTIME) process.env.AWS_LAMBDA_JS_RUNTIME = 'nodejs22.x';
// both packages are loaded on first use; @sparticuz/chromium ships as an ES module
let libs = null;
async function loadLibs() {
  if (!libs) {
    const [c, p] = await Promise.all([import('@sparticuz/chromium'), import('puppeteer-core')]);
    libs = { chromium: c.default || c, puppeteer: p.default || p };
  }
  return libs;
}

const crypto = require('crypto');
const KEYS = /^[ABC]{9}$/;
const EMAIL = /^[^\s@<>"',;:\\()[\]]{1,64}@[A-Za-z0-9.-]{1,190}\.[A-Za-z]{2,24}$/;
const clip = (v, n) => String(v == null ? '' : v).replace(/[\u0000-\u001f\u007f]/g, ' ').trim().slice(0, n);
const MAX_BODY = 8 * 1024;

/* ---------- where the report page lives: fixed by configuration, never taken from the request ---------- */
const BASE = (process.env.REPORT_BASE_URL
  || (process.env.VERCEL_PROJECT_PRODUCTION_URL ? 'https://' + process.env.VERCEL_PROJECT_PRODUCTION_URL : '')
  || (process.env.VERCEL_URL ? 'https://' + process.env.VERCEL_URL : '')).replace(/\/+$/, '');
const BASE_ORIGIN = BASE ? new URL(BASE).origin : '';
const ALLOWED = new Set([BASE_ORIGIN, ...(process.env.ALLOWED_ORIGINS || '').split(',').map(s => s.trim().replace(/\/+$/, ''))].filter(Boolean));
// the headless browser may only load the site itself and the two public libraries the report uses
const BROWSER_ALLOW = [BASE_ORIGIN, 'https://cdnjs.cloudflare.com', 'https://fonts.googleapis.com', 'https://fonts.gstatic.com'];

/* ---------- simple abuse limits (per server instance; add a Vercel Firewall rate-limit rule for a global limit) ---------- */
const hits = new Map();
function limited(key, max, ms) {
  const now = Date.now(), list = (hits.get(key) || []).filter(t => now - t < ms);
  list.push(now); hits.set(key, list);
  if (hits.size > 5000) for (const [k, v] of hits) if (!v.some(t => now - t < ms)) hits.delete(k);
  return list.length > max;
}

async function makePdf(data) {
  if (!BASE) throw new Error('REPORT_BASE_URL is not set');
  const { chromium, puppeteer } = await loadLibs();
  const local = process.env.CHROME_PATH;
  if (!local) chromium.setGraphicsMode = true;           // software WebGL (SwiftShader) for the 3D cafeteria
  const browser = await puppeteer.launch(local
    ? { executablePath: local, headless: true, args: ['--use-angle=swiftshader', '--enable-unsafe-swiftshader'] }
    : { executablePath: await chromium.executablePath(), headless: 'shell', args: chromium.args });
  try {
    const page = await browser.newPage();
    await page.setRequestInterception(true);
    page.on('request', r => {
      const u = r.url();
      if (u.startsWith('data:') || u.startsWith('blob:') || BROWSER_ALLOW.some(o => u === o || u.startsWith(o + '/'))) r.continue();
      else r.abort();
    });
    await page.setViewport({ width: 1240, height: 1754, deviceScaleFactor: 1 });
    await page.evaluateOnNewDocument(d => { window.REPORT_DATA = d; }, data);
    await page.goto(BASE + '/report/report.html', { waitUntil: 'load', timeout: 30000 });
    await page.waitForFunction('window.REPORT_READY === true', { timeout: 25000 });
    const pdf = await page.pdf({ printBackground: true, preferCSSPageSize: true });
    return Buffer.from(pdf);
  } finally { await browser.close(); }
}

function readData(src) {
  const answers = String(src.answers || '').toUpperCase().replace(/[^ABC]/g, '');
  if (!KEYS.test(answers)) throw Object.assign(new Error('Need 9 answers, each A, B or C'), { status: 400 });
  const lead = src.lead || {};
  return {
    answers: answers.split(''),
    lead: { name: clip(lead.name, 80), role: clip(lead.role, 80), institution: clip(lead.institution, 90) },
    variant: src.variant === 'uni' ? 'uni' : src.variant === 'k12' ? 'k12' : undefined,
    date: new Date().toISOString().slice(0, 10)
  };
}

/* ---------- lead log: one row per guest in the Chartwells team's Google Sheet (optional) ----------
   SHEET_WEBHOOK_URL is the sheet's Apps Script web-app address, SHEET_SECRET the shared key it checks.
   Logging never blocks or fails the guest's email. */
const POINTS = { A: 1, B: 3, C: 5 };
const STAGES = [[15, 'Foundation'], [22, 'Emerging'], [30, 'Developing'], [38, 'Advanced'], [45, 'Future-Ready']];
function scoreOf(data) {
  const pts = data.answers.map(k => POINTS[k]);
  const sum = a => a.reduce((x, y) => x + y, 0), total = sum(pts);
  return { total, stage: STAGES.find(s => total <= s[0])[1], nutrition: sum(pts.slice(0, 3)), safety: sum(pts.slice(3, 6)), awareness: sum(pts.slice(6, 9)) };
}
async function logLead(data, email, phone, status) {
  const url = process.env.SHEET_WEBHOOK_URL, secret = process.env.SHEET_SECRET;
  if (!url || !secret) return;
  const sc = scoreOf(data);
  // a leading = + - @ would make the sheet run the cell as a formula; prefix such cells with an apostrophe
  const cell = v => /^[=+\-@]/.test(String(v)) ? "'" + v : v;
  const row = {
    secret, time: new Date().toISOString(), name: cell(data.lead.name), email: cell(email), phone: cell(phone),
    role: cell(data.lead.role), institution: cell(data.lead.institution), variant: data.variant || '',
    total: sc.total, stage: sc.stage, nutrition: sc.nutrition, safety: sc.safety, awareness: sc.awareness,
    answers: data.answers.join(''), status
  };
  try { await fetch(url, { method: 'POST', headers: { 'Content-Type': 'text/plain' }, body: JSON.stringify(row), signal: AbortSignal.timeout(8000) }); }
  catch (err) { console.error('sheet log failed:', err.message); }
}

const TEAM = (process.env.REPORT_TEAM_TO || process.env.REPORT_BCC || '').split(',').map(a => a.trim()).filter(a => /^[^\s@<>]+@[^\s@<>]+\.[^\s@<>]+$/.test(a)).slice(0, 3);
const esc = s => String(s).replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' })[c]);
function emailHtml(d) {
  const hi = d.lead.name ? 'Dear ' + esc(d.lead.name.split(' ')[0]) + ',' : 'Hello,';
  const who = d.lead.institution ? esc(d.lead.institution) + '’s' : 'your institution’s';
  const logo = `<img src="${BASE}/img/email-logo-k12.png" width="240" height="47" alt="Chartwells" style="display:block;border:0">`;
  return `<div style="font-family:Arial,Helvetica,sans-serif;font-size:15px;line-height:1.55;color:#14181F;max-width:560px">
<p style="margin:0 0 20px">${logo}</p>
<p>${hi}</p>
<p>Thank you for taking the Nutrition Readiness check with Chartwells. Your personal report is attached as a PDF.</p>
<p>It shows ${who} score across nutrition, food safety, and awareness and sustainability, what is already working, and a game plan for the next 30 days, this term and this year.</p>
<p>Your Chartwells contact will be in touch to go through it with you. You can also write to us at sarika.chauhan@compass-group.co.in or call +91 96671 80483.</p>
<p>Warm regards,<br>The Chartwells team<br><span style="color:#5A6170">Chartwells by Compass India · Serving up happy &amp; healthy</span></p>
</div>`;
}

// the details email the Chartwells contact gets for each guest (internal; reply goes straight to the guest)
function teamHtml(d, email, phone, status) {
  const sc = scoreOf(d);
  const row = (k, v) => `<tr><td style="padding:6px 12px 6px 0;color:#5A6170;vertical-align:top;width:38%">${k}</td><td style="padding:6px 0;font-weight:bold;vertical-align:top;word-break:break-word">${v || '–'}</td></tr>`;
  const when = new Date().toLocaleString('en-IN', { timeZone: 'Asia/Kolkata', day: '2-digit', month: 'short', year: 'numeric', hour: '2-digit', minute: '2-digit' });
  return `<div style="font-family:Arial,Helvetica,sans-serif;font-size:14px;line-height:1.5;color:#14181F;max-width:560px">
<p style="margin:0 0 4px;font-size:16px;font-weight:bold">New guest at the Nutrition Readiness check</p>
<p style="margin:0 0 14px;color:#5A6170">${when} IST · their report is attached · reply to this email to write to the guest</p>
<table style="border-collapse:collapse;font-size:14px;width:100%;table-layout:fixed">
${row('Name', esc(d.lead.name))}${row('Email', esc(email))}${row('Phone', esc(phone))}${row('Role', esc(d.lead.role))}${row('Institution', esc(d.lead.institution))}${row('Type', d.variant === 'uni' ? 'Higher Ed' : 'K-12')}
${row('Score', sc.total + ' / 45 · ' + sc.stage + ' stage')}${row('Nutrition', sc.nutrition + ' / 15')}${row('Food Safety', sc.safety + ' / 15')}${row('Awareness &amp; Sustain&shy;ability', sc.awareness + ' / 15')}${row('Report email', status === 'sent' ? 'Sent to the guest' : 'Could not be sent to the guest: please send them this PDF')}
</table>
</div>`;
}

/* ---------- sending ---------- */
let graphToken = null;
const b64url = b => Buffer.from(b).toString('base64').replace(/=+$/, '').replace(/\+/g, '-').replace(/\//g, '_');
// a client secret, or a signed client assertion (certificate credential, RFC 7523)
function graphCredential() {
  if (process.env.GRAPH_CERT_PRIVATE_KEY && process.env.GRAPH_CERT_THUMBPRINT) return { client_assertion_type: 'urn:ietf:params:oauth:client-assertion-type:jwt-bearer', client_assertion: graphAssertion() };
  return { client_secret: process.env.GRAPH_CLIENT_SECRET };
}
function graphAssertion() {
  const now = Math.floor(Date.now() / 1000), id = process.env.GRAPH_CLIENT_ID;
  const head = { alg: 'RS256', typ: 'JWT', x5t: b64url(Buffer.from(process.env.GRAPH_CERT_THUMBPRINT.replace(/[^0-9a-f]/gi, ''), 'hex')) };
  const body = { aud: `https://login.microsoftonline.com/${process.env.GRAPH_TENANT_ID}/oauth2/v2.0/token`, iss: id, sub: id, jti: crypto.randomUUID(), nbf: now, exp: now + 600 };
  const data = b64url(JSON.stringify(head)) + '.' + b64url(JSON.stringify(body));
  const key = process.env.GRAPH_CERT_PRIVATE_KEY.replace(/\\n/g, '\n');   // PEM pasted on one line with \n
  return data + '.' + b64url(crypto.createSign('RSA-SHA256').update(data).sign(key));
}
async function graphAccessToken() {
  if (graphToken && graphToken.exp > Date.now() + 60000) return graphToken.value;
  const r = await fetch(`https://login.microsoftonline.com/${encodeURIComponent(process.env.GRAPH_TENANT_ID)}/oauth2/v2.0/token`, {
    method: 'POST', headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
    body: new URLSearchParams(Object.assign({ client_id: process.env.GRAPH_CLIENT_ID, scope: 'https://graph.microsoft.com/.default', grant_type: 'client_credentials' }, graphCredential()))
  });
  const j = await r.json().catch(() => ({}));
  if (!r.ok || !j.access_token) throw new Error('Graph token: ' + (j.error_description || j.error || r.status));
  graphToken = { value: j.access_token, exp: Date.now() + (j.expires_in || 3600) * 1000 };
  return graphToken.value;
}
async function sendGraph(to, subject, html, pdf, replyTo = process.env.REPORT_REPLY_TO) {
  const token = await graphAccessToken();
  const r = await fetch(`https://graph.microsoft.com/v1.0/users/${encodeURIComponent(process.env.GRAPH_SENDER)}/sendMail`, {
    method: 'POST', headers: { Authorization: 'Bearer ' + token, 'Content-Type': 'application/json' },
    body: JSON.stringify({
      message: {
        subject, body: { contentType: 'HTML', content: html },
        toRecipients: [].concat(to).map(a => ({ emailAddress: { address: a } })),
        replyTo: replyTo ? [{ emailAddress: { address: replyTo } }] : [],
        attachments: [{ '@odata.type': '#microsoft.graph.fileAttachment', name: 'Nutrition-Readiness-Report.pdf', contentType: 'application/pdf', contentBytes: pdf.toString('base64') }]
      },
      saveToSentItems: false
    })
  });
  if (r.status !== 202) { const t = await r.text().catch(() => ''); throw new Error('Graph sendMail ' + r.status + ' ' + t.slice(0, 300)); }
}
async function sendResend(to, subject, html, pdf, replyTo = process.env.REPORT_REPLY_TO) {
  const r = await fetch('https://api.resend.com/emails', {
    method: 'POST', headers: { Authorization: 'Bearer ' + process.env.RESEND_API_KEY, 'Content-Type': 'application/json' },
    body: JSON.stringify({ from: process.env.REPORT_FROM || 'Chartwells Dining Readiness <onboarding@resend.dev>', to: [].concat(to), reply_to: replyTo || undefined, subject, html,
      attachments: [{ filename: 'Nutrition-Readiness-Report.pdf', content: pdf.toString('base64') }] })
  });
  if (!r.ok) { const t = await r.text().catch(() => ''); throw new Error('Resend ' + r.status + ' ' + t.slice(0, 300)); }
}
const provider = () => (process.env.GRAPH_TENANT_ID && process.env.GRAPH_CLIENT_ID && process.env.GRAPH_SENDER && (process.env.GRAPH_CLIENT_SECRET || (process.env.GRAPH_CERT_PRIVATE_KEY && process.env.GRAPH_CERT_THUMBPRINT))) ? 'graph' : process.env.RESEND_API_KEY ? 'resend' : '';

/* ---------- handler ---------- */
module.exports = async (req, res) => {
  res.setHeader('Cache-Control', 'no-store');
  res.setHeader('X-Content-Type-Options', 'nosniff');
  const origin = (req.headers.origin || '').replace(/\/+$/, '');
  const fromAllowed = origin && ALLOWED.has(origin);
  if (fromAllowed && origin !== BASE_ORIGIN) {          // other approved sites (for example a test copy) get CORS
    res.setHeader('Access-Control-Allow-Origin', origin);
    res.setHeader('Access-Control-Allow-Methods', 'POST, OPTIONS');
    res.setHeader('Access-Control-Allow-Headers', 'Content-Type, X-Event-Key');
    res.setHeader('Vary', 'Origin');
  }
  if (req.method === 'OPTIONS') return res.status(fromAllowed ? 204 : 403).end();
  const ip = String(req.headers['x-forwarded-for'] || req.socket?.remoteAddress || '').split(',')[0].trim();
  try {
    if (req.method === 'GET') {
      if (process.env.ALLOW_PREVIEW !== '1') return res.status(404).json({ ok: false, error: 'Not found' });
      if (limited('get:' + ip, 10, 10 * 60e3)) return res.status(429).json({ ok: false, error: 'Too many requests' });
      const q = req.query || {};
      const pdf = await makePdf(readData({ answers: q.a, lead: { name: q.n, role: q.r, institution: q.i }, variant: q.v }));
      res.setHeader('Content-Type', 'application/pdf');
      res.setHeader('Content-Disposition', 'inline; filename="Nutrition-Readiness-Report.pdf"');
      return res.status(200).send(pdf);
    }
    if (req.method !== 'POST') return res.status(405).json({ ok: false, error: 'Use POST' });

    // only the quiz's own site (and approved test sites) may ask for a report
    if (!fromAllowed) return res.status(403).json({ ok: false, error: 'Not allowed' });
    if (process.env.EVENT_KEY && req.headers['x-event-key'] !== process.env.EVENT_KEY) return res.status(403).json({ ok: false, error: 'Not allowed' });
    if (!/^application\/json\b/i.test(req.headers['content-type'] || '')) return res.status(415).json({ ok: false, error: 'Send JSON' });
    if (+(req.headers['content-length'] || 0) > MAX_BODY) return res.status(413).json({ ok: false, error: 'Request too large' });
    let body;   // Vercel parses JSON lazily and throws on bad input, so read it inside a try
    try { body = req.body; if (typeof body === 'string') body = JSON.parse(body || '{}'); } catch (e) { return res.status(400).json({ ok: false, error: 'Invalid JSON' }); }
    body = body || {};
    if (Buffer.byteLength(JSON.stringify(body)) > MAX_BODY) return res.status(413).json({ ok: false, error: 'Request too large' });

    if (body.consent !== true) return res.status(400).json({ ok: false, error: 'Consent is required' });
    const email = clip(body.lead && body.lead.email, 120);
    if (!EMAIL.test(email)) return res.status(400).json({ ok: false, error: 'A valid email is required' });
    const data = readData({ answers: Array.isArray(body.answers) ? body.answers.join('') : '', lead: body.lead, variant: body.variant });
    // venue tablets share one internet address: allow a busy stand (100 per 10 minutes), or skip this limit with EVENT_KEY
    if ((!process.env.EVENT_KEY && limited('ip:' + ip, 100, 10 * 60e3)) || limited('to:' + email.toLowerCase(), 3, 60 * 60e3)) return res.status(429).json({ ok: false, error: 'Too many requests' });
    const via = provider();
    if (!via) return res.status(503).json({ ok: false, error: 'Email is not set up yet' });

    const pdf = await makePdf(data);
    const to = process.env.REPORT_TO_OVERRIDE || email;
    const subject = `${data.lead.name ? data.lead.name.split(' ')[0] + ', your' : 'Your'} Nutrition Readiness Report for ${data.lead.institution || 'your institution'}`;
    const phone = clip(body.lead && body.lead.phone, 30);
    const send = via === 'graph' ? sendGraph : sendResend;
    let status = 'sent';
    try { await send(to, subject, emailHtml(data), pdf); }
    catch (err) { console.error('send failed:', err.message); status = 'email failed'; }
    // the Chartwells contact gets the guest's details and report either way, so no guest is lost
    if (TEAM.length) {
      const sc = scoreOf(data);
      const teamSubject = `New guest: ${data.lead.name || email}${data.lead.institution ? ', ' + data.lead.institution : ''} (${sc.total}/45)`;
      try { await send(TEAM, teamSubject, teamHtml(data, email, phone, status), pdf, email); }
      catch (err) { console.error('team email failed:', err.message); }
    }
    await logLead(data, email, phone, status);
    if (status !== 'sent') return res.status(502).json({ ok: false, error: 'The email could not be sent' });
    return res.status(200).json({ ok: true });
  } catch (err) {
    console.error(err);
    return res.status(err.status || 500).json({ ok: false, error: err.status ? err.message : 'Report could not be made' });
  }
};
