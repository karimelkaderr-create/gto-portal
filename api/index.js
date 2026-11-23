// GTO Operations Portal – server API (Vercel serverless function)
// The Excel files are the database. Only the opsportal account can open them;
// staff never get file access. Every request is checked against the Access list.
import crypto from 'node:crypto';
import { buildIndex, answer as localAnswer } from './assist.js';
import { readSheet } from '../lib/xlsx-lite.js';
import { sectionsOf } from '../lib/doc-lite.js';
const blob = () => globalThis.__blobMock || import('@vercel/blob');

const CLIENT_ID = process.env.CLIENT_ID || 'b57c476e-4a3e-498c-a114-2aa2638997d0';
const TENANT = process.env.TENANT_ID || '082517b5-3937-4f49-87e1-9a3dec0551bd';
const SECRET = process.env.CLIENT_SECRET || '';
const OPS = (process.env.OPS_ACCOUNT || 'opsportal@gto.org.uk').toLowerCase();
const URLS = {
  portal: process.env.PORTAL_DATA_URL || 'https://inhouseworlduk-my.sharepoint.com/personal/opsportal_gto_org_uk/Documents/GTO Portal Data.xlsx',
  tracker: process.env.TRACKER_URL || 'https://inhouseworlduk-my.sharepoint.com/:x:/g/personal/karim_abdelkader_gto_org_uk/IQAGYUtqR1vLSKqP0FlZ6RATAdBiU8BgwfIHE4l4MyEuylk',
  rota: process.env.ROTA_URL || 'https://inhouseworlduk-my.sharepoint.com/:x:/g/personal/hui-fen_inhouseworld_co_uk/IQCsgxuG2GMfQrdxchuLp-g-AUXcRe-tUdm0up6CarphlSA',
  credits: process.env.CREDITS_URL || 'https://inhouseworlduk-my.sharepoint.com/:x:/g/personal/hui-fen_inhouseworld_co_uk/IQCAIRDbCANzQYGu3vQA99zVAXuw_Brrgq6CWaW3YbmRRCU',
  kb: process.env.KB_URL || 'https://inhouseworlduk-my.sharepoint.com/:x:/g/personal/elsaid_zaki_gto_org_uk/IQDRm5MVVCi3QZPYHMDlHTYTAYohlhdIhvHRb7r5Tgs25gI',
};
const AI_KEY = process.env.ANTHROPIC_API_KEY || '';
const AI_MODEL = process.env.AI_MODEL || 'claude-haiku-5-5';
const TRACKER_TABLE = process.env.TRACKER_TABLE || 'Table1';
const BALANCE_SHEET = process.env.BALANCE_SHEET || 'Balance';
const T = { access: 'tblAccess', att: 'tblAttendance', req: 'tblRequests', set: 'tblSettings', swap: 'tblSwaps' };
const ROLES = ['Agent', 'Team Leader', 'Duty Manager', 'General Manager', 'HR', 'Quality'];
// people who report straight to the General Manager: their leave skips the manager step
const REPORTS_TO_GM = ['Team Leader', 'Duty Manager', 'HR'];
const MANAGER_ROLES = ['Team Leader', 'Duty Manager', 'General Manager', 'HR'];
// '.default' = every Graph permission the admin has approved for the app (Files.ReadWrite.All, User.Read)
const SCOPE = 'offline_access openid profile https://graph.microsoft.com/.default';

/* ---------------- helpers ---------------- */
const pad = n => String(n).padStart(2, '0');
const dkey = d => `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}`;
const fromKey = k => { const [y, m, d] = k.split('-').map(Number); return new Date(y, m - 1, d); };
const serialToDate = n => { const u = new Date(Math.round((n - 25569) * 864e5)); return new Date(u.getUTCFullYear(), u.getUTCMonth(), u.getUTCDate()); };
const dateToSerial = d => (Date.UTC(d.getFullYear(), d.getMonth(), d.getDate()) - Date.UTC(1899, 11, 30)) / 864e5;
const norm = s => String(s ?? '').replace(/\u00a0|\u202f/g, ' ').toLowerCase().replace(/\(.*?\)/g, m => ' ' + m.slice(1, -1) + ' ').replace(/[^a-z\u0600-\u06ff ]+/g, ' ').replace(/\s+/g, ' ').trim();
const normEmail = s => String(s ?? '').trim().toLowerCase();
const yes = v => /^(y|yes|true|1|x|✓|✔|admin)$/i.test(String(v??'').trim()) || v===true || v===1;
const isNo = v => /^no$/i.test(String(v ?? '').trim());
const uid = p => p + Date.now().toString(36).toUpperCase() + crypto.randomBytes(2).toString('hex').toUpperCase();
const isKey = k => /^\d{4}-\d{2}-\d{2}$/.test(String(k || ''));
function parseDateCell(v) {
  if (v == null || v === '') return null;
  if (typeof v === 'number') return v > 20000 && v < 80000 ? serialToDate(v) : null;
  const s = String(v).trim();
  let m = s.match(/^(\d{4})-(\d{2})-(\d{2})/); if (m) return new Date(+m[1], +m[2] - 1, +m[3]);
  m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{2,4})$/); if (m) return new Date(m[3].length === 2 ? 2000 + +m[3] : +m[3], +m[2] - 1, +m[1]);
  if (/^\d+(\.\d+)?$/.test(s)) return parseDateCell(Number(s));
  const t = Date.parse(s); return isNaN(t) ? null : new Date(new Date(t).setHours(0, 0, 0, 0));
}
const toKey = v => { const d = parseDateCell(v); return d ? dkey(d) : String(v || ''); };
function tsString(v) { if (typeof v === 'number') return new Date(Math.round((v - 25569) * 864e5)).toISOString(); return String(v || ''); }
function rowsToObjects(values) { const h = values[0].map(x => String(x ?? '').trim()); return { headers: h, rows: values.slice(1).map((r, i) => { const o = { _i: i }; h.forEach((k, j) => o[k] = r[j]); return o; }) }; }
class HttpError extends Error { constructor(status, msg) { super(msg); this.status = status; } }
const deny = msg => { throw new HttpError(403, msg || "You don't have permission to do that."); };
const bad = msg => { throw new HttpError(400, msg); };

/* ---------------- rota parsing (same as the page) ---------------- */
const DAYRX = /^(mon|tue|wed|thu|fri|sat|sun)[a-z]*\.?\s+(\d{1,2})\s*\/\s*(\d{1,2})/i;
function dateFromDM(d, m, near) { const base = near || new Date(); let best = null, bd = 1e15; for (const y of [base.getFullYear() - 1, base.getFullYear(), base.getFullYear() + 1]) { const x = new Date(y, m - 1, d); const diff = Math.abs(x - base); if (diff < bd) { bd = diff; best = x; } } return best; }
// Tab names carry no year ("WK 02.11"). Every week tab starts on a Monday, so take the year in which that date is a Monday
// (nearest to today); only if none is, fall back to the nearest date.
function sheetWeekStart(name, today = new Date()) {
  const m = String(name).match(/(\d{1,2})\s*\.\s*(\d{1,2})/); if (!m) return null; const d = +m[1], mo = +m[2];
  const cands = [-1, 0, 1].map(k => new Date(today.getFullYear() + k, mo - 1, d)).filter(x => x.getMonth() === mo - 1);
  const mondays = cands.filter(x => x.getDay() === 1).sort((a, b) => Math.abs(a - today) - Math.abs(b - today));
  return mondays[0] || dateFromDM(d, mo, today);
}
function parseRotaSheet(grid, weekStart) {
  const people = []; const seen = new Set();
  for (let r = 0; r < grid.length; r++) {
    const row = grid[r] || [];
    for (let c = 1; c < row.length; c++) {
      if (!DAYRX.test(String(row[c] || '').trim()) || DAYRX.test(String(row[c - 1] || '').trim())) continue;
      let n = 0; while (n < 7 && DAYRX.test(String(row[c + n] || '').trim())) n++;
      if (n < 5) continue;
      const rawLabel = String(row[c - 1] || '').trim();
      if (/^shift$/i.test(rawLabel) || /counter/i.test(rawLabel)) continue;
      const label = /^agent name$/i.test(rawLabel) ? 'Agents' : /^tl\s*\/\s*dm$/i.test(rawLabel) ? 'TL / DM' : /^name$/i.test(rawLabel) ? 'Team' : rawLabel.replace(/\s+member$/i, '');
      const dates = []; for (let k = 0; k < n; k++) { const m = String(row[c + k]).trim().match(DAYRX); dates.push(dateFromDM(+m[2], +m[3], weekStart)); }
      for (let rr = r + 1; rr < grid.length; rr++) {
        const rw = grid[rr] || []; const name = String(rw[c - 1] || '').replace(/\u00a0/g, ' ').trim();
        if (!name || DAYRX.test(String(rw[c] || '').trim()) || /counter|^total$|^shift$/i.test(name)) break;
        const days = {}; dates.forEach((d, k) => { days[dkey(d)] = String(rw[c + k] ?? '').replace(/\u00a0/g, ' ').trim(); });
        const key = norm(name) + '|' + label; if (seen.has(key)) continue; seen.add(key);
        people.push({ name, team: label || 'Team', days, at: Object.fromEntries(dates.map((d, k) => [dkey(d), [rr, c + k]])), row: rr, c });
      }
    }
  }
  return people;
}

/* ---------------- opsportal token (stored encrypted in Vercel Blob) ---------------- */
const KEY = () => crypto.createHash('sha256').update('gto-portal-token:' + SECRET).digest();
function enc(obj) { const iv = crypto.randomBytes(12); const c = crypto.createCipheriv('aes-256-gcm', KEY(), iv); const data = Buffer.concat([c.update(JSON.stringify(obj), 'utf8'), c.final()]); return Buffer.concat([iv, c.getAuthTag(), data]).toString('base64'); }
function dec(s) { const b = Buffer.from(s, 'base64'); const d = crypto.createDecipheriv('aes-256-gcm', KEY(), b.subarray(0, 12)); d.setAuthTag(b.subarray(12, 28)); return JSON.parse(Buffer.concat([d.update(b.subarray(28)), d.final()]).toString('utf8')); }
const TOKEN_PATH = 'auth/ops-token.txt';
async function saveRT(rt) { const { put } = await blob(); await put(TOKEN_PATH, enc({ rt, at: Date.now() }), { access: 'private', addRandomSuffix: false, allowOverwrite: true, contentType: 'text/plain' }); }
async function loadRT() {
  const { get } = await blob(); let r;
  try { r = await get(TOKEN_PATH, { access: 'private', useCache: false }); } catch (e) { if (/not.?found/i.test(e.name + e.message)) return null; throw e; }
  if (!r) return null;
  const body = r.stream ? await new Response(r.stream).text() : typeof r.text === 'function' ? await r.text() : String(r);
  try { const o = dec(body); lastSaved = o.at || 0; return o.rt; } catch (e) { return null; }
}
let opsCache = null, lastSaved = 0;
async function opsToken() {
  if (opsCache && opsCache.exp > Date.now() + 60e3) return opsCache.token;
  const rt = await loadRT();
  if (!rt) throw new HttpError(503, 'The portal is not connected to the Excel files yet. The opsportal account needs to complete setup.');
  const r = await fetch(`https://login.microsoftonline.com/${TENANT}/oauth2/v2.0/token`, { method: 'POST', body: new URLSearchParams({ client_id: CLIENT_ID, client_secret: SECRET, grant_type: 'refresh_token', refresh_token: rt, scope: SCOPE }) });
  const j = await r.json();
  if (!r.ok) throw new HttpError(503, 'The portal lost its connection to the Excel files (' + (j.error || r.status) + '). The opsportal account needs to run setup again.');
  // Microsoft rotates the refresh token each time; older ones stay valid, so only store a new one once a day (keeps storage use tiny)
  if (j.refresh_token && j.refresh_token !== rt && Date.now() - lastSaved > 24 * 3600e3) { await saveRT(j.refresh_token); lastSaved = Date.now(); }
  opsCache = { token: j.access_token, exp: Date.now() + j.expires_in * 1000 };
  return opsCache.token;
}

/* ---------------- Graph / Excel ---------------- */
const books = {};
const sessions = {}; // workbook base -> { id, at }
async function sessionFor(base) {
  const s = sessions[base]; if (s && Date.now() - s.at < 4 * 60e3) { s.at = Date.now(); return s.id; }
  try { const r = await g('POST', base + '/createSession', { persistChanges: true }, 0, true); sessions[base] = { id: r.id, at: Date.now() }; return r.id; } catch (e) { return null; }
}
async function g(method, path, body, tries = 0, noSession = false) {
  const wb = !noSession && path.match(/^(\/drives\/[^/]+\/items\/[^/]+\/workbook)\//);
  const sid = wb ? await sessionFor(wb[1]) : null;
  const headers = { Authorization: 'Bearer ' + await opsToken(), 'Content-Type': 'application/json' }; if (sid) headers['workbook-session-id'] = sid;
  if (path.startsWith('/shares/')) headers.Prefer = 'redeemSharingLinkIfNecessary'; // lets links meant for specific people work for this account
  const res = await fetch('https://graph.microsoft.com/v1.0' + path, { method, headers, body: body ? JSON.stringify(body) : undefined });
  if (sid && (res.status === 404 || res.status === 400) && tries < 1) { const txt = await res.clone().text(); if (/session/i.test(txt)) { delete sessions[wb[1]]; return g(method, path, body, tries + 1, noSession); } }
  if ([429, 409, 503, 504].includes(res.status) && tries < 4) { await new Promise(r => setTimeout(r, 700 * (tries + 1))); return g(method, path, body, tries + 1, noSession); }
  if (!res.ok) { let t = ''; try { t = (await res.json()).error?.message || ''; } catch (e) { } throw new HttpError(502, `Excel: ${t || res.statusText}`); }
  return res.status === 204 ? null : res.json();
}
async function book(url) {
  if (books[url]) return books[url];
  const direct = url.match(/^https:\/\/([^/]+)(\/personal\/[^/]+)\/Documents\/(.+)$/);
  let it;
  if (direct) { const site = await g('GET', `/sites/${direct[1]}:${direct[2]}?$select=id`); it = await g('GET', `/sites/${site.id}/drive/root:/${direct[3].split('/').map(encodeURIComponent).join('/')}?$select=id,parentReference`); }
  else { const enc = 'u!' + Buffer.from(url.split('?')[0]).toString('base64').replace(/=+$/, '').replace(/\//g, '_').replace(/\+/g, '-'); it = await g('GET', `/shares/${enc}/driveItem?$select=id,parentReference`); }
  return books[url] = `/drives/${it.parentReference.driveId}/items/${it.id}/workbook`;
}
async function table(url, name) { const b = await book(url); const r = await g('GET', `${b}/tables/${encodeURIComponent(name)}/range?$select=values`); return rowsToObjects(r.values); }
async function tableTail(url, name, maxRows) {
  const b = await book(url);
  const hdr = await g('GET', `${b}/tables/${encodeURIComponent(name)}/headerRowRange?$select=values`);
  const body = await g('GET', `${b}/tables/${encodeURIComponent(name)}/dataBodyRange?$select=address`);
  const m = body.address.match(/^(.*)!([A-Z]+)(\d+):([A-Z]+)(\d+)$/); const sheet = m[1].replace(/^'|'$/g, '').replace(/''/g, "'");
  const first = +m[3], last = +m[5]; const start = Math.max(first, last - maxRows + 1);
  const r = await g('GET', `${b}/worksheets('${encodeURIComponent(sheet.replace(/'/g, "''"))}')/range(address='${m[2]}${start}:${m[4]}${last}')?$select=values`);
  const out = rowsToObjects([hdr.values[0], ...r.values]); out.rows.forEach(o => o._i += (start - first)); return out;
}
async function addRow(url, name, values) { const b = await book(url); return g('POST', `${b}/tables/${encodeURIComponent(name)}/rows/add`, { values: [values] }); }
async function setRow(url, name, index, values) {
  const b = await book(url);
  const body = await g('GET', `${b}/tables/${encodeURIComponent(name)}/dataBodyRange?$select=address`);
  const m = body.address.match(/^(.*)!([A-Z]+)(\d+):([A-Z]+)(\d+)$/); const sheet = m[1].replace(/^'|'$/g, '').replace(/''/g, "'"); const row = +m[3] + index;
  return g('PATCH', `${b}/worksheets('${encodeURIComponent(sheet.replace(/'/g, "''"))}')/range(address='${m[2]}${row}:${m[4]}${row}')`, { values: [values] });
}
async function sheetNames(url) { const b = await book(url); return (await g('GET', `${b}/worksheets?$select=name`)).value.map(w => w.name); }
// visible tabs only: hidden tabs in the rota are old weeks kept for reference
async function visibleSheetNames(url) { const b = await book(url); return (await g('GET', `${b}/worksheets?$select=name,visibility`)).value.filter(w => !w.visibility || w.visibility === 'Visible').map(w => w.name); }
async function sheetText(url, sheet) {
  const b = await book(url); const r = await g('GET', `${b}/worksheets('${encodeURIComponent(sheet.replace(/'/g, "''"))}')/usedRange(valuesOnly=true)?$select=text,address`);
  const m = r.address.match(/!([A-Z]+)(\d+)/); let col = 0; for (const ch of m[1]) col = col * 26 + ch.charCodeAt(0) - 64;
  const grid = []; for (let i = 0; i < +m[2] - 1; i++) grid.push([]); r.text.forEach(row => grid.push([...Array(col - 1).fill(''), ...row])); return grid;
}
async function sheetValues(url, sheet) { const b = await book(url); return (await g('GET', `${b}/worksheets('${encodeURIComponent(sheet.replace(/'/g, "''"))}')/usedRange(valuesOnly=true)?$select=values`)).values; }

/* ---------------- short-lived memory cache (per server instance) ---------------- */
const memo = {};
async function cached(key, ttlMs, fn) { const c = memo[key]; if (c && Date.now() - c.at < ttlMs) return structuredClone(c.v); const v = await fn(); memo[key] = { at: Date.now(), v }; return structuredClone(v); }
const forget = (...keys) => keys.forEach(k => delete memo[k]);

/* ---------------- data loaders ---------------- */
async function loadAccess() { return cached('access', 60000.0, loadAccessRaw); }
// "Sondos" or "Sondos Ahmed" → the one team leader / manager whose full name starts with those words (null if none or ambiguous)
function findLeader(name, rows) {
  const toks = norm(name).split(' ').filter(Boolean); if (!toks.length) return null;
  const hits = rows.filter(r => MANAGER_ROLES.includes(String(r['Portal Role'] || '').trim()) && !isNo(r.Active) && r.Email).filter(r => {
    const full = norm(r['Full Name']).split(' '); return toks.every((t, i) => full[i] === t || (t.length >= 5 && full[i] && lev1(t, full[i])));
  });
  return hits.length === 1 ? hits[0] : null;
}
async function loadAccessRaw() {
  const t = await table(URLS.portal, T.access);
  t.rows = t.rows.filter(r => r.Email || r['Full Name']);
  t.rows.forEach(r => { r.Email = normEmail(r.Email); r['Manager Email'] = normEmail(r['Manager Email']); r['Personal Email'] = normEmail(r['Personal Email']); r['Portal Role'] = String(r['Portal Role'] || '').trim(); });
  const byName = new Map(t.rows.map(r => [norm(r['Full Name']), r])); t.rows.forEach(r => { [r['Roster Name'], ...String(r['Previous Name'] || '').split(';')].forEach(x => { if (norm(x) && !byName.has(norm(x))) byName.set(norm(x), r); }); });
  t.rows.forEach(r => { const tl = r['Team Leader'] ? (byName.get(norm(r['Team Leader'])) || findLeader(r['Team Leader'], t.rows)) : null; r._tl = tl ? tl['Full Name'] : (r['Team Leader'] || ''); if (!r['Manager Email'] && tl) r['Manager Email'] = tl.Email; });
  return t;
}
async function loadSettings() { return cached('settings', 60000.0, loadSettingsRaw); }
async function loadSettingsRaw() { const t = await table(URLS.portal, T.set); const o = { BreakMinutes: 60, EmergencyLimit: 6, LateGraceMinutes: 5, LeaveYear: new Date().getFullYear() }; t.rows.forEach(r => { if (r.Key) o[String(r.Key).trim()] = isNaN(+r.Value) ? r.Value : +r.Value; }); return o; }
async function loadAttendance() { return cached('att', 15000.0, loadAttendanceRaw); }
async function loadAttendanceRaw() { const t = await tableTail(URLS.portal, T.att, 20000); t.rows = t.rows.filter(r => r.ID).map(r => ({ ...r, Email: normEmail(r.Email), 'Shift Date': toKey(r['Shift Date']), Timestamp: tsString(r.Timestamp) })); return t; }
async function loadRequests() { return cached('req', 15000.0, loadRequestsRaw); }
async function loadRequestsRaw() { let t = await table(URLS.portal, T.req);
  for (const name of ['Exception', 'Exception Note']) if (!t.headers.includes(name)) { try { const b = await book(URLS.portal); await g('POST', `${b}/tables/${T.req}/columns/add`, { name }); t = await table(URLS.portal, T.req); } catch (e) { console.error('requests column', name, e.message); } }
  t.rows = t.rows.filter(r => r.ID); t.rows.forEach(r => { r.Email = normEmail(r.Email); r['Submitted By'] = normEmail(r['Submitted By']); r['Submitted At'] = tsString(r['Submitted At']); r.dateList = String(r.Dates || '').split(/[,\s]+/).filter(Boolean); }); return t; }
async function loadSwaps() { return cached('swaps', 15000.0, loadSwapsRaw); }
async function loadSwapsRaw() { let t; try { t = await table(URLS.portal, T.swap); } catch (e) { return { headers: [], rows: [], missing: true }; }
  for (const col of ['Type', 'Written To Rota']) if (!t.headers.includes(col)) { try { const b = await book(URLS.portal); await g('POST', `${b}/tables/${T.swap}/columns/add`, { name: col }); t = await table(URLS.portal, T.swap); } catch (e) { console.error('swap column', col, e.message); } } t.rows = t.rows.filter(r => r.ID); t.rows.forEach(r => { r['Requester Email'] = normEmail(r['Requester Email']); r['Teammate Email'] = normEmail(r['Teammate Email']); r['Requester Date'] = toKey(r['Requester Date']); r['Teammate Date'] = toKey(r['Teammate Date']); r['Submitted At'] = tsString(r['Submitted At']); }); return t; }
async function loadTracker() { return cached('tracker', 60000.0, loadTrackerRaw); }
async function loadTrackerRaw() {
  const t = await table(URLS.tracker, TRACKER_TABLE); const H = t.headers; const col = n => H.find(h => h.toLowerCase() === n.toLowerCase());
  const cN = col('Name'), cV = col('Vacation Date'), cR = col('Role'), cT = col('Leave type'), cS = col('Status'), cB = col('Requested by'), cC = col('Additional Comment');
  t.items = t.rows.filter(r => r[cN] && String(r[cN]).trim()).map(r => {
    const d = parseDateCell(r[cV]); let type = String(r[cT] || '').trim(); if (/^sick/i.test(type)) type = 'Sick'; const status = String(r[cS] || '').trim();
    return { _i: r._i, name: String(r[cN]), n: norm(r[cN]), date: d ? dkey(d) : null, role: String(r[cR] || '').trim(), type, status, by: String(r[cB] || ''), comment: String(r[cC] || ''), approved: /^approved/i.test(status) && !/cancel/i.test(status), pending: status === '', cancelled: /cancel|reject/i.test(status) || /cancel/i.test(type) };
  });
  return t;
}
async function loadBalance() { return cached('balance', 300000.0, loadBalanceRaw); }
async function loadBalanceRaw() {
  const v = await sheetValues(URLS.tracker, BALANCE_SHEET); const h = v[0].map(x => String(x ?? '').trim().toLowerCase()); const last = n => h.lastIndexOf(n);
  const cA = last('agent'), cAl = last('allowance'), cS = last('start date'), cAs = last('allowance start date'); const out = {};
  v.slice(1).forEach(r => { if (r[cA]) { const s = parseDateCell(r[cS]), as = parseDateCell(r[cAs]); out[norm(r[cA])] = { allowance: +r[cAl] || 0, start: s ? dkey(s) : null, allowanceStart: as ? dkey(as) : null }; } });
  return out;
}
async function loadWeeks() { return cached('weeks', 300000.0, loadWeeksRaw); }
async function loadWeeksRaw() { return (await visibleSheetNames(URLS.rota)).map(n => ({ name: n, start: sheetWeekStart(n) })).filter(x => x.start && /^w[kc]/i.test(x.name.trim())).sort((a, b) => a.start - b.start).map(w => ({ name: w.name, start: dkey(w.start) })); }

/* ---------------- who is calling ---------------- */
let jwks = null, jwksAt = 0;
async function signingKey(kid) {
  if (!jwks || Date.now() - jwksAt > 3600e3 || !jwks.some(k => k.kid === kid)) { jwks = (await (await fetch(`https://login.microsoftonline.com/${TENANT}/discovery/v2.0/keys`)).json()).keys || []; jwksAt = Date.now(); }
  const k = jwks.find(x => x.kid === kid); if (!k) throw new Error('unknown key');
  return crypto.createPublicKey({ key: { kty: k.kty, n: k.n, e: k.e }, format: 'jwk' });
}
async function verifyIdToken(tok) {
  const [h, p, sig] = String(tok).split('.'); if (!sig) throw new Error('malformed');
  const header = JSON.parse(Buffer.from(h, 'base64url').toString()); if (header.alg !== 'RS256') throw new Error('alg');
  const ok = crypto.verify('RSA-SHA256', Buffer.from(h + '.' + p), await signingKey(header.kid), Buffer.from(sig, 'base64url')); if (!ok) throw new Error('signature');
  const payload = JSON.parse(Buffer.from(p, 'base64url').toString()); const now = Date.now() / 1000;
  if (payload.iss !== `https://login.microsoftonline.com/${TENANT}/v2.0` || payload.aud !== CLIENT_ID || payload.tid !== TENANT) throw new Error('claims');
  if (payload.exp < now - 300 || (payload.nbf && payload.nbf > now + 300)) throw new Error('expired');
  return payload;
}
async function caller(req, acc) {
  const tok = String(req.headers.authorization || '').replace(/^Bearer\s+/i, '');
  if (!tok) throw new HttpError(401, 'Please sign in.');
  let p; try { p = await verifyIdToken(tok); } catch (e) { throw new HttpError(401, 'Your sign-in has expired. Refresh the page.'); }
  const email = normEmail(p.preferred_username || p.email || p.upn);
  const me = acc.rows.find(r => r.Email === email) || acc.rows.find(r => r['Personal Email'] && r['Personal Email'] === email);
  if (!me || isNo(me.Active) || !ROLES.includes(me['Portal Role'])) throw new HttpError(403, `${email} isn't on the portal's access list. Ask your Duty Manager to add you.`);
  return { email: me.Email || email, name: me['Full Name'], role: me['Portal Role'], admin: yes(me['Portal Admin']), row: me };
}
const isMgr = u => MANAGER_ROLES.includes(u.role);
// swaps: Duty Managers and above, plus anyone the admin ticks 'Can Approve Swaps' for. Rota edits: admins, the GM, or 'Can Edit Rota'.
const canApproveSwaps = u => ['Duty Manager', 'General Manager'].includes(u.role) || yes(u.row['Can Approve Swaps']);
const canEditRota = u => u.admin || u.role === 'General Manager' || yes(u.row['Can Edit Rota']);
// every name a person goes by in the Leaves Tracker: their current name, plus earlier names kept in 'Previous Name' after a rename
const nameKeys = p => [...new Set([p && p['Full Name'], ...String((p && p['Previous Name']) || '').split(';')].map(norm).filter(Boolean))];
const teamEmails = (u, acc) => new Set(acc.rows.filter(r => r['Manager Email'] === u.email).map(r => r.Email).filter(Boolean));
const seesAll = u => ['Duty Manager', 'General Manager', 'HR'].includes(u.role);
const canRaiseForOthers = u => ['Team Leader', 'Duty Manager', 'General Manager'].includes(u.role);
const publicPerson = r => ({ _i: r._i, Email: r.Email, 'Full Name': r['Full Name'], 'Portal Role': r['Portal Role'], 'Team Leader': r['Team Leader'], 'Manager Email': r['Manager Email'], Team: r.Team, 'Tracker Role': r['Tracker Role'], 'Rota Name': r['Rota Name'], Active: r.Active, 'Roster Name': r['Roster Name'], 'Previous Name': r['Previous Name'], 'Quality Team': r['Quality Team'], 'Can Approve Swaps': r['Can Approve Swaps'], 'Can Edit Rota': r['Can Edit Rota'], _tl: r._tl });

/* ---------------- read: everything the page needs, filtered by role ---------------- */
async function readData(u, acc, parts) {
  const want = p => !parts || parts.includes(p); const out = { me: { email: u.email, name: u.name, role: u.role, admin: u.admin, quality: yes(u.row['Quality Team']) } };
  const team = teamEmails(u, acc);
  if (want('access')) {
    let rows;
    if (u.admin) rows = acc.rows; else if (isMgr(u)) rows = acc.rows.map(publicPerson);
    else rows = acc.rows.filter(r => r.Email === u.email || (u.row['Manager Email'] && r['Manager Email'] === u.row['Manager Email'] && r['Portal Role'] === 'Agent' && !isNo(r.Active)) || (r.Email && r.Email === u.row['Manager Email'])).map(publicPerson);
    out.access = { headers: acc.headers, rows };
  }
  if (want('settings')) out.settings = await loadSettings();
  const jobs = []; out.warnings = [];
  const safe = (label, pr) => pr.catch(e => { out.warnings.push(`${label}: ${e.message}`); });
  if (want('attendance')) jobs.push(safe('Log-ins', loadAttendance().then(t => { t.rows = seesAll(u) ? t.rows : t.rows.filter(r => r.Email === u.email || (u.role === 'Team Leader' && team.has(r.Email))); out.attendance = t; })));
  if (want('requests')) jobs.push(safe('Leave requests', loadRequests().then(t => { t.rows = seesAll(u) ? t.rows : t.rows.filter(r => r.Email === u.email || r['Submitted By'] === u.email || (u.role === 'Team Leader' && team.has(r.Email))); out.requests = t; })));
  if (want('swaps')) jobs.push(safe('Swaps', loadSwaps().then(t => { t.rows = (seesAll(u) || canApproveSwaps(u)) ? t.rows : t.rows.filter(r => [r['Requester Email'], r['Teammate Email']].includes(u.email) || (u.role === 'Team Leader' && (team.has(r['Requester Email']) || team.has(r['Teammate Email'])))); out.swaps = t; })));
  if (want('tracker')) jobs.push(safe('Leaves Tracker', loadTracker().then(t => {
    const mine = new Set([norm(u.name), ...nameKeys(u.row)]);
    const items = isMgr(u) ? t.items : t.items.map(x => mine.has(x.n) ? x : { _i: -1, name: '', n: '', date: x.date, role: x.role, type: '', status: x.approved ? 'Approved' : x.status ? 'Other' : '', by: '', comment: '', approved: x.approved, pending: x.pending, cancelled: x.cancelled });
    out.tracker = { headers: [], items };
  })));
  if (want('balance')) jobs.push(safe('Leaves Tracker (Balance tab)', loadBalance().then(b => { out.balance = isMgr(u) ? b : Object.fromEntries([norm(u.name), ...nameKeys(u.row)].filter(k => b[k]).map(k => [k, b[k]])); })));
  if (want('rotaedits')) jobs.push(safe('Schedule changes', loadRotaEdits().then(t => { out.rotaedits = { rows: t.rows.map(r => ({ Email: r.Email, Date: r.Date, Shift: r.Shift, Original: r.Original, Note: r.Note, 'Edited By': r['Edited By'], 'Edited At': r['Edited At'], Synced: r.Synced })) }; })));
  if (want('weeks')) jobs.push(safe('Rota', loadWeeks().then(w => { out.weeks = w; })));
  await Promise.all(jobs);
  return out;
}

/* ---------------- writes ---------------- */
function rowFor(headers, o) { return headers.map(k => o[k] ?? ''); }
async function doAttendance(u, acc, b) {
  const action = String(b.action || '');
  if (!['Clock In', 'Break Start', 'Break End', 'Coaching Start', 'Coaching End', 'Clock Out'].includes(action)) bad('Unknown action.');
  const target = normEmail(b.email || u.email);
  const person = acc.rows.find(r => r.Email === target); if (!person) bad('Person not found.');
  let ts = new Date(), note = String(b.note || '').slice(0, 300);
  if (target !== u.email) {
    const team = teamEmails(u, acc);
    if (!(seesAll(u) && u.role !== 'HR') && !(u.role === 'Team Leader' && team.has(target))) deny('Only Duty Managers, the GM, or the person\'s team leader can add entries for someone else.');
    if (!note.trim()) bad('A reason is required for manual entries.');
    const t = new Date(b.ts); if (isNaN(t)) bad('Invalid time.'); ts = t; note = 'Manual: ' + note.replace(/^Manual:\s*/i, '');
  }
  const shiftDate = isKey(b.shiftDate) ? b.shiftDate : dkey(new Date(ts.getTime() + 2 * 3600e3));
  if (Math.abs(fromKey(shiftDate) - ts) > 3 * 864e5) bad('Shift date is too far from the entry time.');
  const t = await tableTail(URLS.portal, T.att, 1);
  const row = { ID: uid('A'), Email: person.Email, Name: person['Full Name'], 'Shift Date': shiftDate, Action: action, Timestamp: ts.toISOString(), 'Recorded By': u.email, Note: note };
  await addRow(URLS.portal, T.att, rowFor(t.headers, row));
  return { row };
}
async function ptoLeft(person, reqs, excludeId) {
  const keys = nameKeys(person); const [bal, tr] = await Promise.all([loadBalance(), loadTracker()]); const b = keys.map(k => bal[k]).find(Boolean); if (!b) return null;
  const used = tr.items.filter(x => keys.includes(x.n) && x.approved && (x.type === 'Holiday' || x.type === 'Emergency')).length;
  const pending = reqs.filter(r => r.Email === person.Email && r.ID !== excludeId && /^Pending/.test(r.Status) && ['Holiday', 'Emergency'].includes(r['Leave Type'])).reduce((a, r) => a + r.dateList.length, 0);
  return { allowance: b.allowance, used, pending, left: b.allowance - used - pending };
}
const plural = (n, w) => { n = +(+n).toFixed(1); return `${n} ${w}${n === 1 ? '' : 's'}`; };
/* ---- holiday capacity: the most people from one group who may be off on the same day ---- */
const CAPACITY = [
  { key: 'enquiry', label: 'Enquiry agents', cap: 7, re: /^enquiry agent/ },
  { key: 'resv', label: 'Reservation agents', cap: 2, re: /^reservation agent|^resv agent/ },
  { key: 'dm', label: 'Duty Managers', cap: 1, re: /duty man/ },
  { key: 'tl-enq', label: 'Enquiry team leaders', cap: 2, re: /^tl\s*-?\s*enq/ },
  { key: 'tl-resv', label: 'Reservation team leaders', cap: 1, re: /^tl\s*-?\s*res/ },
];
const capGroup = role => { const r = String(role || '').trim().toLowerCase().replace(/\s+/g, ' '); const hit = CAPACITY.find(c => c.re.test(r)); return hit ? { key: hit.key, label: hit.label, cap: hit.cap } : { key: 'other:' + (r || 'unknown'), label: String(role || 'This group').trim() || 'This group', cap: 1 }; };
const LEADS = ['Team Leader', 'Duty Manager', 'General Manager', 'HR'];
async function capacityFull(person, dates, reqs) {
  const grp = capGroup(person['Tracker Role']); const tr = await loadTracker(); const mine = new Set(nameKeys(person)); const per = {};
  const add = (d, n) => { if (n && !mine.has(n)) (per[d] = per[d] || new Set()).add(n); };
  tr.items.forEach(x => { if (x.date && dates.includes(x.date) && !x.cancelled && (x.approved || x.pending) && !/cancell/i.test(x.type) && capGroup(x.role).key === grp.key) add(x.date, x.n); });
  reqs.forEach(r => { if (/^Pending/.test(r.Status) && r.Email !== person.Email && capGroup(r['Tracker Role']).key === grp.key) r.dateList.forEach(d => { if (dates.includes(d)) add(d, norm(r.Name)); }); });
  return { group: grp, full: dates.filter(d => (per[d] ? per[d].size : 0) >= grp.cap).map(d => ({ date: d, off: per[d].size, cap: grp.cap })) };
}
async function doRequest(u, acc, b) {
  const target = normEmail(b.email || u.email); const person = acc.rows.find(r => r.Email === target); if (!person) bad('Person not found.');
  if (target !== u.email && !canRaiseForOthers(u)) deny('You can only request leave for yourself.');
  const type = String(b.type || ''); const allowed = canRaiseForOthers(u) ? ['Holiday', 'Emergency', 'UPL', 'Sick'] : ['Holiday', 'Emergency'];
  if (!allowed.includes(type)) deny(`You can't request ${type} leave.`);
  const dates = [...new Set((b.dates || []).filter(isKey))].sort(); if (!dates.length || dates.length > 62) bad('Pick between 1 and 62 days.');
  const t = await loadRequestsRaw();
  if (t.rows.some(r => r.Email === target && /^Pending|^Approved/.test(r.Status) && r.dateList.some(d => dates.includes(d)))) bad('There is already a request covering one of these days.');
  if (type === 'Holiday') { const bal = await ptoLeft(person, t.rows); if (bal && dates.length > bal.left) bad(`Not enough PTO balance: ${person['Full Name']} has ${plural(Math.max(0, bal.left), 'day')} left${bal.pending ? ` (after ${plural(bal.pending, 'day')} already waiting for approval)` : ''}, and this request is for ${plural(dates.length, 'day')}.`); }
  let exception = false, exNote = '';
  if (['Holiday', 'UPL'].includes(type)) {
    const cap = await capacityFull(person, dates, t.rows);
    if (cap.full.length) {
      const txt = cap.full.map(f => `${f.date} (${f.off} of ${f.cap} already off)`).join(', ');
      if (!LEADS.includes(u.role)) { const e = new HttpError(409, `${cap.group.label} are at capacity on ${txt}. Please pick other days.`); e.capacity = { blocked: true, ...cap }; throw e; }
      if (b.check) return { ok: true, check: true, exception: true, blocked: false, ...cap };
      if (b.exception !== true) { const e = new HttpError(409, `${cap.group.label} are at capacity on ${txt}. This can be sent as an exception request.`); e.capacity = { needsException: true, ...cap }; throw e; }
      exception = true; exNote = `Over capacity (${cap.group.label}, max ${cap.group.cap}): ${txt}`;
    }
    if (b.check) return { ok: true, check: true, exception: false, blocked: false, ...cap };
  } else if (b.check) return { ok: true, check: true, exception: false, blocked: false, full: [] };
  const row = { ID: uid('R'), 'Submitted At': new Date().toISOString(), Email: target, Name: person['Full Name'], 'Tracker Role': person['Tracker Role'], 'Leave Type': type, Dates: dates.join(', '), Days: dates.length, Reason: String(b.reason || '').slice(0, 300), 'Submitted By': u.email, Status: (exception || REPORTS_TO_GM.includes(person['Portal Role'])) ? 'Pending GM' : 'Pending Manager', 'Written To Tracker': 'No', Exception: exception ? 'Yes' : '', 'Exception Note': exNote };
  await addRow(URLS.portal, T.req, rowFor(t.headers, row));
  return { ok: true, id: row.ID };
}
async function updateReq(t, r, patch) { Object.assign(r, patch); await setRow(URLS.portal, T.req, r._i, t.headers.map(k => k === 'Dates' ? r.dateList.join(', ') : (r[k] ?? ''))); }
async function writeTracker(t, r, acc) {
  if (isTestMode(await loadSettings())) { await updateReq(t, r, { 'Written To Tracker': 'Test' }); return; } // test mode: the Leaves Tracker is left alone
  const tr = await table(URLS.tracker, TRACKER_TABLE); const H = tr.headers;
  let lastFilled = -1; tr.rows.forEach(x => { if (H.some(h => x[h] !== '' && x[h] != null)) lastFilled = x._i; });
  let next = lastFilled + 1; const by = (acc.rows.find(x => x.Email === r['Submitted By']) || {})['Full Name'] || r['Submitted By'];
  for (const d of r.dateList) {
    const vals = H.map(h => { const k = h.toLowerCase(); return k === 'date' ? dateToSerial(new Date()) : k === 'vacation date' ? dateToSerial(fromKey(d)) : k === 'name' ? r.Name : k === 'role' ? r['Tracker Role'] : k === 'requested by' ? by : k === 'leave type' ? r['Leave Type'] : k === 'status' ? 'Approved ' : k === 'additional comment' ? ('Portal ' + r.ID + (r.Reason ? ' – ' + r.Reason : '')) : ''; });
    if (next < tr.rows.length) await setRow(URLS.tracker, TRACKER_TABLE, next, vals); else await addRow(URLS.tracker, TRACKER_TABLE, vals);
    next++;
  }
  await updateReq(t, r, { 'Written To Tracker': 'Yes' });
}
async function doDecide(u, acc, b) {
  const t = await loadRequestsRaw(); const r = t.rows.find(x => x.ID === b.id); if (!r) bad('Request not found.');
  if (r.Email === u.email) deny("You can't approve your own request.");
  const ok = !!b.ok, note = String(b.note || '').slice(0, 300), now = new Date().toISOString();
  if (r.Status === 'Pending Manager' && REPORTS_TO_GM.includes((acc.rows.find(x => x.Email === r.Email) || {})['Portal Role'])) {
    // older request from a leader / Duty Manager / HR: only the General Manager decides, and that is the final approval
    if (u.role !== 'General Manager') deny('This person reports to the General Manager, so only the General Manager can approve it.');
    if (!ok) { await updateReq(t, r, { Status: 'Rejected', 'Stage 2 By': u.email, 'Stage 2 At': now, 'Stage 2 Note': note }); return { status: 'Rejected' }; }
    if (r['Leave Type'] === 'Holiday') { const p = acc.rows.find(x => x.Email === r.Email); const bal = p && await ptoLeft(p, t.rows, r.ID); if (bal && r.dateList.length > bal.left) bad(`Can't approve: ${r.Name} only has ${plural(Math.max(0, bal.left), 'PTO day')} left. Reject it so they can request fewer days.`); }
    await updateReq(t, r, { Status: 'Approved', 'Stage 2 By': u.email, 'Stage 2 At': now, 'Stage 2 Note': note });
    try { await writeTracker(t, r, acc); } catch (e) { throw new HttpError(502, 'Approved, but writing to the Leaves Tracker failed: ' + e.message + '. Use "Retry tracker write".'); }
    return { status: 'Approved' };
  }
  if (r.Status === 'Pending Manager') {
    const mgr = (acc.rows.find(x => x.Email === r.Email) || {})['Manager Email'];
    if (r['Submitted By'] === u.email) deny("You raised this request, so another manager has to approve it.");
    if (!(u.role === 'Duty Manager' || (u.role === 'Team Leader' && mgr === u.email))) deny('Only the team leader or a Duty Manager can approve this step.');
    if (ok && r['Leave Type'] === 'Holiday') { const p = acc.rows.find(x => x.Email === r.Email); const bal = p && await ptoLeft(p, t.rows, r.ID); if (bal && r.dateList.length > bal.left) bad(`Can't approve: ${r.Name} only has ${plural(Math.max(0, bal.left), 'PTO day')} left. Reject it so they can request fewer days.`); }
    if (!ok) { await updateReq(t, r, { Status: 'Rejected', 'Stage 1 By': u.email, 'Stage 1 At': now, 'Stage 1 Note': note }); return { status: 'Rejected' }; }
    const isGM = (acc.rows.find(x => x.Email === r.Email) || {})['Portal Role'] === 'General Manager';
    await updateReq(t, r, { Status: isGM ? 'Approved' : 'Pending GM', 'Stage 1 By': u.email, 'Stage 1 At': now, 'Stage 1 Note': note });
    if (isGM) await writeTracker(t, r, acc);
    return { status: r.Status };
  }
  if (r.Status === 'Pending GM') {
    if (u.role !== 'General Manager') deny('Only the General Manager can give final approval.');
    if (ok && r['Leave Type'] === 'Holiday') { const p = acc.rows.find(x => x.Email === r.Email); const bal = p && await ptoLeft(p, t.rows, r.ID); if (bal && r.dateList.length > bal.left) bad(`Can't approve: ${r.Name} only has ${plural(Math.max(0, bal.left), 'PTO day')} left. Reject it so they can request fewer days.`); }
    if (!ok) { await updateReq(t, r, { Status: 'Rejected', 'Stage 2 By': u.email, 'Stage 2 At': now, 'Stage 2 Note': note }); return { status: 'Rejected' }; }
    await updateReq(t, r, { Status: 'Approved', 'Stage 2 By': u.email, 'Stage 2 At': now, 'Stage 2 Note': note });
    try { await writeTracker(t, r, acc); } catch (e) { throw new HttpError(502, 'Approved, but writing to the Leaves Tracker failed: ' + e.message + '. Use "Retry tracker write".'); }
    return { status: 'Approved' };
  }
  bad('This request is no longer waiting for approval.');
}
async function doCancel(u, b) { const t = await loadRequestsRaw(); const r = t.rows.find(x => x.ID === b.id); if (!r) bad('Request not found.'); if (r.Email !== u.email && r['Submitted By'] !== u.email) deny(); if (!/^Pending/.test(r.Status)) bad('Only pending requests can be cancelled.'); await updateReq(t, r, { Status: 'Cancelled' }); return { ok: true }; }
async function doRetry(u, acc, b) { if (!['Duty Manager', 'General Manager'].includes(u.role)) deny(); const t = await loadRequestsRaw(); const r = t.rows.find(x => x.ID === b.id); if (!r || r.Status !== 'Approved' || r['Written To Tracker'] === 'Yes') bad('Nothing to retry.'); await writeTracker(t, r, acc); return { ok: true }; }
async function doSwap(u, acc, b) {
  const mate = acc.rows.find(r => r.Email === normEmail(b.teammate)); if (!mate || isNo(mate.Active)) bad('Colleague not found.');
  if (mate.Email === u.email) bad('Pick a colleague to swap with.');
  const dept = x => norm(x.Team);
  if (!dept(u.row) || dept(u.row) !== dept(mate) || String(mate['Portal Role']).trim() !== u.role) deny(!dept(u.row) ? "Your department isn't set yet. Ask your Duty Manager to add it in Access." : 'You can only swap with someone in the same role and department (' + u.role + ', ' + u.row.Team + ').');
  const type = b.type === 'dayoff' ? 'Day-off swap' : 'Shift swap';
  if (type === 'Shift swap' && b.myDate !== b.theirDate) bad('A shift swap is for one day. Choose the same date, or use a day-off swap.');
  if (type === 'Day-off swap' && b.myDate === b.theirDate) bad('A day-off swap needs two different days.');
  if (!isKey(b.myDate) || !isKey(b.theirDate)) bad('Pick both days.');
  const t = await loadSwapsRaw(); if (t.missing) bad('The data file has no Swaps table.');
  if (t.rows.some(r => r.Status === 'Pending' && [r['Requester Email'], r['Teammate Email']].some(e => [u.email, mate.Email].includes(e)) && [r['Requester Date'], r['Teammate Date']].some(x => [b.myDate, b.theirDate].includes(x)))) bad('There is already a pending swap on one of these days.');
  const row = { ID: uid('S'), 'Submitted At': new Date().toISOString(), 'Requester Email': u.email, 'Requester Name': u.name, 'Requester Date': b.myDate, 'Requester Shift': String(b.myShift || '').slice(0, 80), 'Teammate Email': mate.Email, 'Teammate Name': mate['Full Name'], 'Teammate Date': b.theirDate, 'Teammate Shift': String(b.theirShift || '').slice(0, 80), Reason: String(b.reason || '').slice(0, 300), Status: 'Pending', Type: type };
  await addRow(URLS.portal, T.swap, rowFor(t.headers, row));
  return { ok: true };
}
async function doSwapUpdate(u, b, decide) {
  const t = await loadSwapsRaw(); const r = t.rows.find(x => x.ID === b.id); if (!r) bad('Swap not found.'); if (r.Status !== 'Pending') bad('This swap is no longer pending.');
  let patch;
  if (decide) { if (!canApproveSwaps(u)) deny('You are not set up to approve swaps. Ask a portal admin.'); if ([r['Requester Email'], r['Teammate Email']].includes(u.email)) deny("You can't approve a swap you're part of."); patch = { Status: b.ok ? 'Approved' : 'Declined', 'Decided By': u.email, 'Decided At': new Date().toISOString(), Note: String(b.note || '').slice(0, 300) }; }
  else { if (r['Requester Email'] !== u.email) deny(); patch = { Status: 'Cancelled' }; }
  Object.assign(r, patch); await setRow(URLS.portal, T.swap, r._i, t.headers.map(k => r[k] ?? ''));
  if (decide && b.ok) { const acc = await loadAccessRaw(); const st = await syncSwap(t, r, acc, u.email); return { ok: true, rota: st === 'Yes' ? 'written' : st === 'Test' ? 'test' : 'failed', rotaDetail: /^Failed/.test(st) ? st : undefined }; }
  return { ok: true };
}
/* ---------------- backup sync: the Excel files stay the source of truth ----------------
   Whatever is done in the portal is also written to the rota, the Leaves Tracker and the Promotions sheet,
   so the sheets are complete if the portal is ever unavailable. Nothing is written in test mode.
   A failed write never blocks the action: it is logged (Sync Log) and can be retried from the Access tab. */
const SYNC_HEADERS = ['ID', 'At', 'Kind', 'Ref', 'Summary', 'Status', 'Detail', 'By'];
async function syncLog(kind, ref, summary, status, detail, by) {
  try { await ensureTable(URLS.portal, 'Sync Log', 'tblSyncLog', SYNC_HEADERS); await addRow(URLS.portal, 'tblSyncLog', [uid('Y'), new Date().toISOString(), kind, ref, String(summary).slice(0, 200), status, String(detail || '').slice(0, 400), by || '']); forget('synclog'); }
  catch (e) { console.error('sync log', e.message); }
}
const colLetters = n => { let t = '', x = n + 1; while (x > 0) { const m = (x - 1) % 26; t = String.fromCharCode(65 + m) + t; x = Math.floor((x - 1) / 26); } return t; };
const cellAddr = (r, c) => colLetters(c) + (r + 1);
const wsOf = async (url, sheet) => `${await book(url)}/worksheets('${encodeURIComponent(sheet.replace(/'/g, "''"))}')`;
async function readCell(url, sheet, r, c) { return (await g('GET', `${await wsOf(url, sheet)}/range(address='${cellAddr(r, c)}')?$select=values`)).values[0][0]; }
async function writeCell(url, sheet, r, c, v, mark) {
  const w = await wsOf(url, sheet), a = cellAddr(r, c);
  await g('PATCH', `${w}/range(address='${a}')`, { values: [[v]] });
  if (mark) { try { await g('PATCH', `${w}/range(address='${a}')/format/fill`, { color: '#FFF2CC' }); } catch (e) { /* colour is only a visual hint */ } }
}
const PC_HEADERS = ['At', 'Change', 'Person', 'Date', 'Was', 'Now', 'By', 'Ref'];
async function portalChange(rows) {
  try { await ensureTable(URLS.rota, 'Portal Changes', 'tblPortalChanges', PC_HEADERS); for (const r of rows) await addRow(URLS.rota, 'tblPortalChanges', r); }
  catch (e) { console.error('portal changes tab', e.message); }
}
async function rotaTabFor(d) { let pick = null; const day = fromKey(d); for (const w of await loadWeeks()) if (fromKey(w.start) <= day) pick = w; return pick && (day - fromKey(pick.start)) <= 13 * 864e5 ? pick : null; }
async function loadRotaGrid(cache, sheet) { if (!cache.has(sheet)) { const grid = await sheetText(URLS.rota, sheet); cache.set(sheet, { grid, people: parseRotaSheet(grid, sheetWeekStart(sheet)) }); } return cache.get(sheet); }
function rotaFind(people, row) {
  const keys = new Set([...nameKeys(row), norm(row['Rota Name']), norm(row['Roster Name'])].filter(Boolean));
  const hits = people.filter(p => keys.has(norm(p.name)));
  if (hits.length > 1) { const t = norm(row.Team); const f = hits.filter(p => t && norm(p.team).includes(t)); return f.length === 1 ? f[0] : 'ambiguous'; }
  return hits[0] || null;
}
async function rotaLocate(row, d, cache) {
  const w = await rotaTabFor(d); if (!w) throw new Error(`there is no rota tab for ${d}`);
  const { people } = await loadRotaGrid(cache, w.name); const p = rotaFind(people, row);
  if (!p || p === 'ambiguous') throw new Error(`${row['Full Name']} ${p ? 'matches more than one row' : "isn't found"} in rota tab ${w.name}. Set their Rota Name in Access`);
  const at = p.at[d]; if (!at) throw new Error(`${d} is not a day column in ${w.name}`);
  return { sheet: w.name, r: at[0], c: at[1], text: p.days[d] };
}
async function rotaApplySwap(sw, acc) {
  const A = acc.rows.find(x => x.Email === sw['Requester Email']), B = acc.rows.find(x => x.Email === sw['Teammate Email']); if (!A || !B) throw new Error('one of the people is not on the Access list');
  const cache = new Map(), plan = [];
  for (const d of [...new Set([sw['Requester Date'], sw['Teammate Date']])]) plan.push({ d, la: await rotaLocate(A, d, cache), lb: await rotaLocate(B, d, cache) });
  for (const p of plan) { p.va = await readCell(URLS.rota, p.la.sheet, p.la.r, p.la.c); p.vb = await readCell(URLS.rota, p.lb.sheet, p.lb.r, p.lb.c); }
  const done = [];
  for (const p of plan) { await writeCell(URLS.rota, p.la.sheet, p.la.r, p.la.c, p.vb, true); done.push(p); await writeCell(URLS.rota, p.lb.sheet, p.lb.r, p.lb.c, p.va, true); }
  for (const sh of new Set(plan.flatMap(p => [p.la.sheet, p.lb.sheet]))) forget('rota:' + sh);
  const now = new Date().toISOString();
  await portalChange(plan.flatMap(p => [[now, sw.Type || 'Swap', A['Full Name'], p.d, String(p.va), String(p.vb), sw['Decided By'] || '', sw.ID], [now, sw.Type || 'Swap', B['Full Name'], p.d, String(p.vb), String(p.va), sw['Decided By'] || '', sw.ID]]));
}
async function syncSwap(t, sw, acc, by) {
  const label = `${sw.Type || 'Swap'}: ${sw['Requester Name']} and ${sw['Teammate Name']}`;
  let status;
  if (isTestMode(await loadSettings())) status = 'Test';
  else { try { await rotaApplySwap(sw, acc); status = 'Yes'; await syncLog('swap', sw.ID, label, 'Ok', 'Written to the rota', by); } catch (e) { status = 'Failed: ' + e.message.slice(0, 160); await syncLog('swap', sw.ID, label, 'Failed', e.message, by); } }
  sw['Written To Rota'] = status; await setRow(URLS.portal, T.swap, sw._i, t.headers.map(k => sw[k] ?? '')); forget('swaps');
  return status;
}
async function rotaWriteEdit(row, d, text) {
  const cache = new Map(); const l = await rotaLocate(row, d, cache); const was = await readCell(URLS.rota, l.sheet, l.r, l.c);
  await writeCell(URLS.rota, l.sheet, l.r, l.c, text, true); forget('rota:' + l.sheet); return { was: String(was ?? ''), l };
}
async function syncNewPerson(row, by) {
  const out = []; const name = row['Full Name']; const today = new Date(); today.setHours(0, 0, 0, 0);
  if (isTestMode(await loadSettings())) return [['rota', 'Test'], ['balance', 'Test']];
  // 1. Leaves Tracker, Balance sheet
  try {
    const b = await book(URLS.tracker); const ws = await wsOf(URLS.tracker, BALANCE_SHEET); const ur = await g('GET', `${ws}/usedRange(valuesOnly=true)?$select=address,values`);
    const h = ur.values[0].map(x => String(x ?? '').trim().toLowerCase()); const cA = h.lastIndexOf('agent'); if (cA < 0) throw new Error('no Agent column in the Balance sheet');
    if (ur.values.slice(1).some(r => norm(r[cA]) === norm(name))) out.push(['balance', 'Already there']);
    else {
      const m = ur.address.match(/!([A-Z]+)(\d+):([A-Z]+)(\d+)$/) || ur.address.match(/!([A-Z]+)(\d+)/); const startRow = +m[2]; const next = startRow + ur.values.length;
      const st = await loadSettings(); const vals = h.map(k => k === 'agent' ? name : k === 'role' ? (row['Tracker Role'] || '') : k === 'allowance' && st.DefaultAllowance ? +st.DefaultAllowance : '');
      const last = colLetters(vals.length - 1); await g('PATCH', `${ws}/range(address='A${next}:${last}${next}')`, { values: [vals] }); forget('balance'); out.push(['balance', 'Added (check allowance and start date)']);
    }
    await syncLog('person', row.Email, `${name}: Balance sheet`, 'Ok', out.at(-1)[1], by);
  } catch (e) { out.push(['balance', 'Failed: ' + e.message.slice(0, 120)]); await syncLog('person', row.Email, `${name}: Balance sheet`, 'Failed', e.message, by); }
  // 2. rota: this week and next week, in the same block as a colleague of the same team
  try {
    const acc = await loadAccessRaw(); const mates = acc.rows.filter(x => x.Email !== row.Email && !isNo(x.Active) && norm(x.Team) && norm(x.Team) === norm(row.Team) && x['Portal Role'] === row['Portal Role']);
    const weeks = (await loadWeeks()).filter(w => fromKey(w.start) >= new Date(today.getTime() - 6 * 864e5)).slice(0, 2); if (!weeks.length) throw new Error('no current rota tab');
    const cache = new Map(); const res = [];
    for (const w of weeks) {
      const { grid, people } = await loadRotaGrid(cache, w.name);
      if (rotaFind(people, row) && rotaFind(people, row) !== 'ambiguous') { res.push(w.name + ': already there'); continue; }
      let anchor = null; for (const mt of mates) { const p = rotaFind(people, mt); if (p && p !== 'ambiguous') { anchor = p; break; } }
      if (!anchor) throw new Error(`no colleague from ${row.Team || 'the same team'} found in ${w.name} to place them next to`);
      const block = people.filter(p => p.team === anchor.team && p.c === anchor.c); const lastRow = Math.max(...block.map(p => p.row)); const nr = lastRow + 1; const rw = grid[nr] || [];
      const cells = []; for (let k = -1; k < 7; k++) cells.push(String(rw[anchor.c + k] ?? '').trim());
      if (cells.some(Boolean)) throw new Error(`no free row under the ${anchor.team} block in ${w.name}. Add them by hand`);
      await writeCell(URLS.rota, w.name, nr, anchor.c - 1, name, false); forget('rota:' + w.name); res.push(w.name + ': added'); 
    }
    out.push(['rota', res.join('; ')]); await syncLog('person', row.Email, `${name}: rota`, 'Ok', res.join('; '), by);
  } catch (e) { out.push(['rota', 'Failed: ' + e.message.slice(0, 160)]); await syncLog('person', row.Email, `${name}: rota`, 'Failed', e.message, by); }
  return out;
}
async function pushUpdateToPromotions(r, prevTitle) {
  const sheet = (await sheetNames(URLS.kb)).find(n => /promo/i.test(n) && !/posh/i.test(n)); if (!sheet) throw new Error('no Promotions sheet in the workbook');
  const ws = await wsOf(URLS.kb, sheet); const ur = await g('GET', `${ws}/usedRange(valuesOnly=true)?$select=address,values`);
  const m = ur.address.match(/!([A-Z]+)(\d+)/); let c0 = 0; for (const ch of m[1]) c0 = c0 * 26 + ch.charCodeAt(0) - 64; c0--; const r0 = +m[2] - 1; const v = ur.values;
  let nr = -1, lc = -1;
  for (let i = 0; i < Math.min(v.length, 10) && nr < 0; i++) for (let c = 0; c < Math.min((v[i] || []).length, 5); c++) if (/promotion details|product details/i.test(String(v[i][c] || ''))) { nr = i; lc = c; break; }
  if (nr < 0) throw new Error("the Promotions sheet layout wasn't recognised");
  const width = Math.max(...v.map(x => x.length));
  let col = r['Sheet Col'] !== '' && r['Sheet Col'] != null && !isNaN(+r['Sheet Col']) ? +r['Sheet Col'] : null;
  if (col != null && prevTitle && norm((v[nr] || [])[col - c0]) !== norm(prevTitle)) col = null; // the column moved or was changed by hand: add a fresh one
  if (col == null) col = c0 + width;
  let detail = -1, first = -1; for (let i = nr + 1; i < v.length; i++) { const lab = String(v[i][lc] ?? '').trim(); if (/details?|description|benefit|info|message/i.test(lab)) { detail = i; break; } if (first < 0 && lab) first = i; }
  if (detail < 0) detail = first; if (detail < 0) throw new Error('no field rows in the Promotions sheet');
  await writeCell(URLS.kb, sheet, r0 + nr, col, r.Title, false); await writeCell(URLS.kb, sheet, r0 + detail, col, r.Body, false); forget('kb');
  return col;
}
async function syncUpdate(t, r, prevTitle, by) {
  let status;
  if (isTestMode(await loadSettings())) { r['Sheet Col'] = 'Test'; return; }
  try { r['Sheet Col'] = await pushUpdateToPromotions(r, prevTitle); await syncLog('update', r.ID, r.Title, 'Ok', 'Written to Promotions', by); }
  catch (e) { await syncLog('update', r.ID, r.Title, 'Failed', e.message, by); }
}
async function readSyncLog(u) {
  if (!(canEditRota(u) || ['Duty Manager', 'General Manager'].includes(u.role))) deny();
  const t = await cached('synclog', 20e3, async () => { try { await ensureTable(URLS.portal, 'Sync Log', 'tblSyncLog', SYNC_HEADERS); return await table(URLS.portal, 'tblSyncLog'); } catch (e) { return { headers: SYNC_HEADERS, rows: [] }; } });
  const latest = new Map(); for (const r of t.rows.filter(x => x.ID)) latest.set(r.Kind + '|' + r.Ref, { ...r, At: tsString(r.At) });
  const rows = [...latest.values()].reverse().slice(0, 40); return { rows, failed: rows.filter(r => r.Status === 'Failed').length };
}
async function doSyncRetry(u, acc, b) {
  if (!(canEditRota(u) || ['Duty Manager', 'General Manager'].includes(u.role))) deny();
  const kind = String(b.kind || ''), ref = String(b.ref || '');
  if (kind === 'swap') { const t = await loadSwapsRaw(); const sw = t.rows.find(x => x.ID === ref); if (!sw || sw.Status !== 'Approved') bad('Swap not found.'); if (sw['Written To Rota'] === 'Yes') bad('This swap is already in the rota.'); return { status: await syncSwap(t, sw, acc, u.email) }; }
  if (kind === 'rotaedit') { const t = await loadRotaEditsRaw(); const r = t.rows.find(x => x.ID === ref); if (!r) bad('Change not found.'); if (r.Synced === 'Yes') bad('Already in the rota.'); const who = acc.rows.find(x => x.Email === r.Email);
    try { const w = await rotaWriteEdit(who, r.Date, r.Shift); r.Original = r.Original || w.was; r.Synced = 'Yes'; await syncLog('rotaedit', r.ID, `${r.Name} ${r.Date}`, 'Ok', 'Written to the rota', u.email); } catch (e) { await syncLog('rotaedit', r.ID, `${r.Name} ${r.Date}`, 'Failed', e.message, u.email); bad('Still failing: ' + e.message); }
    await setRow(URLS.portal, 'tblRotaEdits', r._i, t.headers.map(k => r[k] ?? '')); forget('rotaedits'); return { status: 'Yes' }; }
  if (kind === 'person') { const who = acc.rows.find(x => x.Email === ref); if (!who) bad('Person not found.'); return { results: await syncNewPerson(who, u.email) }; }
  if (kind === 'update') { const t = await loadUpdates(); const r = t.rows.find(x => x.ID === ref); if (!r) bad('Update not found.'); await syncUpdate(t, r, '', u.email); await setRow(URLS.portal, 'tblUpdates', r._i, t.headers.map(k => r[k] ?? '')); return { status: r['Sheet Col'] === '' ? 'Failed' : 'Yes' }; }
  bad('Unknown item.');
}
const REDIT_HEADERS = ['ID', 'Email', 'Name', 'Date', 'Shift', 'Original', 'Note', 'Edited By', 'Edited At', 'Synced'];
async function loadRotaEditsRaw() {
  await ensureTable(URLS.portal, 'Rota Edits', 'tblRotaEdits', REDIT_HEADERS);
  let t = await table(URLS.portal, 'tblRotaEdits');
  if (!t.headers.includes('Synced')) { const bk = await book(URLS.portal); await g('POST', `${bk}/tables/tblRotaEdits/columns/add`, { name: 'Synced' }); t = await table(URLS.portal, 'tblRotaEdits'); }
  t.rows = t.rows.filter(r => r.ID && r.Email); t.rows.forEach(r => { r.Email = normEmail(r.Email); r.Date = toKey(r.Date); r['Edited At'] = tsString(r['Edited At']); });
  return t;
}
const loadRotaEdits = () => cached('rotaedits', 15000.0, loadRotaEditsRaw);
async function doRotaEdit(u, acc, b) {
  if (!canEditRota(u)) deny('You are not set up to change the schedule. Ask a portal admin.');
  const who = acc.rows.find(r => r.Email === normEmail(b.email)); if (!who || isNo(who.Active)) bad('Person not found.');
  if (!isKey(b.date)) bad('Pick a date.');
  const shift = String(b.shift || '').trim().slice(0, 120); if (!shift) bad('Enter a shift (for example 09:00-18:00) or choose Day off.');
  const t = await loadRotaEditsRaw();
  const old = t.rows.find(r => r.Email === who.Email && r.Date === b.date);
  const row = { ID: old ? old.ID : uid('R'), Email: who.Email, Name: who['Full Name'], Date: b.date, Shift: shift, Original: old ? old.Original : String(b.original || '').slice(0, 120), Note: String(b.note || '').slice(0, 300), 'Edited By': u.email, 'Edited At': new Date().toISOString(), Synced: '' };
  let rota = 'test';
  if (!isTestMode(await loadSettings())) {
    try { const w = await rotaWriteEdit(who, b.date, shift); if (!old) row.Original = w.was; row.Synced = 'Yes'; rota = 'written'; await portalChange([[row['Edited At'], 'Schedule change', who['Full Name'], b.date, w.was, shift, u.email, row.ID]]); await syncLog('rotaedit', row.ID, `${who['Full Name']} ${b.date}`, 'Ok', 'Written to the rota', u.email); }
    catch (e) { row.Synced = 'Failed: ' + e.message.slice(0, 120); rota = 'failed'; await syncLog('rotaedit', row.ID, `${who['Full Name']} ${b.date}`, 'Failed', e.message, u.email); }
  } else row.Synced = 'Test';
  if (old) await setRow(URLS.portal, 'tblRotaEdits', old._i, rowFor(t.headers, row)); else await addRow(URLS.portal, 'tblRotaEdits', rowFor(t.headers, row));
  forget('rotaedits'); return { ok: true, rota, rotaDetail: rota === 'failed' ? row.Synced : undefined };
}
async function doRotaEditClear(u, b) {
  if (!canEditRota(u)) deny('You are not set up to change the schedule. Ask a portal admin.');
  const t = await loadRotaEditsRaw(); const r = t.rows.find(x => x.Email === normEmail(b.email) && x.Date === b.date); if (!r) return { ok: true };
  if (r.Synced === 'Yes' && r.Original !== '' && r.Original != null) {
    const who = (await loadAccessRaw()).rows.find(x => x.Email === r.Email);
    try { await rotaWriteEdit(who, r.Date, r.Original); await portalChange([[new Date().toISOString(), 'Restored rota', r.Name, r.Date, r.Shift, String(r.Original), u.email, r.ID]]); }
    catch (e) { await syncLog('rotaedit', r.ID, `${r.Name} ${r.Date}`, 'Failed', 'Restore failed: ' + e.message, u.email); bad("The Excel rota couldn't be restored (" + e.message + '). Fix the cell by hand, then try again.'); }
  }
  const bk = await book(URLS.portal); await g('DELETE', `${bk}/tables/tblRotaEdits/rows/itemAt(index=${r._i})`);
  forget('rotaedits'); return { ok: true };
}
async function doAccess(u, accCached, b) {
  if (!u.admin) deny('Only portal admins can edit the access list.');
  let acc = await loadAccessRaw();
  const o = b.row || {};
  if (!String(o['Full Name'] || '').trim()) bad('Name is required.');
  for (const c of ['Can Approve Swaps', 'Can Edit Rota']) if (!acc.headers.includes(c) && c in o) { await ensureAccessColumn(acc, c); acc = await loadAccessRaw(); }
  if (o['Portal Role'] && !ROLES.includes(o['Portal Role'])) bad('Unknown role.');
  const old = Number.isInteger(o._i) && o._i >= 0 ? acc.rows.find(x => x._i === o._i) : null;
  const renamed = old && norm(old['Full Name']) !== norm(o['Full Name']);
  if (renamed && !acc.headers.includes('Previous Name')) { await ensureAccessColumn(acc, 'Previous Name'); acc = await loadAccessRaw(); }
  const H = acc.headers;
  const tl = acc.rows.find(x => nameKeys(x).includes(norm(o['Team Leader']))) || findLeader(o['Team Leader'], acc.rows); const clean = { ...o, Email: normEmail(o.Email), 'Manager Email': tl ? tl.Email : normEmail(o['Manager Email']) };
  if (tl && norm(o['Team Leader']) !== norm(tl['Full Name'])) clean['Team Leader'] = tl['Full Name'];
  // Renaming someone: keep the old name in 'Previous Name' so their Leaves Tracker history, balance and team links still match
  if (renamed) { const prev = String(old['Previous Name'] || '').split(';').map(x => x.trim()).filter(Boolean); if (!prev.some(x => norm(x) === norm(old['Full Name']))) prev.push(String(old['Full Name']).trim()); clean['Previous Name'] = prev.filter(x => norm(x) !== norm(clean['Full Name'])).join('; '); }
  if (clean.Email && acc.rows.some(x => x.Email === clean.Email && x._i !== o._i)) bad('That email is already on the list.');
  const vals = H.map(k => clean[k] ?? '');
  if (old) { if (o._i === u.row._i && (!yes(clean['Portal Admin']) || isNo(clean.Active))) bad("You can't remove your own admin access."); await setRow(URLS.portal, T.access, o._i, vals); }
  else await addRow(URLS.portal, T.access, vals);
  let moved = 0;
  if (renamed) for (const r of acc.rows) {   // people whose team leader was written with the old name
    if (r._i === o._i || norm(r['Team Leader']) !== norm(old['Full Name'])) continue;
    r['Team Leader'] = clean['Full Name']; if (!r['Manager Email'] && clean.Email) r['Manager Email'] = clean.Email;
    await setRow(URLS.portal, T.access, r._i, H.map(k => r[k] ?? '')); moved++;
  }
  qCache = null; let sync; if (!old && clean.Email && !isNo(clean.Active)) { try { sync = await syncNewPerson(clean, u.email); } catch (e) { console.error('person sync', e.message); } }
  return { ok: true, renamed: !!renamed, teamUpdated: moved, sync };
}

/* ---------------- quality: evaluations + disputes ---------------- */
const DEFAULT_QUALITY_URL = process.env.QUALITY_URL || 'https://inhouseworlduk-my.sharepoint.com/:x:/g/personal/karim_abdelkader_gto_org_uk/IQCm4e8B9OfhTLZuFVv_YcYqAcy92G6NTopq-d0MHsDT9MI';
const DISPUTE_HEADERS = ['ID', 'Raised At', 'Eval ID', 'Month', 'Agent Email', 'Agent Name', 'Chat Date', 'Client', 'Score', 'Raised By', 'Reason', 'Status', 'Responded By', 'Responded At', 'Response', 'Email Sent', 'Manager Email', 'Reviewed By', 'Reviewed At', 'Manager Note'];
const QFILE_HEADERS = ['Month', 'File Link'];
async function tableExists(url, name) { const b = await book(url); try { await g('GET', `${b}/tables/${encodeURIComponent(name)}?$select=name`); return true; } catch (e) { return false; } }
async function ensureTable(url, sheet, name, headers, firstRows = []) {
  if (await tableExists(url, name)) return;
  const b = await book(url);
  try { await g('POST', `${b}/worksheets/add`, { name: sheet }); } catch (e) { /* sheet may exist */ }
  const last = String.fromCharCode(64 + headers.length);
  const rows = [headers, ...firstRows];
  await g('PATCH', `${b}/worksheets('${encodeURIComponent(sheet)}')/range(address='A1:${last}${rows.length}')`, { values: rows });
  const t = await g('POST', `${b}/tables/add`, { address: `'${sheet}'!A1:${last}${rows.length}`, hasHeaders: true });
  await g('PATCH', `${b}/tables/${encodeURIComponent(t.id || t.name)}`, { name, style: 'TableStyleMedium2' });
}
async function ensureAccessColumn(acc, col) {
  if (acc.headers.includes(col)) return;
  const b = await book(URLS.portal);
  await g('POST', `${b}/tables/${T.access}/columns/add`, { name: col });
  forget('access');
}
async function qualityFiles() {
  await ensureTable(URLS.portal, 'Quality Files', 'tblQualityFiles', QFILE_HEADERS, [['Oct 2026', DEFAULT_QUALITY_URL]]);
  const t = await table(URLS.portal, 'tblQualityFiles');
  const label = m => { const d = typeof m === 'number' ? serialToDate(m) : null; return d ? d.toLocaleDateString('en-GB', { month: 'short', year: 'numeric' }) : (String(m || '').trim() || 'Month'); };
  return t.rows.filter(r => String(r['File Link'] || '').startsWith('https://')).map(r => ({ month: label(r.Month), url: String(r['File Link']).trim() }));
}
const pct = v => { const n = typeof v === 'number' ? v : parseFloat(String(v).replace('%', '')) / (String(v).includes('%') ? 100 : 1); return isNaN(n) ? null : n; };
function lev1(a, b) { if (Math.abs(a.length - b.length) > 1) return false; let i = 0, j = 0, e = 0; while (i < a.length && j < b.length) { if (a[i] === b[j]) { i++; j++; continue; } if (++e > 1) return false; if (a.length > b.length) i++; else if (b.length > a.length) j++; else { i++; j++; } } return e + (a.length - i) + (b.length - j) <= 1; }
function matchPerson(name, acc) {
  const n = norm(name); if (!n) return null;
  const exact = acc.rows.find(r => norm(r['Roster Name']) === n || norm(r['Full Name']) === n || (r['Quality Name'] && norm(r['Quality Name']) === n)); if (exact) return exact;
  const toks = n.split(' '); let best = null, bs = 0, tie = false;
  for (const r of acc.rows) {
    if (MANAGER_ROLES.includes(r['Portal Role'])) continue; // only an exact name links an evaluation to a team leader or manager
    const full = norm(r['Full Name']).split(' '); if (full[0] !== toks[0] && !(toks[0].length >= 5 && lev1(toks[0], full[0]))) continue;
    if (!toks.slice(1).every(t => full.slice(1).some(f => f === t || (t.length > 3 && f.length > 3 && (f.startsWith(t) || t.startsWith(f) || lev1(t, f)))))) continue;
    const sc = toks.length + (full[0] === toks[0] ? 0.5 : 0) - 0.1 * Math.max(0, full.length - toks.length) + (r['Portal Role'] === 'Agent' ? 0.05 : 0); if (sc > bs) { bs = sc; best = r; tie = false; } else if (sc === bs) tie = true;
  }
  return tie ? null : best;
}
let qCache = null;
async function loadEvaluations(acc) {
  if (qCache && Date.now() - qCache.at < 120e3) return qCache.items;
  const files = await qualityFiles(); const items = [];
  for (const [fi, f] of files.entries()) {
    const names = (await sheetNames(f.url)).filter(n => /evaluation/i.test(n));
    for (const sheet of names) {
      const v = await sheetValues(f.url, sheet); const hr = v.findIndex(r => r.some(c => /agent'?s name/i.test(String(c)))); if (hr < 0) continue;
      const H = v[hr].map(x => String(x || '').trim());
      const idx = n => H.findIndex(h => h.toLowerCase().startsWith(n.toLowerCase()));
      const all = (n) => H.map((h, i) => [h, i]).filter(([h]) => h.toLowerCase() === n.toLowerCase()).map(([, i]) => i);
      const cAgent = idx("agent's name") >= 0 ? idx("agent's name") : idx('agent'), cTL = idx('tl name'), cWeek = idx('week'), cDate = idx('chat date'), cLink = idx('thread link'), cClient = idx("client"), cScore = H.lastIndexOf(H.find(h => /^score$/i.test(h)) || 'Score');
      const metrics = ['RT', 'CRT', 'HRT', 'ETA', 'ICA', 'RN', 'IQ'].map(k => [k, H.findIndex(h => h.toUpperCase().startsWith(k + ' ('))]).filter(([, i]) => i >= 0);
      const errCats = [['CTC', 'Critical-to-Customer'], ['CTB', 'Critical-to-Business'], ['CTQ', 'Critical-to-Quality'], ['NCE', 'Non-Critical Error'], ['ZTP', 'ZTP Violation']];
      const typeCols = all('error type'), incCol = idx('ztp incident');
      v.slice(hr + 1).forEach((r, k) => {
        const agent = String(r[cAgent] || '').trim(); const score = pct(r[cScore]); if (!agent || score == null) return;
        const p = matchPerson(agent, acc); const d = parseDateCell(r[cDate]);
        const errors = [];
        errCats.forEach(([code, head], j) => { const ci = H.findIndex(h => h.toLowerCase() === head.toLowerCase()); if (ci >= 0 && /^yes$/i.test(String(r[ci]).trim())) { const type = code === 'ZTP' ? r[incCol] : r[typeCols[j]]; errors.push({ code, type: String(type || '').trim() }); } });
        items.push({ id: `${fi}|${sheet}|${hr + 2 + k}`, month: f.month, sheet, week: r[cWeek] ?? '', agent, agentEmail: p ? p.Email : '', agentFull: p ? p['Full Name'] : agent, tl: (p && p._tl && p['Manager Email']) ? String(p._tl).trim().split(/\s+/)[0] : String(r[cTL] || '').trim(), tlEmail: p ? p['Manager Email'] : '',
          date: d ? dkey(d) : '', link: /^https?:/i.test(String(r[cLink] || '')) ? String(r[cLink]) : '', client: String(r[cClient] || '').trim(), score,
          metrics: Object.fromEntries(metrics.map(([m, i]) => [m, pct(r[i])])), resolved: String(r[idx('issue resolution')] ?? ''), irresolution: String(r[idx('irresolution')] ?? ''),
          errors, inquiry: String(r[idx('inquiry type')] ?? ''), findings: String(r[idx('findings')] ?? ''), actions: String(r[idx('actions')] ?? '') });
      });
    }
  }
  qCache = { at: Date.now(), items }; return items;
}
async function loadDisputes() {
  await ensureTable(URLS.portal, 'Disputes', 'tblDisputes', DISPUTE_HEADERS);
  let t = await table(URLS.portal, 'tblDisputes');
  const missing = DISPUTE_HEADERS.filter(h => !t.headers.includes(h));
  if (missing.length) { const b = await book(URLS.portal); for (const name of missing) await g('POST', `${b}/tables/tblDisputes/columns/add`, { name }); t = await table(URLS.portal, 'tblDisputes'); }
  t.rows = t.rows.filter(r => r.ID);
  t.rows.forEach(r => { r['Agent Email'] = normEmail(r['Agent Email']); r['Raised By'] = normEmail(r['Raised By']); r['Manager Email'] = normEmail(r['Manager Email']); r['Raised At'] = tsString(r['Raised At']); r['Responded At'] = tsString(r['Responded At']); r['Reviewed At'] = tsString(r['Reviewed At']); });
  return t;
}
const isQuality = u => u.role === 'Quality' || yes(u.row['Quality Team']);
function canSeeEval(u, e, team) {
  if (seesAll(u) || isQuality(u)) return true;
  if (e.agentEmail === u.email) return true;
  if (u.role === 'Team Leader') return team.has(e.agentEmail) || (!e.agentEmail && norm(e.tl) && norm(u.name).split(' ')[0] === norm(e.tl).split(' ')[0]);
  return false;
}
async function readQuality(u, acc) {
  await ensureAccessColumn(acc, 'Quality Team');
  const team = teamEmails(u, acc); const [ev, dt] = await Promise.all([loadEvaluations(acc), loadDisputes()]);
  const evals = ev.filter(e => canSeeEval(u, e, team));
  const ids = new Set(evals.map(e => e.id));
  const disputes = dt.rows.filter(d => {
    if (d.Status === 'With Manager') return seesAll(u) || d['Manager Email'] === u.email || d['Agent Email'] === u.email || d['Raised By'] === u.email;
    return ids.has(d['Eval ID']) || d['Raised By'] === u.email || isQuality(u) || seesAll(u);
  });
  const isManagerRole = ['Team Leader', 'Duty Manager', 'General Manager'].includes(u.role);
  return { evals, disputes, canDispute: u.role !== 'HR' && !(isQuality(u) && !isManagerRole && u.role !== 'Agent'), canReview: isManagerRole, canRespond: isQuality(u), quality: isQuality(u) };
}
function esch(s) { return String(s ?? '').replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c])); }
async function doDispute(u, acc, b) {
  const reason = String(b.reason || '').trim(); if (reason.length < 5) bad('Please explain the dispute.');
  const team = teamEmails(u, acc); const ev = (await loadEvaluations(acc)).find(e => e.id === b.evalId); if (!ev) bad('Evaluation not found.');
  const own = ev.agentEmail === u.email;
  const isManagerRole = ['Team Leader', 'Duty Manager', 'General Manager'].includes(u.role);
  if (!own && !isManagerRole) deny('You can only dispute your own evaluations.');
  if (!own && u.role === 'Team Leader' && !canSeeEval(u, ev, team)) deny('You can only dispute evaluations for your own team.');
  const t = await loadDisputes(); if (t.rows.some(d => d['Eval ID'] === ev.id && ['Open', 'With Manager'].includes(d.Status))) bad('This evaluation already has a dispute in progress.');
  // An agent's dispute goes to their direct manager first; a manager raising one for the team goes straight to Quality.
  const person = acc.rows.find(r => r.Email === ev.agentEmail);
  const mgrEmail = own ? normEmail((u.row['Manager Email'] || (person && person['Manager Email']) || ev.tlEmail || '')) : '';
  const row = { ID: uid('D'), 'Raised At': new Date().toISOString(), 'Eval ID': ev.id, Month: ev.month, 'Agent Email': ev.agentEmail, 'Agent Name': ev.agentFull, 'Chat Date': ev.date, Client: ev.client, Score: ev.score, 'Raised By': u.email, Reason: reason.slice(0, 1000), Status: own ? 'With Manager' : 'Open', 'Email Sent': 'Portal only', 'Manager Email': mgrEmail };
  await addRow(URLS.portal, 'tblDisputes', rowFor(t.headers, row));
  qCache = null; return { ok: true, status: row.Status };
}
async function doDisputeReview(u, b) {
  if (!['Team Leader', 'Duty Manager', 'General Manager'].includes(u.role)) deny('Only managers can review disputes.');
  const t = await loadDisputes(); const d = t.rows.find(x => x.ID === b.id); if (!d) bad('Dispute not found.');
  if (d.Status !== 'With Manager') bad('This dispute is no longer waiting for a manager.');
  if (d['Agent Email'] === u.email) deny("You can't review your own dispute.");
  if (!(u.email === d['Manager Email'] || ['Duty Manager', 'General Manager'].includes(u.role))) deny("Only the agent's direct manager (or a Duty Manager) can review this dispute.");
  Object.assign(d, { Status: b.ok ? 'Open' : 'Declined', 'Reviewed By': u.email, 'Reviewed At': new Date().toISOString(), 'Manager Note': String(b.note || '').slice(0, 1000) });
  await setRow(URLS.portal, 'tblDisputes', d._i, t.headers.map(k => d[k] ?? ''));
  qCache = null; return { ok: true, status: d.Status };
}
async function doDisputeRespond(u, acc, b) {
  if (!isQuality(u)) deny('Only the quality team can respond to disputes.');
  const t = await loadDisputes(); const d = t.rows.find(x => x.ID === b.id); if (!d) bad('Dispute not found.'); if (d.Status !== 'Open') bad('This dispute is already closed.');
  const status = b.decision === 'Accepted' ? 'Accepted' : b.decision === 'Rejected' ? 'Rejected' : null; if (!status) bad('Choose Accept or Reject.');
  Object.assign(d, { Status: status, 'Responded By': u.email, 'Responded At': new Date().toISOString(), Response: String(b.note || '').slice(0, 1000) });
  await setRow(URLS.portal, 'tblDisputes', d._i, t.headers.map(k => d[k] ?? ''));
  qCache = null; return { ok: true };
}


/* ---------------- Knowledge base (promotions & products workbook) ---------------- */
// Each sheet: column A holds field names ("PROMOTION DETAILS" = the offer's name), every other column is one offer.
const MONTHS = ['jan', 'feb', 'mar', 'apr', 'may', 'jun', 'jul', 'aug', 'sep', 'oct', 'nov', 'dec'];
const MRX = '(jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec)[a-z]*\\.?';
function datesIn(text, fallbackYear) {
  const t = String(text || '').toLowerCase().replace(/\u00a0|\u202f/g, ' '); const out = [];
  const yr = y => { if (!y) return fallbackYear || null; y = +y; return y < 100 ? 2000 + y : y; };
  const push = (d, m, y) => { y = yr(y); if (y && m >= 0 && d >= 1 && d <= 31 && y > 2015 && y < 2100) out.push(new Date(y, m, d)); };
  let m; let rx = new RegExp(`\\b(\\d{1,2})(?:st|nd|rd|th)?\\s*(?:of\\s+)?${MRX},?\\s*(\\d{4}|\\d{2}(?!\\d|:|\\s*(?:am|pm)))?`, 'g');
  while ((m = rx.exec(t))) push(+m[1], MONTHS.indexOf(m[2]), m[3]);
  rx = new RegExp(`${MRX}\\s+(\\d{1,2})(?:st|nd|rd|th)?(?!\\d)(?:,?\\s*(\\d{4}))?`, 'g');
  while ((m = rx.exec(t))) push(+m[2], MONTHS.indexOf(m[1]), m[3]);
  rx = /\b(\d{1,2})[\/.-](\d{1,2})[\/.-](\d{2,4})\b/g;
  while ((m = rx.exec(t))) { const a = +m[1], b = +m[2]; if (b > 12 && a <= 12) push(b, a - 1, m[3]); else push(a, b - 1, m[3]); }
  return out;
}
function kbValue(v, field) {
  if (typeof v === 'number') {
    if (v > 20000 && v < 80000) { const d = serialToDate(v); return d.toLocaleDateString('en-GB', { day: 'numeric', month: 'short', year: 'numeric' }); }
    if (v > 0 && v <= 1 && /discount|rate|benefit|hotels|flights|attraction|theme|quotation/i.test(field)) return Math.round(v * 1000) / 10 + '%';
    return String(v);
  }
  return String(v ?? '').replace(/\r/g, '').replace(/\u00a0|\u202f/g, ' ').replace(/[ \t]+\n/g, '\n').replace(/\n{3,}/g, '\n\n').trim();
}
const KB_COVER = [['Refundable hotels', /^refundable hotels/i], ['Non-refundable hotels', /^nr hotels/i], ['International flights', /^international flights/i], ['Domestic flights', /^domestic flights/i], ['Attractions & transfers', /^attraction/i], ['Theme parks & car rental', /^theme parks/i]];
function kbCover(v) { const s = String(v || '').trim().toLowerCase(); if (!s || /^[-_—.]+$/.test(s)) return null; if (/^(x|no|n|✗|×|none|not)\b/.test(s)) return false; return true; }
function kbStatus(item, neighbourYear) {
  const all = Object.values(item.fields).join(' ').toLowerCase(); const offer = item.fields['Offer Period'] || '';
  const nameYear = (String(item.name).match(/20(2\d)/) || String(item.name).match(/\/(2\d)\b/) || [])[1];
  const fbYear = nameYear ? 2000 + +nameYear : neighbourYear || null;
  if (/paused/i.test(offer)) return { status: 'Paused', note: offer };
  // The "Sold out" column of the sheet says whether the offer still takes deposits and subscriptions.
  const so = Object.entries(item.fields).find(([k]) => /sold[\s-]*out/i.test(k));
  if (so) { const v = String(so[1] ?? '').trim(); if (v && !/^(n|no|false|0|-|n\/a|na|open|available)$/i.test(v) && (/^(y|yes|true|1|x|✓|✔)$/i.test(v) || /sold|closed|full|finished|no longer|not (open|accepting|available)/i.test(v)) && !/^(not sold|no\b)/i.test(v)) return { status: 'Ended', soldOut: true, note: 'Sold out' }; }
  // Closed offers: the sheet says so in plain words (for example "no longer accepting deposits").
  const everything = String(item.name) + ' ' + Object.values(item.fields).join(' ');
  if (/no longer (accept|available|valid|on sale)|not (currently )?accepting|(sale|offer|promotion) (has |is )?(now )?(ended|closed|over)|closed for (new )?(deposits?|bookings?)|deposits? (are |is )?closed/i.test(everything)) return { status: 'Ended', note: 'Closed in the sheet' };
  // A flash sale, or an offer that runs for a number of hours, starts on the date written for it (the activation date) and lasts that long.
  const hrs = (everything.match(/\b(\d{1,3})\s*-?\s*hours?\b/i) || [])[1];
  if (/flash/i.test(everything) || hrs) {
    const fd = datesIn(String(item.name) + ' ' + offer, fbYear);
    if (fd.length) { const start = new Date(Math.min(...fd)); const end = new Date(start.getTime() + (hrs ? +hrs : 24) * 3600e3); end.setHours(0, 0, 0, 0); const t0 = new Date(); t0.setHours(0, 0, 0, 0); return { status: end >= t0 ? 'On sale' : 'Ended', end: dkey(end) }; }
  }
  if (/quota|while stock|sold out|until it.?s gone/i.test(offer) && !/\b(until|ends?|by|before)\b[^.\n]{0,70}\d/i.test(offer)) return { status: 'While quota lasts' };
  const ds = datesIn(offer, fbYear); const today = new Date(); today.setHours(0, 0, 0, 0);
  if (ds.length) { const end = new Date(Math.max(...ds)); return { status: end >= today ? 'On sale' : 'Ended', end: dkey(end) }; }
  if (/all year|every (year|nov|month)|always|ongoing|no end/i.test(offer + ' ' + (item.fields['Buying Availability'] || ''))) return { status: 'Always available' };
  return { status: '' };
}
function parseKbSheet(name, v) {
  let lc = -1, nr = -1;
  for (let r = 0; r < Math.min(v.length, 10) && nr < 0; r++) for (let c = 0; c < Math.min((v[r] || []).length, 5); c++) if (/promotion details|product details/i.test(String(v[r][c] || ''))) { nr = r; lc = c; break; }
  if (nr < 0) return [];
  const cat = /posh/i.test(name) ? 'Posh Saver' : /product/i.test(name) ? 'Product' : /promo/i.test(name) ? 'Promotion' : name.replace(/\s*alf order\s*/i, '').trim();
  const fields = []; for (let r = nr + 1; r < v.length; r++) { const f = String(v[r][lc] ?? '').replace(/\s+/g, ' ').trim(); if (f) fields.push([r, f]); }
  const items = []; const width = Math.max(...v.map(r => r.length));
  for (let c = lc + 1; c < width; c++) {
    const nm = kbValue(v[nr][c], ''); if (!nm) continue;
    const it = { id: `${({ Promotion: 'PR', Product: 'PD', 'Posh Saver': 'PS' })[cat] || cat.slice(0, 2).toUpperCase()}${c}`, cat, sheet: name, order: c, name: nm.replace(/\s*\n\s*/g, ' '), fields: {} };
    fields.forEach(([r, f]) => { const val = kbValue((v[r] || [])[c], f); if (val) it.fields[f] = val; });
    it.at = { c, rows: Object.fromEntries(fields.map(([r, f]) => [f, r])) };
    it.cover = KB_COVER.map(([label, rx]) => { const f = fields.find(([, x]) => rx.test(x)); return [label, f ? kbCover(it.fields[f[1]]) : null]; });
    const tc = Object.entries(it.fields).find(([k]) => /terms/i.test(k)); const link = tc && (tc[1].match(/https?:\/\/\S+/) || [])[0]; if (link) it.link = link.replace(/[)\].,]+$/, '');
    Object.assign(it, kbStatus(it)); items.push(it);
  }
  // dates written without a year ("until 22 March"): borrow the year of the nearest neighbouring offer (columns are in date order)
  items.forEach((it, i) => { if (it.status) return; const near = [...items.slice(0, i).reverse(), ...items.slice(i + 1)].find(x => x.end); if (near) Object.assign(it, kbStatus(it, +near.end.slice(0, 4))); });
  return items;
}
// Last good copy of the workbook, kept encrypted in the portal's private Blob store (refreshed at most once a day),
// so the tab keeps working if the live file is briefly unavailable. It never goes into the code repository.
const KB_BACKUP = 'kb/last-good.txt'; let kbBackupAt = null;
async function readKbBackup() {
  const { get } = await blob(); let r; try { r = await get(KB_BACKUP, { access: 'private', useCache: false }); } catch (e) { return null; }
  if (!r) return null; const body = r.stream ? await new Response(r.stream).text() : typeof r.text === 'function' ? await r.text() : String(r);
  try { const o = dec(body); kbBackupAt = o.at || 0; return o; } catch (e) { return null; }
}
async function saveKbBackup(sheets) {
  try {
    if (kbBackupAt == null) { const o = await readKbBackup(); kbBackupAt = o ? o.at : 0; }
    if (Date.now() - kbBackupAt < 24 * 3600e3) return;
    const { put } = await blob(); const at = Date.now();
    await put(KB_BACKUP, enc({ at, sheets }), { access: 'private', addRandomSuffix: false, allowOverwrite: true, contentType: 'text/plain' }); kbBackupAt = at;
  } catch (e) { console.error('KB backup not saved', e.message); }
}
/* The offers register: one tab with one row per offer (Offer | T&Cs | notice period | activation | 1st day to book | special terms | combined | sold out).
   A ticked "sold out" box means the offer no longer takes deposits or subscriptions; an unticked one means it is on sale. */
const kbNorm = x => String(x ?? '').toLowerCase().replace(/[^a-z0-9]+/g, '');
const tick = v => v === true || /^(true|yes|y|1|x|✓|✔)$/i.test(String(v ?? '').trim());
function parseRegister(sheets) {
  const rows = [];
  for (const [sheet, v] of Object.entries(sheets)) {
    let hr = -1, nc = -1, sc = -1;
    for (let r = 0; r < Math.min((v || []).length, 6) && hr < 0; r++) { const row = v[r] || []; const n = row.findIndex(c => /^\s*offer\s*$/i.test(String(c ?? ''))); const so = row.findIndex(c => /sold[\s-]*o/i.test(String(c ?? ''))); if (n >= 0 && so >= 0) { hr = r; nc = n; sc = so; } }
    if (hr < 0) continue;
    const head = v[hr].map(h => String(h ?? '').replace(/\s+/g, ' ').trim());
    const label = h => /notice/i.test(h) ? 'Notice period' : /^activation/i.test(h) ? 'Activation' : /1st day|first day/i.test(h) ? '1st day to book' : /spe+i?c?ial|special/i.test(h) ? 'Special terms' : /combined/i.test(h) ? 'Combined with other offers' : /sold/i.test(h) ? 'Sold out' : '';
    for (let r = hr + 1; r < v.length; r++) {
      const row = v[r] || []; const name = String(row[nc] ?? '').replace(/\s+/g, ' ').trim(); if (!name) continue;
      const fields = {}; let sold = false;
      head.forEach((h, c) => { const l = label(h); if (!l || c === nc) return; const raw = row[c]; if (l === 'Sold out') { sold = tick(raw); fields[l] = sold ? 'Yes' : 'No'; return; } if (l === 'Combined with other offers') { fields[l] = tick(raw) ? 'Yes' : 'No'; return; } const val = kbValue(raw, l); if (val) fields[l] = val; });
      rows.push({ name, sold, fields, sheet, row: r });
    }
  }
  return rows;
}
function applyRegister(items, rows) {
  if (!rows.length) return { used: 0, unmatched: [] };
  const matches = new Map(); rows.forEach(rw => matches.set(rw, []));
  const claimed = new Set();
  const lookup = new Map(); items.forEach(it => { const k = kbNorm(it.name); (lookup.get(k) || lookup.set(k, []).get(k)).push(it); });
  for (const rw of rows) {
    const k = kbNorm(rw.name); let hit = lookup.get(k) || [];
    if (!hit.length && k.length >= 8) hit = items.filter(it => { const n = kbNorm(it.name); return n.length >= 8 && (n.includes(k) || k.includes(n)); });
    hit.forEach(it => { if (!claimed.has(it)) { claimed.add(it); matches.get(rw).push(it); } });
  }
  const seq = { n: 0 };
  for (const rw of rows.slice().reverse()) {
    let list = matches.get(rw);
    if (!list.length) { const it = { id: 'RG' + (++seq.n) + '-' + rw.row, cat: 'Product', sheet: rw.sheet, order: 100000 + rw.row, name: rw.name, fields: {}, cover: KB_COVER.map(([l]) => [l, null]) }; items.push(it); list = [it]; claimed.add(it); }
    for (const it of list) {
      Object.assign(it.fields, rw.fields); it.reg = true; delete it.note;
      const expired = /expired|no longer/i.test(rw.fields['Special terms'] || '');
      if (rw.sold || expired) { it.status = 'Ended'; it.soldOut = rw.sold; it.note = rw.sold ? 'Sold out' : 'Expired'; delete it.end; }
      else { it.status = 'On sale'; delete it.end; delete it.soldOut; }
    }
  }
  // Anything still marked on sale that is not in the register is not current.
  for (const it of items) if (!it.reg && (it.status === 'On sale' || it.status === 'While quota lasts')) { it.status = 'Ended'; it.note = 'Not in the current sheet'; delete it.end; }
  return { used: claimed.size, unmatched: [] };
}
async function loadKb() {
  return cached('kb', 600e3, async () => {
    let source = 'live', sheets = {}, live = {}, error = '';
    try { const names = await sheetNames(URLS.kb); for (const n of names) live[n] = await sheetValues(URLS.kb, n); sheets = live; if (!Object.entries(sheets).some(([n, v]) => parseKbSheet(n, v || []).length)) throw new Error('no offers found in the sheet layout'); await saveKbBackup(sheets); }
    catch (e) {
      error = e.message; const b = await readKbBackup(); const reg0 = parseRegister(live);
      if (!b && !reg0.length) throw new HttpError(503, "The knowledge base workbook can't be opened by the portal (" + e.message.slice(0, 120) + '). Share the file with opsportal@gto.org.uk.');
      if (b) { sheets = b.sheets; source = 'backup ' + new Date(b.at).toLocaleDateString('en-GB', { day: 'numeric', month: 'short', year: 'numeric' }); } else { sheets = {}; }
    }
    const items = Object.entries(sheets).flatMap(([n, v]) => parseKbSheet(n, v || []));
    const reg = parseRegister(live); applyRegister(items, reg);
    const cats = [...new Set(items.map(i => i.cat))]; items.forEach(i => { const same = items.filter(x => x.cat === i.cat); i.rank = same.length - same.indexOf(i); });
    return { items, cats, source, register: reg.length > 0, soldUnmatched: [], error: error.slice(0, 200), loadedAt: new Date().toISOString() };
  });
}
/* Portal-side edits of the workbook's offers: the workbook itself is never changed. */
const KBEDIT_HEADERS = ['Key', 'Hidden', 'Fields', 'Edited By', 'Edited At'];
async function loadKbEdits() {
  try { await ensureTable(URLS.portal, 'KB Edits', 'tblKbEdits', KBEDIT_HEADERS); const t = await table(URLS.portal, 'tblKbEdits'); t.rows = t.rows.filter(r => r.Key); return t; }
  catch (e) { console.error('kb edits', e.message); return { headers: KBEDIT_HEADERS, rows: [], missing: true }; }
}
function deriveItem(it, base) {
  it.cover = KB_COVER.map(([label, rx]) => { const f = Object.keys(it.fields).find(x => rx.test(x)); return [label, f ? kbCover(it.fields[f]) : null]; });
  const tc = Object.entries(it.fields).find(([k]) => /terms/i.test(k)); const link = tc && (tc[1].match(/https?:\/\/\S+/) || [])[0]; it.link = link ? link.replace(/[)\].,]+$/, '') : undefined;
  delete it.soldOut; delete it.note;
  if (base.reg) { it.status = base.status; it.end = base.end; if (base.soldOut) it.soldOut = base.soldOut; if (base.note) it.note = base.note; return; }
  const st = kbStatus(it); if (st && st.status) Object.assign(it, st); else { it.status = base.status; it.end = base.end; }
}
async function loadKbView() {
  const [kb, ed] = await Promise.all([loadKb(), cached('kbedits', 30e3, loadKbEdits)]);
  const by = new Map(ed.rows.map(r => [r.Key, r])); const items = [], hidden = [];
  for (const base of kb.items) {
    const key = base.cat + '|' + base.name; const e = by.get(key);
    if (!e) { items.push({ ...base, key }); continue; }
    if (yes(e.Hidden)) { hidden.push({ key, name: base.name, cat: base.cat }); continue; }
    let over = {}; try { over = JSON.parse(e.Fields || '{}'); } catch (_) { /* ignore */ }
    const closed = over.__closed === 'Yes'; delete over.__closed;
    const it = { ...base, key, fields: { ...base.fields, ...over }, edited: Object.keys(over).length > 0 }; if (it.edited) deriveItem(it, base);
    if (closed) { it.status = 'Ended'; it.closed = true; it.note = 'Marked as no longer on sale in the portal'; }
    items.push(it);
  }
  return { ...kb, items, hidden, editsSig: ed.rows.map(r => r.Key + r['Edited At']).join(',') };
}
const canEditKb = u => u.admin || ['Duty Manager', 'General Manager'].includes(u.role);
async function doKbEdit(u, b) {
  if (!canEditKb(u)) deny('Only Duty Managers, the GM and admins can edit or hide topics.');
  const key = String(b.key || ''); const kb = await loadKb(); const base = kb.items.find(i => i.cat + '|' + i.name === key); if (!base) bad('Topic not found.');
  const t = await loadKbEdits(); if (t.missing) bad("The data file can't hold topic edits yet.");
  let r = t.rows.find(x => x.Key === key); const isNew = !r; if (isNew) r = { Key: key, Hidden: '', Fields: '', 'Edited By': '', 'Edited At': '' };
  if (b.reset) { r.Hidden = ''; r.Fields = ''; }
  if (b.hide != null) r.Hidden = b.hide ? 'Yes' : '';
  if (b.closed != null) { let cur = {}; try { cur = JSON.parse(r.Fields || '{}'); } catch (_) { /* ignore */ } if (b.closed) cur.__closed = 'Yes'; else delete cur.__closed; const js = JSON.stringify(cur); r.Fields = js === '{}' ? '' : js; }
  if (b.fields && typeof b.fields === 'object') {
    let cur = {}; try { cur = JSON.parse(r.Fields || '{}'); } catch (_) { /* ignore */ }
    for (const [k, v] of Object.entries(b.fields)) { if (!(k in base.fields)) bad('Unknown field: ' + k); const val = String(v ?? '').slice(0, 8000); if (val === base.fields[k]) delete cur[k]; else cur[k] = val; }
    const js = JSON.stringify(cur); if (js.length > 30000) bad('That is too much text for one topic.'); r.Fields = js === '{}' ? '' : js;
  }
  r['Edited By'] = u.email; r['Edited At'] = new Date().toISOString();
  if (isNew) await addRow(URLS.portal, 'tblKbEdits', rowFor(t.headers, r)); else await setRow(URLS.portal, 'tblKbEdits', r._i, t.headers.map(k => r[k] ?? ''));
  forget('kbedits'); ixCache = null;
  // A correction made in the portal is also written into the offers workbook, so the sheet stays the single source of truth.
  let sheetNote = '';
  if (b.fields && typeof b.fields === 'object' && base.at && kb.source === 'live') {
    const done = [], failed = [];
    for (const [k, v] of Object.entries(b.fields)) {
      const row = base.at.rows[k]; if (row == null || String(v ?? '') === base.fields[k]) continue;
      try { await writeCell(URLS.kb, base.sheet, row, base.at.c, String(v ?? '').slice(0, 8000), true); done.push(k); }
      catch (e) { failed.push(k); await syncLog('offer', key.slice(0, 60), 'Offer "' + base.name.slice(0, 60) + '": ' + k, 'Failed', e.message, u.email); }
    }
    if (done.length) { forget('kb'); await syncLog('offer', key.slice(0, 60), 'Offer "' + base.name.slice(0, 60) + '" updated in the sheet: ' + done.join(', '), 'Ok', 'Written to the offers workbook', u.email); }
    sheetNote = failed.length ? 'failed' : done.length ? 'written' : '';
  }
  return { ok: true, sheet: sheetNote };
}
const UPDATE_HEADERS = ['ID', 'Posted At', 'Posted By', 'Posted By Name', 'Title', 'Body', 'Category', 'Pinned', 'Active', 'Edited At', 'Edited By', 'Sheet Col'];
async function loadUpdates() {
  await ensureTable(URLS.portal, 'Updates', 'tblUpdates', UPDATE_HEADERS);
  let t = await table(URLS.portal, 'tblUpdates');
  const missing = UPDATE_HEADERS.filter(h => !t.headers.includes(h));
  if (missing.length) { const b = await book(URLS.portal); for (const name of missing) await g('POST', `${b}/tables/tblUpdates/columns/add`, { name }); t = await table(URLS.portal, 'tblUpdates'); }
  t.rows = t.rows.filter(r => r.ID); t.rows.forEach(r => { r['Posted At'] = tsString(r['Posted At']); r['Edited At'] = tsString(r['Edited At']); });
  return t;
}
const canPost = u => u.admin || ['Team Leader', 'Duty Manager', 'General Manager', 'HR', 'Quality'].includes(u.role);
/* New offers added to the Promotions workbook by hand become an update, so everyone is notified (the normal update poll picks it up). */
const KB_SEEN = 'kb/seen.txt'; let lastOfferCheck = 0;
async function detectNewOffers() {
  if (Date.now() - lastOfferCheck < 5 * 60e3) return; lastOfferCheck = Date.now();
  try {
    const st = await loadSettings(); if (isTestMode(st)) return;
    const kb = await loadKb(); if (kb.source !== 'live' || !kb.items.length) return;
    const raw = await blobText(KB_SEEN); let seen = null; try { seen = raw ? new Set(dec(raw).keys) : null; } catch (e) { seen = null; }
    const keys = kb.items.map(i => i.cat + '|' + i.name);
    const save = async set => { const { put } = await blob(); await put(KB_SEEN, enc({ keys: [...set] }), { access: 'private', addRandomSuffix: false, allowOverwrite: true, contentType: 'text/plain' }); };
    if (!seen) { await save(new Set(keys)); return; }   // first run: today's offers are the baseline, nobody is notified about them
    const fresh = kb.items.filter(i => !seen.has(i.cat + '|' + i.name)); if (!fresh.length) return;
    const up = await cached('updates', 30e3, loadUpdates); const have = new Set(up.rows.map(r => norm(r.Title)));
    const pickF = (it, rx) => { const k = Object.keys(it.fields).find(x => rx.test(x)); return k ? `${k}: ${String(it.fields[k]).replace(/\s+/g, ' ').slice(0, 220)}` : ''; };
    for (const it of fresh.slice(0, 5)) {
      seen.add(it.cat + '|' + it.name);
      if ([...have].some(h => h.includes(norm(it.name)))) continue;   // already announced through the portal
      const body = [pickF(it, /offer period|period/i), pickF(it, /discount|benefit/i), pickF(it, /booking period/i), pickF(it, /details?|description/i)].filter(Boolean).join('\n') || 'Added to the Promotions sheet. Open the offer in Updates for the details.';
      const row = { ID: 'K' + crypto.createHash('sha1').update(it.cat + '|' + it.name).digest('hex').slice(0, 10), 'Posted At': new Date().toISOString(), 'Posted By': 'promotions-sheet', 'Posted By Name': 'Promotions sheet', Title: `New ${it.cat === 'Product' ? 'product' : it.cat === 'Posh Saver' ? 'Posh Saver' : 'promotion'}: ${it.name}`.slice(0, 140), Body: body.slice(0, 1500), Category: 'Offers', Pinned: 'No', Active: 'Yes', 'Sheet Col': 'Sheet' };
      if (up.rows.some(r => r.ID === row.ID)) continue;
      await addRow(URLS.portal, 'tblUpdates', rowFor(up.headers, row));
    }
    forget('updates'); await save(seen);
  } catch (e) { console.error('new offer check', e.message); }
}
async function readUpdates() {
  await detectNewOffers();
  const up = await cached('updates', 30e3, loadUpdates);
  const st = await loadSettings();
  return { flags: { test: isTestMode(st), credits: creditsMode(st) }, updates: up.rows.filter(r => !isNo(r.Active)).sort((a, b) => (yes(b.Pinned) - yes(a.Pinned)) || String(b['Posted At']).localeCompare(String(a['Posted At']))) };
}
async function readKb(u) {
  const [kb, up] = await Promise.all([loadKbView(), cached('updates', 30e3, loadUpdates)]);
  return { ...kb, hidden: canEditKb(u) ? kb.hidden : [], canEditKb: canEditKb(u), updates: up.rows.filter(r => !isNo(r.Active)).sort((a, b) => (yes(b.Pinned) - yes(a.Pinned)) || String(b['Posted At']).localeCompare(String(a['Posted At']))), canPost: canPost(u), ai: true, aiPlus: !!AI_KEY, seesLog: ['Duty Manager', 'General Manager', 'Quality'].includes(u.role) || u.admin || yes(u.row['Quality Team']) };
}
async function doUpdate(u, b) {
  if (!canPost(u)) deny('Only team leaders and managers can post updates.');
  const t = await loadUpdates();
  if (b.id) { const r = t.rows.find(x => x.ID === b.id); if (!r) bad('Update not found.'); if (r['Posted By'] !== u.email && !['Duty Manager', 'General Manager'].includes(u.role) && !u.admin) deny('You can only change your own updates.');
    const prevTitle = r.Title; if (b.remove) r.Active = 'No'; if (b.pin != null) r.Pinned = b.pin ? 'Yes' : 'No';
    if (b.title != null || b.body != null) {
      const title = String(b.title ?? r.Title).trim().slice(0, 140), body = String(b.body ?? r.Body).trim().slice(0, 4000); if (title.length < 3) bad('Add a title.'); if (body.length < 3) bad('Write the update.');
      Object.assign(r, { Title: title, Body: body, Category: String(b.category || r.Category || 'General').slice(0, 40), 'Edited At': new Date().toISOString(), 'Edited By': u.email });
    }
    if (b.title != null || b.body != null) await syncUpdate(t, r, prevTitle, u.email);
    await setRow(URLS.portal, 'tblUpdates', r._i, t.headers.map(k => r[k] ?? '')); return { ok: true }; }
  const title = String(b.title || '').trim().slice(0, 140), body = String(b.body || '').trim().slice(0, 4000); if (title.length < 3) bad('Add a title.'); if (body.length < 3) bad('Write the update.');
  const row = { ID: uid('U'), 'Posted At': new Date().toISOString(), 'Posted By': u.email, 'Posted By Name': u.name, Title: title, Body: body, Category: String(b.category || 'General').slice(0, 40), Pinned: b.pin ? 'Yes' : 'No', Active: 'Yes' };
  row['Sheet Col'] = ''; await syncUpdate(t, row, '', u.email);
  await addRow(URLS.portal, 'tblUpdates', rowFor(t.headers, row)); return { ok: true, id: row.ID };
}
function kbText(items) {
  return items.map(i => `### ${i.name}\n[${i.cat} · ref ${i.id}${i.soldOut ? ' · SOLD OUT, no longer taking deposits or subscriptions' : i.status ? ' · ' + i.status + (i.end ? ' (offer ends ' + i.end + ')' : '') : ''}]\n` + Object.entries(i.fields).filter(([k]) => !/promotion details/i.test(k)).map(([k, v]) => `${k}: ${v}`).join('\n')).join('\n\n');
}
const askLog = {}; let ixCache = null;
const ASKLOG_HEADERS = ['ID', 'Asked At', 'Email', 'Name', 'Question', 'Answer Type'];
async function assistIndex() {
  const [kb, up] = await Promise.all([loadKbView(), cached('updates', 30e3, loadUpdates)]);
  const ups = up.rows.filter(r => !isNo(r.Active));
  const key = kb.loadedAt + '|' + kb.editsSig + '|' + ups.map(r => r.ID + (r['Edited At'] || '')).join(',');
  if (!ixCache || ixCache.key !== key) ixCache = { key, ix: buildIndex(kb.items, ups), kb, ups };
  return ixCache;
}
async function logQuestion(u, q, type) {
  try { await ensureTable(URLS.portal, 'Assistant Log', 'tblAskLog', ASKLOG_HEADERS); await addRow(URLS.portal, 'tblAskLog', [uid('Q'), new Date().toISOString(), u.email, u.name, q.slice(0, 500), type]); }
  catch (e) { console.error('ask log', e.message); }
}
/* ---------------- QA library: playbook and guides, kept in the private Blob store ---------------- */
const QA_INDEX = 'qa/index.txt', QA_FILE = id => `qa/file-${id}.txt`;
async function blobText(path) {
  const { get } = await blob(); let r; try { r = await get(path, { access: 'private', useCache: false }); } catch (e) { return null; } if (!r) return null;
  return r.stream ? await new Response(r.stream).text() : typeof r.text === 'function' ? await r.text() : String(r);
}
const loadQa = () => cached('qalib', 60e3, async () => { const t = await blobText(QA_INDEX); if (!t) return { docs: [] }; try { return dec(t); } catch (e) { return { docs: [] }; } });
async function saveQa(lib) { const { put } = await blob(); await put(QA_INDEX, enc(lib), { access: 'private', addRandomSuffix: false, allowOverwrite: true, contentType: 'text/plain' }); forget('qalib'); }
const canManageQa = u => u.admin || isQuality(u) || ['Duty Manager', 'General Manager'].includes(u.role);
async function readQaLib(u) { const lib = await loadQa(); return { canManage: canManageQa(u), docs: lib.docs.map(d => ({ id: d.id, title: d.title, file: d.file, kind: d.kind, by: d.by, at: d.at, size: d.size, sections: d.sections })) }; }
async function doQaUpload(u, b) {
  if (!canManageQa(u)) deny('Only the Quality team, Duty Managers, the GM and admins can add guides.');
  const name = String(b.name || '').replace(/[^\w .()\-–]/g, '').slice(0, 120); if (!/\.(docx|pptx)$/i.test(name)) bad('Upload a Word (.docx) or PowerPoint (.pptx) file.');
  const buf = Buffer.from(String(b.data || ''), 'base64'); if (buf.length < 500) bad('The file looks empty.'); if (buf.length > 3.2e6) bad('The file is over 3 MB. Please compress the images or split it.');
  let sections; try { sections = sectionsOf(name, buf); } catch (e) { bad("This file couldn't be read: " + e.message); }
  if (!sections.length) bad('No text was found in this file.');
  const lib = await loadQa(); const title = String(b.title || '').trim().slice(0, 120) || name.replace(/\.(docx|pptx)$/i, '').replace(/[_]+/g, ' ');
  const old = lib.docs.find(d => norm(d.title) === norm(title)); const id = old ? old.id : uid('G');
  const doc = { id, title, file: name, kind: /pptx$/i.test(name) ? 'Slides' : 'Guide', by: u.name, at: new Date().toISOString(), size: buf.length, sections };
  const { put } = await blob(); await put(QA_FILE(id), enc({ name, data: buf.toString('base64') }), { access: 'private', addRandomSuffix: false, allowOverwrite: true, contentType: 'text/plain' });
  lib.docs = old ? lib.docs.map(d => d.id === id ? doc : d) : [...lib.docs, doc]; await saveQa(lib); ixCache = null;
  return { ok: true, id, sections: sections.length, replaced: !!old };
}
async function doQaDelete(u, b) {
  if (!canManageQa(u)) deny(); const lib = await loadQa(); lib.docs = lib.docs.filter(d => d.id !== b.id); await saveQa(lib); return { ok: true };
}
async function readQaFile(u, id) {
  const lib = await loadQa(); if (!lib.docs.some(d => d.id === id)) bad('Guide not found.');
  const t = await blobText(QA_FILE(id)); if (!t) bad('The original file is not stored.'); return dec(t);
}
const QA_STOP = new Set('the and for with from that this what when how can you your are was were not any all has have had does did who why where which into out about then than them they will would should could please need needs just get got give say said ask asks client clients agent agents offer offers promo promotion'.split(' '));
const qaTok = s => [...new Set(String(s || '').toLowerCase().match(/[a-z0-9]{3,}/g) || [])].filter(t => !QA_STOP.has(t)).map(t => t.replace(/(ing|ed|es|s)$/, ''));
function qaSearch(lib, query, max = 2) {
  const q = qaTok(query); if (!q.length) return [];
  const out = [];
  for (const d of lib.docs) for (const sc of d.sections) {
    const head = qaTok(sc.heading + ' ' + d.title), body = new Set(qaTok(sc.text));
    let score = 0; for (const t of q) { if (head.includes(t)) score += 3; else if (body.has(t)) score += 1; }
    if (score >= 4) out.push({ d, sc, score });
  }
  return out.sort((a, b) => b.score - a.score).slice(0, max);
}
const qaSnippet = (sc, q, n = 420) => { const lines = sc.text.split('\n'); const qt = qaTok(q); const rank = lines.map((l, i) => [i, qaTok(l).filter(t => qt.includes(t)).length]).sort((a, b) => b[1] - a[1] || a[0] - b[0]); const pick = new Set(rank.filter(r => r[1] > 0).slice(0, 4).map(r => r[0])); if (!pick.size) lines.slice(0, 3).forEach((_, i) => pick.add(i)); let t = lines.filter((_, i) => pick.has(i)).map(l => '- ' + l.trim()).join('\n'); return t.length > n ? t.slice(0, n).replace(/\s+\S*$/, '') + '…' : t; };
async function doAsk(u, b) {
  const now = Date.now(); askLog[u.email] = (askLog[u.email] || []).filter(t => now - t < 3600e3); if (askLog[u.email].length >= 200) throw new HttpError(429, 'You have asked a lot of questions this hour. Please try again shortly.'); askLog[u.email].push(now);
  const msgs = (Array.isArray(b.messages) ? b.messages : []).slice(-10).map(m => ({ role: m.role === 'assistant' ? 'assistant' : 'user', content: String(m.content || '').slice(0, 3000) })).filter(m => m.content);
  const q = String(b.question || (msgs.length && msgs.at(-1).role === 'user' ? msgs.at(-1).content : '')).trim().slice(0, 1000);
  if (q.length < 2) bad('Ask a question.');
  const { ix, kb, ups } = await assistIndex();
  const local = localAnswer(ix, q, { ctx: b.ctx ? String(b.ctx) : null, item: b.item ? String(b.item) : null });
  // GTO Assist answers on its own; an AI key (optional) is only used for questions it isn't sure about
  if (local.conf < 0.5) await logQuestion(u, q, local.conf <= 0.1 ? 'Not found' : 'Partial');
  if (AI_KEY && local.conf < 0.6 && msgs.length) {
    try { const ai = await askClaude(u, msgs, kb, ups, await loadQa()); return { ...ai, engine: 'ai', item: local.item, follow: [] }; } catch (e) { console.error('AI fallback failed', e.message); }
  }
  let answer = local.answer, qa = [];
  try { const lib = await loadQa(); const itemName = ((kb.items.find(i => i.id === local.item) || {}).name) || ''; qa = qaSearch(lib, q + ' ' + itemName); if (qa.length) answer += '\n\n**From the QA guides**\n' + qa.map(m => `**${m.d.title} › ${m.sc.heading}**\n${qaSnippet(m.sc, q)}`).join('\n\n'); } catch (e) { console.error('qa lookup', e.message); }
  return { answer, refs: local.refs, follow: local.follow, item: local.item, conf: Math.max(local.conf, qa.length ? 0.6 : 0), engine: 'local', qa: qa.map(m => ({ doc: m.d.id, section: m.sc.heading })) };
}
async function readAskLog(u) {
  if (!(['Duty Manager', 'General Manager', 'Quality'].includes(u.role) || u.admin || yes(u.row['Quality Team']))) deny();
  if (!(await tableExists(URLS.portal, 'tblAskLog'))) return { rows: [] };
  const t = await table(URLS.portal, 'tblAskLog'); return { rows: t.rows.filter(r => r.ID).map(r => ({ ...r, 'Asked At': tsString(r['Asked At']) })).reverse().slice(0, 100) };
}
async function askClaude(u, msgs, kb, ups, qalib = { docs: [] }) {
  while (msgs.length && msgs[0].role !== 'user') msgs.shift();
  const updates = ups.slice(-30).map(r => `- ${String(r['Posted At']).slice(0, 10)} · ${r.Title}: ${r.Body}`).join('\n');
  const sys = [
    { type: 'text', text: `You are GTO Assist, the internal knowledge assistant for customer-service agents at GTO (Global Travel Organisation), a members' travel club. Answer ONLY from the knowledge base and team updates below; if the answer is not there, say so and suggest checking with the team leader or duty manager. Never invent percentages, dates, prices, links or rules. Be short: lead with the direct answer, then key conditions as bullets. Name the offer(s) used exactly as [[Offer name]]. Later entries are newer; if several versions share a name, use the newest and mention older ones. Say whether an offer is still on sale compared with today's date. Flag confidential / selected-members offers. Reply in the agent's language (English or Arabic).\n\nKNOWLEDGE BASE (${kb.items.length} entries, source: ${kb.source}):\n\n${kbText(kb.items)}`, cache_control: { type: 'ephemeral' } },
    { type: 'text', text: `Today is ${new Date().toLocaleDateString('en-GB', { weekday: 'long', day: 'numeric', month: 'long', year: 'numeric', timeZone: 'Africa/Cairo' })}. The agent asking is ${u.name} (${u.role}).\n\nRecent team updates:\n${updates || '(none)'}\n\nQA playbook and guides (follow these when advising on how to handle offers):\n${qaSearch(qalib, msgs.filter(m => m.role === 'user').slice(-2).map(m => m.content).join(' '), 5).map(m => `## ${m.d.title} › ${m.sc.heading}\n${m.sc.text.slice(0, 1500)}`).join('\n\n') || '(none matched)'}` }
  ];
  const r = await fetch('https://api.anthropic.com/v1/messages', { method: 'POST', headers: { 'content-type': 'application/json', 'x-api-key': AI_KEY, 'anthropic-version': '2023-06-01' }, body: JSON.stringify({ model: AI_MODEL, max_tokens: 1500, system: sys, messages: msgs }) });
  const j = await r.json().catch(() => ({})); if (!r.ok) throw new Error(j.error?.message || 'AI error ' + r.status);
  const text = (j.content || []).filter(c => c.type === 'text').map(c => c.text).join('\n').trim();
  const refs = [...new Set([...text.matchAll(/\[\[([^\]]+)\]\]/g)].map(m => m[1].trim()))].map(n => (kb.items.find(i => norm(i.name) === norm(n)) || {}).id).filter(Boolean);
  return { answer: text, refs, conf: 0.9 };
}


/* ---------------- Client credits: read-only lookup from the "Balance" sheet ---------------- */
const CREDITS_SHEET = 'Balance';
const sNorm = s => String(s ?? '').replace(/\u00a0|\u202f/g, ' ').toLowerCase().replace(/[^a-z0-9\u0600-\u06ff ]+/g, ' ').replace(/\s+/g, ' ').trim();
const digitsOf = s => String(s ?? '').replace(/\D+/g, '');
const serialIso = n => { const d = new Date(Math.round((n - 25569) * 864e5)); return isNaN(d) ? '' : d.toISOString().slice(0, 10); };
function parseCreditText(raw) {
  const t = String(raw || '').trim(); let title = t, issued = '', notice = '';
  const m = t.search(/[\[(]/); if (m > 0) title = t.slice(0, m).trim();
  for (const part of t.matchAll(/[\[(]([^\])]*)[\])]/g)) {
    const x = part[1].trim();
    if (/notice/i.test(x)) notice = x.replace(/\s*notice\s*$/i, '').trim();
    else if (/^\d{1,2}[\/.\-]\d{1,2}[\/.\-]\d{2,4}$/.test(x)) issued = x;
  }
  return { title: title || t, issued, notice };
}
function parseCredits(r) {
  const text = r.text || [], vals = r.values || [];
  let h = -1; for (let i = 0; i < Math.min(40, text.length); i++) if (text[i].some(c => /balance/i.test(c)) && text[i].filter(c => String(c).trim()).length >= 3) { h = i; break; }
  const head = h >= 0 ? text[h].map(c => String(c).trim()) : [];
  const find = (re, fallback) => { const i = head.findIndex(c => re.test(c)); return i >= 0 ? i : fallback; };
  const cName = find(/client|customer|name/i, 0), cNum = find(/number|phone|mobile|contact|msisdn|\bid\b/i, 1), cCred = find(/credit|offer|voucher|description|product|promo|card/i, 2), cDate = find(/date|activation|expir|valid/i, 3), cBal = find(/balance|amount|value/i, 4);
  const clients = []; let cur = null;
  for (let i = h + 1; i < text.length; i++) {
    const t = text[i].map(c => String(c ?? '').trim()); const name = t[cName], num = t[cNum], cred = t[cCred], bal = t[cBal];
    if (name) { const id = (name.match(/\(([^)]*)\)\s*$/) || [])[1] || ''; cur = { name: name.replace(/\s*\([^)]*\)\s*$/, '').trim(), id: id.trim(), number: num || '', credits: [] }; cur.search = sNorm(cur.name + ' ' + cur.id); cur.digits = digitsOf(cur.number); clients.push(cur); }
    else if (cur && num && !cur.number) { cur.number = num; cur.digits = digitsOf(num); }
    if (cur && cred) {
      const raw = vals[i] ? vals[i][cDate] : ''; const dateText = t[cDate];
      const date = typeof raw === 'number' && raw > 20000 ? serialIso(raw) : (() => { const d = new Date(String(dateText).replace(/^[A-Za-z]+,\s*/, '')); return isNaN(d) ? '' : serialIso(Date.UTC(d.getFullYear(), d.getMonth(), d.getDate()) / 864e5 + 25569); })();
      const bn = vals[i] && typeof vals[i][cBal] === 'number' ? vals[i][cBal] : parseFloat(String(bal).replace(/,/g, ''));
      cur.credits.push({ ...parseCreditText(cred), raw: cred, date, dateText, balance: isNaN(bn) ? bal : bn, balanceText: bal });
    }
  }
  return { clients, dateLabel: (head[cDate] || '').trim() || 'Date', balanceLabel: (head[cBal] || '').trim() || 'Balance', foundHeader: h >= 0 };
}
let creditsMemo = null;
export const __resetCredits = () => { creditsMemo = null; };
// when the sharing link can't be used, find the file by name among the files shared with (or searchable by) the portal account
let creditsDiag = '';
async function bookByName(re) {
  const items = [];
  try { items.push(...((await g('GET', '/me/drive/sharedWithMe?$top=200')).value || [])); } catch (e) { console.error('sharedWithMe', e.message); }
  creditsDiag = `shared files visible: ${items.length}`;
  let hit = items.find(x => re.test(x.name || ''));
  if (!hit) { try { hit = ((await g('GET', "/me/drive/search(q='Wallet System')?$top=50")).value || []).find(x => re.test(x.name || '')); } catch (e) { console.error('drive search', e.message); } }
  if (!hit) return null;
  const ri = hit.remoteItem || hit; const drive = (ri.parentReference || {}).driveId; if (!drive || !ri.id) return null;
  return `/drives/${drive}/items/${ri.id}/workbook`;
}
// workbooks Excel for the web can't open (so Graph's workbook API refuses them) are downloaded and read directly
async function readCreditsFile(b) {
  const res = await fetch('https://graph.microsoft.com/v1.0' + b.replace(/\/workbook$/, '') + '/content', { headers: { Authorization: 'Bearer ' + await opsToken() }, redirect: 'follow' });
  if (!res.ok) throw new Error(`download failed (${res.status})`);
  const out = readSheet(Buffer.from(await res.arrayBuffer()), names => names.find(n => n === CREDITS_SHEET) || names.find(n => /balance/i.test(n)) || (names.length === 1 ? names[0] : null));
  if (!out.values) throw new Error(`no tab called "${CREDITS_SHEET}"; the tabs in this file are: ${out.names.join(', ') || 'none'}`);
  return { values: out.values, text: out.text, viaFile: true };
}
async function loadCredits() {
  if (creditsMemo && Date.now() - creditsMemo.at < creditsMemo.ttl) return creditsMemo.v;
  let r, b;
  try { try { b = await book(URLS.credits); } catch (e) { b = await bookByName(/wallet system/i); if (!b) throw e; }
    try {
      const sheets = ((await g('GET', `${b}/worksheets?$select=name,visibility`)).value || []).filter(x => x.visibility !== 'Hidden').map(x => x.name);
      const sheet = sheets.find(n => n === CREDITS_SHEET) || sheets.find(n => /balance/i.test(n)) || (sheets.length === 1 ? sheets[0] : null);
      if (!sheet) throw new Error(`no tab called "${CREDITS_SHEET}"; the tabs in this file are: ${sheets.join(', ') || 'none visible'}`);
      r = await g('GET', `${b}/worksheets('${encodeURIComponent(sheet.replace(/'/g, "''"))}')/usedRange(valuesOnly=true)?$select=values,text`);
    } catch (e) { if (/no tab called/.test(e.message)) throw e; console.error('credits via Excel API failed, reading the file directly:', e.message); r = await readCreditsFile(b); }
  } catch (e) { throw new HttpError(502, `The client credits sheet couldn't be read (${e.message}${creditsDiag ? ' · ' + creditsDiag : ''}). If a tab name is listed above, tell me which one holds the balances.`); }
  const v = parseCredits(r); v.at = new Date().toISOString(); creditsMemo = { at: Date.now(), v, ttl: r.viaFile ? 300e3 : 60e3 }; return v;
}
async function readCredits(u, q) {
  if (!creditsAllowed(u, await loadSettings())) deny('Client credits is not switched on for your role.');
  const data = await loadCredits(); const qs = sNorm(q), qd = digitsOf(q);
  if (qs.length < 2 && qd.length < 4) return { clients: [], hint: 'Type at least 2 letters of the name, or 4 digits of the number.', updatedAt: data.at, dateLabel: data.dateLabel, balanceLabel: data.balanceLabel };
  const toks = qs.split(' ').filter(Boolean);
  let hit = data.clients.filter(c => (toks.length && toks.every(t => c.search.includes(t))) || (qd.length >= 4 && c.digits.includes(qd)));
  hit.sort((a, b) => (b.search.startsWith(qs) - a.search.startsWith(qs)) || a.name.localeCompare(b.name));
  const total = hit.length; hit = hit.slice(0, 15).map(({ search, digits, ...c }) => c);
  return { clients: hit, total, updatedAt: data.at, dateLabel: data.dateLabel, balanceLabel: data.balanceLabel };
}
const CREDITLOG_HEADERS = ['ID', 'Viewed At', 'Email', 'Name', 'Role', 'Client', 'Client ID'];
async function doCreditView(u, b) {
  if (!creditsAllowed(u, await loadSettings())) deny('Client credits is not switched on for your role.');
  try { await ensureTable(URLS.portal, 'Credit Lookups', 'tblCreditLog', CREDITLOG_HEADERS); await addRow(URLS.portal, 'tblCreditLog', [uid('V'), new Date().toISOString(), u.email, u.name, u.role, String(b.name || '').slice(0, 120), String(b.id || '').slice(0, 40)]); }
  catch (e) { console.error('credit log', e.message); }
  return { ok: true };
}
async function readCreditLog(u) {
  if (!isMgr(u)) deny();
  if (!(await tableExists(URLS.portal, 'tblCreditLog'))) return { rows: [] };
  const t = await table(URLS.portal, 'tblCreditLog'); return { rows: t.rows.filter(r => r.ID).map(r => ({ ...r, 'Viewed At': tsString(r['Viewed At']) })).reverse().slice(0, 100) };
}


/* ---------------- portal controls (portal admins): feature switches and test mode ---------------- */
const creditsMode = st => { const v = String(st.ClientCredits ?? 'Everyone').trim().toLowerCase(); return /^(off|no|0|disabled)/.test(v) ? 'Off' : /lead/.test(v) ? 'Leads' : 'Everyone'; };
const isTestMode = st => /^(on|yes|true|1)$/i.test(String(st.TestMode ?? '').trim());
const creditsAllowed = (u, st) => { const m = creditsMode(st); return m === 'Everyone' || (m === 'Leads' && LEADS.includes(u.role)); };
async function putSetting(key, value) {
  const t = await table(URLS.portal, T.set); const row = t.rows.find(r => String(r.Key).trim() === key);
  const vals = t.headers.map(h => h === 'Key' ? key : h === 'Value' ? value : (row ? row[h] ?? '' : ''));
  if (row) await setRow(URLS.portal, T.set, row._i, vals); else await addRow(URLS.portal, T.set, vals);
  forget('settings');
}
async function doSetting(u, b) {
  if (!u.admin) deny();
  if (b.key === 'ClientCredits') { if (!['Off', 'Leads', 'Everyone'].includes(b.value)) bad('Choose Off, Leads or Everyone.'); await putSetting('ClientCredits', b.value); }
  else if (b.key === 'TestMode') { if (b.value !== 'On') bad('Test mode is turned off with "Go live".'); await putSetting('TestStartedAt', new Date().toISOString()); await putSetting('TestMode', 'On'); }
  else bad('Unknown setting.');
  return { ok: true, settings: await loadSettings() };
}
// tables that hold work done in the portal; test mode puts rows in them, going live removes the rows created since test mode started
const RESET_TABLES = [[T.req, 'Submitted At', 'leave requests'], [T.swap, 'Submitted At', 'swap requests'], [T.att, 'Timestamp', 'attendance records'], ['tblDisputes', 'Raised At', 'disputes'], ['tblRotaEdits', 'Edited At', 'schedule changes'], ['tblUpdates', 'Posted At', 'updates'], ['tblAskLog', 'Asked At', 'assistant questions'], ['tblCreditLog', 'Viewed At', 'client lookups']];
async function doGoLive(u, b) {
  if (!u.admin) deny();
  const st = await loadSettings(); if (!isTestMode(st)) bad('Test mode is not on.');
  const start = Date.parse(st.TestStartedAt); if (isNaN(start)) bad('The time test mode started is missing, so nothing can be safely reset. Check the TestStartedAt setting.');
  const plan = [];
  for (const [name, col, label] of RESET_TABLES) {
    if (!(await tableExists(URLS.portal, name))) continue;
    const t = await table(URLS.portal, name);
    const idx = t.rows.filter(r => r.ID && Date.parse(tsString(r[col])) >= start).map(r => r._i);
    if (idx.length) plan.push({ name, label, idx });
  }
  const total = plan.reduce((n, p) => n + p.idx.length, 0);
  if (b.dryRun) return { counts: plan.map(p => ({ label: p.label, n: p.idx.length })), total };
  const bk = await book(URLS.portal); let done = 0;
  outer: for (const p of plan) for (const i of [...p.idx].sort((a, c) => c - a)) { if (done >= 40) break outer; await g('DELETE', `${bk}/tables/${encodeURIComponent(p.name)}/rows/itemAt(index=${i})`); p.idx = p.idx.filter(x => x !== i); done++; }
  const remaining = plan.reduce((n, p) => n + p.idx.length, 0);
  Object.keys(memo).forEach(k => delete memo[k]);
  if (!remaining) { await putSetting('TestMode', 'Off'); await putSetting('TestStartedAt', ''); }
  return { deleted: done, remaining, live: !remaining };
}

/* ---------------- one-time setup for the opsportal account ---------------- */
function base(req) { return `https://${req.headers['x-forwarded-host'] || req.headers.host}`; }
async function setup(req, res) {
  const state = crypto.randomBytes(16).toString('hex');
  res.setHeader('Set-Cookie', `gto_state=${state}; Path=/api; HttpOnly; Secure; SameSite=Lax; Max-Age=600`);
  const q = new URLSearchParams({ client_id: CLIENT_ID, response_type: 'code', redirect_uri: base(req) + '/api/callback', response_mode: 'query', scope: SCOPE, state, login_hint: OPS, prompt: 'select_account' });
  res.writeHead(302, { Location: `https://login.microsoftonline.com/${TENANT}/oauth2/v2.0/authorize?${q}` }).end();
}
async function callback(req, res) {
  const page = (title, msg) => { res.setHeader('Content-Type', 'text/html; charset=utf-8'); res.end(`<!doctype html><meta name=viewport content="width=device-width,initial-scale=1"><body style="font:16px system-ui;background:#0E1013;color:#F1EDE4;display:grid;place-items:center;min-height:100vh;margin:0"><div style="max-width:520px;padding:28px;border:1px solid rgba(227,185,90,.4);border-radius:16px"><h2 style="color:#E3B95A;margin-top:0">${title}</h2><p>${msg}</p><p><a style="color:#E3B95A" href="/">Open the portal</a></p></div>`); };
  const cookie = String(req.headers.cookie || '').match(/gto_state=([a-f0-9]+)/);
  if (!cookie || cookie[1] !== req.query.state) return page('Setup failed', 'The sign-in link expired. Open /api/setup again.');
  if (req.query.error) return page('Setup failed', String(req.query.error_description || req.query.error).replace(/</g, '&lt;'));
  const r = await fetch(`https://login.microsoftonline.com/${TENANT}/oauth2/v2.0/token`, { method: 'POST', body: new URLSearchParams({ client_id: CLIENT_ID, client_secret: SECRET, grant_type: 'authorization_code', code: String(req.query.code || ''), redirect_uri: base(req) + '/api/callback', scope: SCOPE }) });
  const j = await r.json(); if (!r.ok) return page('Setup failed', String(j.error_description || j.error).replace(/</g, '&lt;'));
  const idp = JSON.parse(Buffer.from(String(j.id_token).split('.')[1], 'base64url').toString());
  const who = normEmail(idp.preferred_username || idp.upn);
  if (who !== OPS) return page('Wrong account', `You signed in as ${who}. Setup must be done with ${OPS}.`);
  await saveRT(j.refresh_token); opsCache = null;
  page('Portal connected', `The portal can now read and write the Excel files as ${OPS}. Staff can sign in.`);
}

/* ---------------- router ---------------- */
export default async function handler(req, res) {
  const route = String(req.query.r || (req.url.includes('/api/callback') ? 'callback' : req.url.includes('/api/setup') ? 'setup' : ''));
  try {
    if (route === 'setup') return await setup(req, res);
    if (route === 'callback') return await callback(req, res);
    if (route === 'health') { const rt = await loadRT(); return res.status(200).json({ ok: true, connected: !!rt, secret: !!SECRET }); }
    const acc = await loadAccess();
    const u = await caller(req, acc);
    const b = req.method === 'POST' ? (typeof req.body === 'string' ? JSON.parse(req.body || '{}') : (req.body || {})) : {};
    let out;
    switch (route) {
      case 'data': out = await readData(u, acc, req.query.parts ? String(req.query.parts).split(',') : null); break;
      case 'rota': { const sheet = String(req.query.sheet || ''); if (!(await loadWeeks()).some(w => w.name === sheet)) bad('Unknown rota week.'); const w = sheetWeekStart(sheet); out = await cached('rota:' + sheet, 300e3, async () => ({ name: sheet, start: dkey(w), people: parseRotaSheet(await sheetText(URLS.rota, sheet), w).map(p => ({ name: p.name, team: p.team, days: p.days })) })); break; }
      case 'att': out = await doAttendance(u, acc, b); break;
      case 'req': out = await doRequest(u, acc, b); break;
      case 'decide': out = await doDecide(u, acc, b); break;
      case 'cancel': out = await doCancel(u, b); break;
      case 'retry': out = await doRetry(u, acc, b); break;
      case 'swap': out = await doSwap(u, acc, b); break;
      case 'swapdecide': out = await doSwapUpdate(u, b, true); break;
      case 'swapcancel': out = await doSwapUpdate(u, b, false); break;
      case 'access': out = await doAccess(u, acc, b); break;
      case 'qalib': out = await readQaLib(u); break;
      case 'qaupload': out = await doQaUpload(u, b); break;
      case 'qadelete': out = await doQaDelete(u, b); break;
      case 'qafile': out = await readQaFile(u, String(req.query.id || '')); break;
      case 'synclog': out = await readSyncLog(u); break;
      case 'syncretry': out = await doSyncRetry(u, acc, b); break;
      case 'rotaedit': out = await doRotaEdit(u, acc, b); break;
      case 'rotaeditclear': out = await doRotaEditClear(u, b); break;
      case 'quality': out = await readQuality(u, acc); break;
      case 'dispute': out = await doDispute(u, acc, b); break;
      case 'disputereview': out = await doDisputeReview(u, b); break;
      case 'disputerespond': out = await doDisputeRespond(u, acc, b); break;
      case 'kb': out = await readKb(u); break;
      case 'swapmates': out = { rows: acc.rows.filter(r => r.Email && !isNo(r.Active) && norm(r.Team) && norm(r.Team) === norm(u.row.Team) && String(r['Portal Role']).trim() === u.role).map(r => ({ Email: r.Email, 'Full Name': r['Full Name'], 'Portal Role': r['Portal Role'], Team: r.Team, 'Rota Name': r['Rota Name'], 'Roster Name': r['Roster Name'], 'Previous Name': r['Previous Name'], Active: r.Active })) }; break;
      case 'updates': out = await readUpdates(); break;
      case 'kbedit': out = await doKbEdit(u, b); break;
      case 'update': out = await doUpdate(u, b); break;
      case 'ask': out = await doAsk(u, b); break;
      case 'asklog': out = await readAskLog(u); break;
      case 'credits': out = await readCredits(u, String(req.query.q || '')); break;
      case 'creditview': out = await doCreditView(u, b); break;
      case 'creditlog': out = await readCreditLog(u); break;
      case 'setting': out = await doSetting(u, b); break;
      case 'golive': out = await doGoLive(u, b); break;
      default: throw new HttpError(404, 'Unknown request.');
    }
    ({ att: ['att'], req: ['req'], decide: ['req', 'tracker'], cancel: ['req'], retry: ['req', 'tracker'], swap: ['swaps'], swapdecide: ['swaps'], swapcancel: ['swaps'], access: ['access'], update: ['updates'], kbedit: ['kbedits'] }[route] || []).forEach(k => forget(k));
    res.setHeader('Cache-Control', 'no-store');
    return res.status(200).json(out);
  } catch (e) {
    const status = e.status || 500; if (status >= 500) console.error(e);
    return res.status(status).json({ error: e.message || 'Server error', ...(e.capacity ? { capacity: e.capacity } : {}) });
  }
}
export const __resetOffers = () => { lastOfferCheck = 0; forget('kb', 'kbedits'); };
