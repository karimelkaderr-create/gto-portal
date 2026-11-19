// Minimal .xlsx reader (no dependencies): returns the cell values and text of one sheet.
// Used for workbooks that Excel for the web (and so Microsoft Graph's workbook API) can't open.
import { inflateRawSync } from 'node:zlib';

export function unzip(buf) {
  let eocd = -1; for (let i = buf.length - 22; i >= Math.max(0, buf.length - 70000); i--) if (buf.readUInt32LE(i) === 0x06054b50) { eocd = i; break; }
  if (eocd < 0) throw new Error('not a valid .xlsx file');
  const count = buf.readUInt16LE(eocd + 10); let p = buf.readUInt32LE(eocd + 16); const files = {};
  for (let n = 0; n < count; n++) {
    if (buf.readUInt32LE(p) !== 0x02014b50) break;
    const method = buf.readUInt16LE(p + 10), csize = buf.readUInt32LE(p + 20), nameLen = buf.readUInt16LE(p + 28), extraLen = buf.readUInt16LE(p + 30), commentLen = buf.readUInt16LE(p + 32), off = buf.readUInt32LE(p + 42);
    const name = buf.toString('utf8', p + 46, p + 46 + nameLen); p += 46 + nameLen + extraLen + commentLen;
    files[name] = () => { const lh = off; const start = lh + 30 + buf.readUInt16LE(lh + 26) + buf.readUInt16LE(lh + 28); const data = buf.subarray(start, start + csize); return (method === 8 ? inflateRawSync(data) : data).toString('utf8'); };
  }
  return files;
}
const unesc = s => String(s).replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"').replace(/&apos;/g, "'").replace(/&#(\d+);/g, (_, d) => String.fromCodePoint(+d)).replace(/&#x([0-9a-f]+);/gi, (_, h) => String.fromCodePoint(parseInt(h, 16))).replace(/&amp;/g, '&');
const colIndex = ref => { let n = 0; for (const ch of ref.replace(/\d+/g, '')) n = n * 26 + ch.charCodeAt(0) - 64; return n - 1; };
const textOf = xml => { let out = ''; for (const m of xml.matchAll(/<t(?:\s[^>]*)?>([\s\S]*?)<\/t>/g)) out += m[1]; return unesc(out); };

export function sheetNames(buf) {
  const f = unzip(buf); const wb = f['xl/workbook.xml']; if (!wb) throw new Error('not an Excel workbook');
  return [...wb().matchAll(/<sheet\b[^>]*>/g)].map(m => ({ name: unesc((m[0].match(/\bname="([^"]*)"/) || [])[1] || ''), rid: (m[0].match(/\br:id="([^"]*)"/) || [])[1], hidden: /state="(hidden|veryHidden)"/.test(m[0]) }));
}
export function readSheet(buf, pick) {
  const f = unzip(buf); const sheets = sheetNames(buf).filter(s => !s.hidden);
  const hit = pick(sheets.map(s => s.name)); if (!hit) return { names: sheets.map(s => s.name), values: null, text: null };
  const sh = sheets.find(s => s.name === hit);
  const rels = f['xl/_rels/workbook.xml.rels'](); const rel = [...rels.matchAll(/<Relationship\b[^>]*>/g)].map(m => m[0]).find(x => (x.match(/\bId="([^"]*)"/) || [])[1] === sh.rid);
  let target = (rel.match(/\bTarget="([^"]*)"/) || [])[1]; target = target.startsWith('/') ? target.slice(1) : 'xl/' + target;
  const shared = []; if (f['xl/sharedStrings.xml']) for (const m of f['xl/sharedStrings.xml']().matchAll(/<si\b[^>]*>([\s\S]*?)<\/si>/g)) shared.push(textOf(m[1]));
  const xml = f[target](); const values = [], text = []; let maxCol = 0;
  for (const m of xml.matchAll(/<c\b([^>]*?)(?:\/>|>([\s\S]*?)<\/c>)/g)) {
    const attrs = m[1], inner = m[2] || ''; const ref = (attrs.match(/\br="([A-Z]+\d+)"/) || [])[1]; if (!ref) continue;
    const t = (attrs.match(/\bt="(\w+)"/) || [])[1]; const v = (inner.match(/<v>([\s\S]*?)<\/v>/) || [])[1];
    let val = ''; if (t === 'inlineStr') val = textOf(inner); else if (v !== undefined) val = t === 's' ? (shared[+v] ?? '') : t === 'str' ? unesc(v) : t === 'b' ? v === '1' : t === 'e' ? '' : parseFloat(v);
    if (val === '' || (typeof val === 'number' && isNaN(val))) continue;
    const r = +ref.replace(/[A-Z]+/, '') - 1, c = colIndex(ref); if (c > 60 || r > 200000) continue; maxCol = Math.max(maxCol, c);
    (values[r] = values[r] || [])[c] = val; (text[r] = text[r] || [])[c] = typeof val === 'number' ? String(val) : String(val);
  }
  const rows = values.length, width = maxCol + 1; const fill = a => { for (let r = 0; r < rows; r++) { const row = a[r] || []; a[r] = Array.from({ length: width }, (_, c) => row[c] ?? ''); } return a; };
  return { names: sheets.map(s => s.name), sheet: hit, values: fill(values), text: fill(text) };
}
