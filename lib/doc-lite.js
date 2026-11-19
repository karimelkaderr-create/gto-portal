// Dependency-free text extraction for .docx and .pptx (turns a guide into searchable sections).
import { unzip } from './xlsx-lite.js';
const unesc = s => String(s).replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"').replace(/&apos;/g, "'").replace(/&#(\d+);/g, (_, d) => String.fromCharCode(+d)).replace(/&#x([0-9a-f]+);/gi, (_, h) => String.fromCharCode(parseInt(h, 16))).replace(/&amp;/g, '&');
const runs = (xml, tag) => { let out = ''; for (const m of xml.matchAll(new RegExp(`<${tag}(?:\\s[^>]*)?>([\\s\\S]*?)</${tag}>|<w:tab/>|<w:br/>`, 'g'))) out += m[1] !== undefined ? m[1] : ' '; return unesc(out).replace(/\s+/g, ' ').trim(); };
const clip = (s, n) => String(s).slice(0, n);

export function docxSections(buf) {
  const f = unzip(buf); if (!f['word/document.xml']) throw new Error('not a Word document');
  const xml = f['word/document.xml']();
  const body = xml.replace(/<w:tbl>[\s\S]*?<\/w:tbl>/g, t => '<w:p><w:r><w:t>' + '\u0001' + [...t.matchAll(/<w:tr\b[\s\S]*?<\/w:tr>/g)].map(r => [...r[0].matchAll(/<w:tc>[\s\S]*?<\/w:tc>/g)].map(c => runs(c[0], 'w:t')).join(' | ')).join('\u0001') + '</w:t></w:r></w:p>');
  const sections = []; let cur = { heading: 'Introduction', lines: [] };
  for (const p of body.matchAll(/<w:p\b[\s\S]*?<\/w:p>/g)) {
    const text = runs(p[0], 'w:t'); if (!text) continue;
    const st = (p[0].match(/<w:pStyle w:val="([^"]+)"/) || [])[1] || '';
    const isHead = /^(Heading|Title)/i.test(st) && text.length < 140;
    if (isHead) { if (cur.lines.length) sections.push(cur); cur = { heading: clip(text, 140), lines: [] }; }
    else for (const l of text.split('\u0001')) if (l.trim()) cur.lines.push(l.trim());
  }
  if (cur.lines.length) sections.push(cur);
  // a guide without heading styles becomes chunks of about 1,500 characters
  if (sections.length <= 1) { const all = sections.flatMap(s => s.lines); const out = []; let c = { heading: 'Part 1', lines: [] }, n = 0; for (const l of all) { c.lines.push(l); n += l.length; if (n > 1500) { out.push(c); c = { heading: 'Part ' + (out.length + 1), lines: [] }; n = 0; } } if (c.lines.length) out.push(c); return out.map(s => ({ heading: s.heading, text: s.lines.join('\n') })); }
  return sections.map(s => ({ heading: s.heading, text: s.lines.join('\n') }));
}
export function pptxSections(buf) {
  const f = unzip(buf); const names = Object.keys(f).filter(n => /^ppt\/slides\/slide\d+\.xml$/.test(n)).sort((a, b) => +a.match(/\d+/)[0] - +b.match(/\d+/)[0]);
  if (!names.length) throw new Error('not a PowerPoint file');
  const out = [];
  names.forEach((n, i) => {
    const xml = f[n](); const lines = [];
    for (const sp of xml.matchAll(/<p:(?:sp|graphicFrame)\b[\s\S]*?<\/p:(?:sp|graphicFrame)>/g)) {
      if (/<a:tbl>/.test(sp[0])) { for (const r of sp[0].matchAll(/<a:tr\b[\s\S]*?<\/a:tr>/g)) { const cells = [...r[0].matchAll(/<a:tc\b[\s\S]*?<\/a:tc>/g)].map(c => runs(c[0], 'a:t')); if (cells.some(Boolean)) lines.push(cells.join(' | ')); } continue; }
      for (const p of sp[0].matchAll(/<a:p>[\s\S]*?<\/a:p>/g)) { const t = runs(p[0], 'a:t'); if (t) lines.push(t); }
    }
    if (!lines.length) return;
    out.push({ heading: clip(lines[0], 120) || 'Slide ' + (i + 1), text: lines.slice(lines.length > 1 ? 1 : 0).join('\n'), slide: i + 1 });
  });
  return out;
}
export function sectionsOf(name, buf) { return /\.pptx$/i.test(name) ? pptxSections(buf) : /\.docx$/i.test(name) ? docxSections(buf) : (() => { throw new Error('Only .docx and .pptx files are supported.'); })(); }
