#!/usr/bin/env node
// scrub-test.js — run the scrubber across a whole directory of real documents
// and check every output: zip integrity, XML well-formedness, relationship and
// content-type consistency, re-analysis, and byte-for-byte text preservation.
//
//   node scrub-test.js corpus [--keep out/] [--all]
import { readdirSync, readFileSync, writeFileSync, mkdirSync } from 'node:fs';
import { join } from 'node:path';
import zlib from 'node:zlib';
import { analyze } from './src/engine.js';
import { plan, scrub, verifyScrub } from './src/scrub.js';
import { readZip, parseXML, findAll, walk } from './src/zipxml.js';

const inflate = async (b) => new Uint8Array(zlib.inflateRawSync(b));
const deflate = async (b) => new Uint8Array(zlib.deflateRawSync(b, { level: 6 }));

const args = process.argv.slice(2);
const dir = args.find((a) => !a.startsWith('--')) || 'corpus';
const keepIdx = args.indexOf('--keep');
const keep = keepIdx >= 0 ? args[keepIdx + 1] : null;
const useAll = args.includes('--all');
if (keep) mkdirSync(keep, { recursive: true });

const C = {
  dim: (s) => `\x1b[2m${s}\x1b[0m`,
  b: (s) => `\x1b[1m${s}\x1b[0m`,
  bad: (s) => `\x1b[1;38;5;203m${s}\x1b[0m`,
  ok: (s) => `\x1b[38;5;108m${s}\x1b[0m`,
};

/* ---- a stricter XML check than the engine's tolerant parser ------------- */
// Balanced-tag scan: the engine's parser is deliberately forgiving, so it would
// happily read a document Word refuses to open. This one is not forgiving.
function xmlWellFormed(text) {
  const stack = [];
  const re = /<([!?/]?)([A-Za-z_][\w.:-]*)?([^>]*?)(\/?)>/g;
  let m;
  while ((m = re.exec(text))) {
    const [, kind, name, attrs, selfClose] = m;
    if (kind === '!' || kind === '?') continue;
    if (kind === '/') {
      if (stack.pop() !== name) return `unbalanced </${name}> at ${m.index}`;
      continue;
    }
    if (!name) return `bad tag at ${m.index}`;
    // an unquoted '>' inside an attribute would desync this scan
    const quotes = (attrs.match(/"/g) || []).length;
    if (quotes % 2) return `unbalanced quotes in <${name}> at ${m.index}`;
    if (!selfClose) stack.push(name);
  }
  return stack.length ? `unclosed <${stack[stack.length - 1]}>` : null;
}

async function checkContainer(bytes) {
  const problems = [];
  const zip = await readZip(bytes, inflate);
  const names = new Set(zip.order);
  if (!names.has('[Content_Types].xml')) problems.push('no [Content_Types].xml');
  if (zip.order[0] !== '[Content_Types].xml') problems.push('[Content_Types].xml is not the first entry');
  if (!names.has('word/document.xml')) problems.push('no word/document.xml');

  const types = parseXML(await zip.text('[Content_Types].xml'));
  const defaults = new Set(findAll(types, 'Default').map((d) => (d.attrs.Extension || '').toLowerCase()));
  const overrides = new Set(findAll(types, 'Override').map((o) => (o.attrs.PartName || '').replace(/^\//, '')));
  for (const n of zip.order) {
    if (n.endsWith('/') || n === '[Content_Types].xml') continue;
    const ext = (n.split('.').pop() || '').toLowerCase();
    if (!overrides.has(n) && !defaults.has(ext)) problems.push(`part with no content type: ${n}`);
  }
  for (const pn of overrides) if (!names.has(pn)) problems.push(`content type for a missing part: ${pn}`);

  for (const n of zip.order) {
    if (!/\.(xml|rels)$/i.test(n)) continue;
    const text = await zip.text(n);
    const err = xmlWellFormed(text);
    if (err) problems.push(`${n}: ${err}`);
  }

  // every r:id used in a part must exist in that part's .rels
  for (const n of zip.order) {
    if (!/^word\/.*\.xml$/.test(n) && n !== 'word/document.xml') continue;
    const relsName = n.replace(/([^/]+)$/, '_rels/$1.rels');
    const relIds = new Set();
    if (names.has(relsName)) {
      for (const r of findAll(parseXML(await zip.text(relsName)), 'Relationship')) relIds.add(r.attrs.Id);
    }
    const text = await zip.text(n);
    const used = new Set();
    for (const m of text.matchAll(/\sr:(?:id|embed|link|dm|lo|qs|cs|href|pict)="([^"]+)"/g)) used.add(m[1]);
    for (const id of used) if (!relIds.has(id)) problems.push(`${n} refers to ${id}, which its .rels does not define`);
  }

  // no dangling rels pointing at parts we removed
  for (const n of zip.order) {
    if (!/_rels\/.*\.rels$/.test(n)) continue;
    const dirp = n.replace(/_rels\/[^/]*$/, '');
    for (const r of findAll(parseXML(await zip.text(n)), 'Relationship')) {
      const t = r.attrs.Target || '';
      if (r.attrs.TargetMode === 'External' || /^[a-z]+:/i.test(t) || t.startsWith('//') || t.startsWith('\\\\')) continue;
      const segs = [];
      for (const s of (t.startsWith('/') ? t.slice(1) : dirp + t.replace(/^\.\//, '')).split('/')) {
        if (s === '..') segs.pop();
        else if (s && s !== '.') segs.push(s);
      }
      const resolved = segs.join('/');
      if (!names.has(resolved)) problems.push(`${n} points at missing part ${resolved}`);
    }
  }

  // things the scrub is supposed to have removed must really be gone
  const doc = await zip.text('word/document.xml');
  if (/\sw:rsid[A-Za-z]*="/.test(doc)) problems.push('rsid attributes survive in document.xml');
  if (/<w:(?:del|ins)\b/.test(doc)) problems.push('tracked changes survive in document.xml');
  return problems;
}

/* ---- run ---------------------------------------------------------------- */
const files = readdirSync(dir).filter((f) => f.toLowerCase().endsWith('.docx')).sort();
let okCount = 0;
let failCount = 0;
let textOk = 0;
let textBad = 0;
const totals = { before: 0, after: 0, bytesBefore: 0, bytesAfter: 0 };
const fixTotals = new Map();
const failures = [];
const remainTally = new Map();

for (const f of files) {
  const path = join(dir, f);
  const bytes = new Uint8Array(readFileSync(path));
  let rep;
  try {
    rep = await analyze(bytes, { inflate, name: f });
  } catch {
    continue; // not our problem here; scan.js covers parse failures
  }
  let res;
  let after;
  try {
    const fixes = useAll ? plan(rep).map((x) => x.id) : null;
    res = await scrub(bytes, rep, { inflate, deflate, fixes });
    after = await analyze(res.bytes, { inflate, name: f });
  } catch (e) {
    failCount++;
    failures.push(`${f}: scrub threw — ${e.message}`);
    continue;
  }
  const problems = await checkContainer(res.bytes);
  const v = verifyScrub(rep, after);
  if (v.text && v.text.ok) textOk++;
  else {
    textBad++;
    problems.push(`text changed beyond the intended removals: invented=${JSON.stringify(v.text && v.text.invented)} lost=${JSON.stringify(v.text && v.text.lost)}`);
  }
  for (const a of res.applied) if (!a.id.endsWith('Bytes')) fixTotals.set(a.id, (fixTotals.get(a.id) || 0) + a.n);
  for (const r of v.remaining) remainTally.set(r.replace(/\d+/g, 'N'), (remainTally.get(r.replace(/\d+/g, 'N')) || 0) + 1);
  totals.before += v.before.findings;
  totals.after += v.after.findings;
  totals.bytesBefore += bytes.length;
  totals.bytesAfter += res.bytes.length;
  if (problems.length) {
    failCount++;
    failures.push(`${f}:\n    ${problems.slice(0, 6).join('\n    ')}`);
  } else okCount++;
  if (keep) writeFileSync(join(keep, f.replace(/\.docx$/i, '.clean.docx')), res.bytes);
  process.stdout.write(problems.length ? C.bad('x') : '.');
}

console.log(`\n\nscrubbed ${files.length} documents from ${dir}/`);
console.log(`  ${C.b(String(okCount))} produced a structurally valid .docx with text preserved`);
if (failCount) console.log(`  ${C.bad(String(failCount))} had problems`);
console.log(`  text preserved exactly: ${C.b(`${textOk}/${textOk + textBad}`)}`);
console.log(`  actionable findings: ${totals.before} → ${C.ok(String(totals.after))} (${Math.round((1 - totals.after / Math.max(1, totals.before)) * 100)}% removed)`);
console.log(`  bytes: ${(totals.bytesBefore / 1024).toFixed(0)} KB → ${(totals.bytesAfter / 1024).toFixed(0)} KB`);
console.log('\nwhat was removed across the corpus:');
for (const [id, n] of [...fixTotals].sort((a, b) => b[1] - a[1])) console.log(`  ${id.padEnd(12)} ${String(n).padStart(8)}`);
console.log('\nfindings that survive the scrub:');
for (const [t, n] of [...remainTally].sort((a, b) => b[1] - a[1])) console.log(`  ${String(n).padStart(3)}  ${t}`);
if (failures.length) {
  console.log('\nfailures:');
  for (const f of failures.slice(0, 12)) console.log('  ' + f);
}
process.exit(failCount ? 1 : 0);
