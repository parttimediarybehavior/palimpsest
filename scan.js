#!/usr/bin/env node
// scan.js — run the engine over a directory of .docx files and report both
// robustness (does it survive real-world files?) and prevalence (how often does
// each leak actually occur in documents people published?).
//   node scan.js corpus [--json out.json]
import { readFileSync, readdirSync, writeFileSync } from 'node:fs';
import { join } from 'node:path';
import zlib from 'node:zlib';
import { analyze } from './src/engine.js';

const inflate = async (b) => new Uint8Array(zlib.inflateRawSync(b));
const dir = process.argv[2] || 'corpus';
const jsonOut = process.argv.includes('--json') ? process.argv[process.argv.indexOf('--json') + 1] : null;
const files = readdirSync(dir).filter((f) => /\.docx$/i.test(f)).sort();

const rows = [];
const failures = [];
for (const f of files) {
  const bytes = new Uint8Array(readFileSync(join(dir, f)));
  const t0 = performance.now();
  try {
    const rep = await analyze(bytes, { inflate, name: f });
    const ms = performance.now() - t0;
    const key = rep.findings.filter((x) => x.level === 'key');
    rows.push({
      file: f,
      kb: Math.round(bytes.length / 1024),
      ms: Math.round(ms),
      producer: rep.props.application || '(none)',
      appVersion: rep.props.appVersionLabel || rep.props.appVersion || null,
      words: Math.max(rep.text.words || 0, rep.props.words || 0),
      sessions: rep.settings.rsidCount,
      sessionsWithText: rep.sessions.filter((s) => s.native && s.chars > 0).length,
      ghostSessions: rep.sessions.filter((s) => s.native && !s.chars).length,
      foreign: rep.sessions.filter((s) => !s.native).length,
      topShare: Math.max(0, ...rep.sessions.filter((s) => s.native).map((s) => s.share)),
      names: rep.identities.length,
      emails: (rep.identityIds || []).filter((i) => i.email).length,
      company: !!rep.props.company,
      template: rep.props.template && !/^Normal(\.dotm)?$/.test(rep.props.template) ? rep.props.template : null,
      totalTime: rep.props.totalTime,
      revisions: rep.timeline.revisions,
      printed: !!rep.props.lastPrinted,
      deletions: rep.deletedText.length,
      hidden: rep.hiddenText.length,
      comments: rep.comments.length,
      images: rep.media.length,
      cropped: rep.media.filter((m) => m.cropped && m.visibleFraction < 99).length,
      imageMeta: rep.media.filter((m) => Object.keys(m.meta || {}).length).length,
      gps: rep.media.filter((m) => m.meta && m.meta.GPS).length,
      unusedImages: rep.media.filter((m) => m.unused).length,
      labels: rep.labels.length,
      external: rep.external.length,
      embedded: rep.embedded.filter((e) => e.kind === 'embedded object').length,
      macros: rep.embedded.some((e) => e.kind === 'VBA macro project'),
      glossary: rep.embedded.some((e) => e.kind === 'AutoText / building blocks'),
      locales: rep.locales.length,
      keyFindings: key.length,
      headline: rep.verdict.headline,
      keyTitles: key.map((k) => k.title),
    });
  } catch (e) {
    failures.push({ file: f, error: e.message });
  }
}

const n = rows.length;
const word = rows.filter((r) => r.sessions > 0);
const share = (pred, base = rows) => {
  const m = base.filter(pred).length;
  return `${m}/${base.length} (${base.length ? Math.round((m / base.length) * 100) : 0}%)`;
};
const median = (xs) => {
  const s = xs.filter((x) => x != null).sort((a, b) => a - b);
  return s.length ? s[Math.floor(s.length / 2)] : null;
};

const lines = [];
const say = (s = '') => lines.push(s);
say(`scanned ${n} documents from ${dir}/  (${failures.length} failed to parse)`);
say(`engine time: median ${median(rows.map((r) => r.ms))} ms/file, worst ${Math.max(...rows.map((r) => r.ms))} ms`);
say();
say('producers:');
const byProducer = new Map();
for (const r of rows) byProducer.set(r.producer, (byProducer.get(r.producer) || 0) + 1);
for (const [p, c] of [...byProducer.entries()].sort((a, b) => b[1] - a[1])) say(`  ${String(c).padStart(3)}  ${p}`);
say();
say('what leaks, across the whole corpus:');
say(`  a name (author / last editor / reviewer)      ${share((r) => r.names > 0)}`);
say(`  two or more distinct names                    ${share((r) => r.names > 1)}`);
say(`  an email address or account id                ${share((r) => r.emails > 0)}`);
say(`  employer recorded in Company                  ${share((r) => r.company)}`);
say(`  a named (non-default) template                ${share((r) => !!r.template)}`);
say(`  editing-session ids (Word lineage)            ${share((r) => r.sessions > 0)}`);
say(`  recoverable tracked deletions                 ${share((r) => r.deletions > 0)}`);
say(`  text invisible on screen                      ${share((r) => r.hidden > 0)}`);
say(`  comments left in the file                     ${share((r) => r.comments > 0)}`);
say(`  images cropped with the rest still inside     ${share((r) => r.cropped > 0)}`);
say(`  images carrying their own metadata            ${share((r) => r.imageMeta > 0)}`);
say(`  images with GPS coordinates                   ${share((r) => r.gps > 0)}`);
say(`  images present but never displayed            ${share((r) => r.unusedImages > 0)}`);
say(`  whole embedded files (workbooks etc.)         ${share((r) => r.embedded > 0)}`);
say(`  AutoText / building blocks                    ${share((r) => r.glossary)}`);
say(`  macros                                        ${share((r) => r.macros)}`);
say(`  Microsoft sensitivity label                   ${share((r) => r.labels > 0)}`);
say(`  external / local-path references              ${share((r) => r.external > 0)}`);
say(`  more than one language locale in the runs     ${share((r) => r.locales > 1)}`);
say(`  print history                                 ${share((r) => r.printed)}`);
say();
say(`among the ${word.length} Word-produced files with session ids:`);
say(`  median sessions recorded                      ${median(word.map((r) => r.sessions))}`);
say(`  median sessions that still own text           ${median(word.map((r) => r.sessionsWithText))}`);
say(`  median sessions fully overwritten             ${median(word.map((r) => r.ghostSessions))}`);
say(`  carry save-ids from other documents           ${share((r) => r.foreign > 0, word)}`);
say(`  median foreign ids when present               ${median(word.filter((r) => r.foreign).map((r) => r.foreign))}`);
say(`  one session owns >=85% of the text            ${share((r) => r.topShare >= 85, word)}`);
const typed = word.filter((r) => r.words > 300 && r.totalTime != null);
say(`  >300 words with <=1 recorded editing minute   ${share((r) => r.totalTime <= 1, typed)}`);
say(`  implied typing speed >120 wpm                 ${share((r) => r.totalTime > 0 && r.words / r.totalTime > 120, typed)}`);
say();
say('leakiest documents:');
for (const r of [...rows].sort((a, b) => b.keyFindings - a.keyFindings).slice(0, 12)) {
  say(`  ${String(r.keyFindings).padStart(2)} findings · ${r.file.slice(0, 54).padEnd(54)} ${r.headline}`);
}
if (failures.length) {
  say();
  say('failures:');
  for (const f of failures) say(`  ${f.file}: ${f.error}`);
}
console.log(lines.join('\n'));
if (jsonOut) {
  writeFileSync(jsonOut, JSON.stringify({ generated: new Date().toISOString(), dir, rows, failures, summary: lines }, null, 1));
  console.error(`\nwrote ${jsonOut}`);
}
