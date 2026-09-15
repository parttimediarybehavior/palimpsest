#!/usr/bin/env node
// palimpsest CLI
//   node cli.js <file.docx> [--json] [--text]
//   node cli.js <file.docx> --scrub <out.docx> [--only id,id] [--all]
import { readFileSync, writeFileSync } from 'node:fs';
import zlib from 'node:zlib';
import { analyze } from './src/engine.js';
import { plan, scrub, verifyScrub } from './src/scrub.js';

const inflate = async (b) => new Uint8Array(zlib.inflateRawSync(b));
const deflate = async (b) => new Uint8Array(zlib.deflateRawSync(b, { level: 6 }));

const args = process.argv.slice(2);
const val = (name) => {
  const i = args.indexOf(name);
  return i >= 0 ? args[i + 1] : null;
};
const scrubTo = val('--scrub');
const only = val('--only');
const consumed = new Set([scrubTo, only].filter(Boolean));
const files = args.filter((a) => !a.startsWith('--') && !consumed.has(a));
const flags = new Set(args.filter((a) => a.startsWith('--')));
if (!files.length) {
  console.error('usage: node cli.js <file.docx> [--json] [--text] [--scrub out.docx [--only id,id] [--all]]');
  process.exit(2);
}

const C = {
  dim: (s) => `\x1b[2m${s}\x1b[0m`,
  b: (s) => `\x1b[1m${s}\x1b[0m`,
  key: (s) => `\x1b[1;38;5;208m${s}\x1b[0m`,
  ok: (s) => `\x1b[38;5;108m${s}\x1b[0m`,
};

for (const f of files) {
  const bytes = new Uint8Array(readFileSync(f));
  let rep;
  try {
    rep = await analyze(bytes, { inflate, name: f.split('/').pop() });
  } catch (e) {
    console.error(`${f}: ${e.message}`);
    continue;
  }

  if (scrubTo) {
    const fixes = only ? only.split(',').map((s) => s.trim()) : flags.has('--all') ? plan(rep).map((x) => x.id) : null;
    const res = await scrub(bytes, rep, { inflate, deflate, fixes });
    writeFileSync(scrubTo, res.bytes);
    const after = await analyze(res.bytes, { inflate, name: scrubTo.split('/').pop() });
    const v = verifyScrub(rep, after);
    console.log('\n' + C.b(rep.file.name) + C.dim(` → ${scrubTo}  ${(rep.file.bytes / 1024).toFixed(0)} KB → ${(res.bytes.length / 1024).toFixed(0)} KB`));
    console.log(C.b('\napplied'));
    for (const p of plan(rep)) {
      const hit = res.applied.find((a) => a.id === p.id);
      const chosen = fixes ? fixes.includes(p.id) : p.on;
      console.log(`  ${chosen ? C.ok('✓') : C.dim('·')} ${p.id.padEnd(11)} ${C.dim(p.label.padEnd(42))} ${hit ? C.b(hit.n + '') : C.dim('—')}`);
    }
    for (const n of res.notes) console.log('  ' + C.dim('note: ' + n));
    console.log(C.b('\nbefore → after'));
    for (const k of Object.keys(v.before)) {
      const a = v.before[k];
      const b = v.after[k];
      if (a === b && !a) continue;
      const mark =
        k === 'words' || k === 'chars'
          ? C.dim(a === b ? 'unchanged' : `${b - a} (removed content)`)
          : b === 0 && a > 0
            ? C.ok('clean')
            : b < a
              ? C.ok('reduced')
              : b > a
                ? C.key('worse')
                : '';
      console.log(`  ${k.padEnd(14)} ${String(a).padStart(7)} → ${String(b).padStart(7)}  ${mark}`);
    }
    if (v.text) {
      console.log(
        `  ${'visible text'.padEnd(14)} ${v.text.ok ? C.ok(`${v.text.words.after} of ${v.text.words.before} words kept; everything dropped was flagged content`) : C.key(`MISMATCH invented=${JSON.stringify(v.text.invented)} lost=${JSON.stringify(v.text.lost)}`)}`,
      );
    }
    console.log(C.b('\nfindings still standing'));
    if (!v.remaining.length) console.log('  ' + C.ok('none'));
    for (const r of v.remaining) console.log('  ' + C.dim('· ') + r);
    continue;
  }
  delete rep._model;
  delete rep._plain;
  if (flags.has('--json')) {
    console.log(JSON.stringify(rep, null, 2));
    continue;
  }
  console.log('\n' + C.b(rep.file.name) + C.dim(`  ${(rep.file.bytes / 1024).toFixed(0)} KB · ${rep.container.partCount} parts`));
  console.log(C.key('» ' + rep.verdict.headline));
  for (const l of rep.verdict.lines) console.log('  ' + C.dim('·') + ' ' + l);
  console.log();
  for (const fd of rep.findings) {
    const tag = fd.level === 'key' ? C.key('[!]') : C.dim('[i]');
    console.log(`${tag} ${C.b(fd.title)}`);
    console.log('    ' + C.dim(fd.detail.replace(/\s+/g, ' ')));
  }
  if (rep.sessions.some((s) => s.chars)) {
    console.log('\n' + C.b('sessions') + C.dim('  (layer = position in the inferred partial order)'));
    for (const s of rep.sessions.filter((x) => x.chars).slice(0, 14)) {
      console.log(
        `  ${s.rsid}${s.isRoot ? C.ok(' ROOT') : '     '} layer ${String(s.layer).padStart(2)}  ${String(s.chars).padStart(6)} chars ${String(s.share).padStart(5)}%  ¶${s.span ? s.span[0] + 1 + '-' + (s.span[1] + 1) : '-'}${s.anchoredAt ? C.ok('  @ ' + s.anchoredAt.slice(0, 16).replace('T', ' ')) : ''}${s.langs.length > 1 ? C.dim('  langs:' + s.langs.map((l) => l.name).join('/')) : ''}`,
      );
    }
  }
  if (rep.deletedText.length) {
    console.log('\n' + C.b('recovered deletions'));
    for (const d of rep.deletedText.slice(0, 8)) console.log(`  ${C.dim(d.author || '?')} ${JSON.stringify(d.text.slice(0, 110))}`);
  }
  if (rep.hiddenText.length) {
    console.log('\n' + C.b('invisible text'));
    for (const d of rep.hiddenText.slice(0, 8)) console.log(`  ${C.dim(d.kind)} ${JSON.stringify(d.text.slice(0, 110))}`);
  }
  if (rep.media.length) {
    console.log('\n' + C.b('media'));
    for (const m of rep.media) console.log(`  ${m.part.padEnd(26)} ${String(m.format).padEnd(6)} ${m.width || '?'}x${m.height || '?'} ${m.cropped ? C.key(`crop→${m.visibleFraction}% visible`) : ''} ${Object.keys(m.meta).length ? C.dim(JSON.stringify(m.meta).slice(0, 160)) : ''}`);
  }
  if (flags.has('--text')) console.log('\n' + C.b('text preview') + '\n' + rep.text.preview);
}
