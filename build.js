#!/usr/bin/env node
// build.js — inline every module into one self-contained HTML file.
// No bundler: the modules are plain ES modules with unique top-level names,
// so stripping import/export lines and concatenating is sufficient (and auditable).
import { readFileSync, writeFileSync, existsSync, readdirSync, mkdirSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { spawnSync } from 'node:child_process';
import { tmpdir } from 'node:os';

const here = dirname(new URL(import.meta.url).pathname);
const src = join(here, 'src');
const read = (f) => readFileSync(join(src, f), 'utf8');

const strip = (code) =>
  code
    .replace(/^\s*import[^;]*?;\s*$/gm, '')
    .replace(/^export\s+(?=(const|function|async|class|let|var))/gm, '')
    .replace(/^export\s+\{[^}]*\};?\s*$/gm, '');

const MODULE_FILES = ['zipxml.js', 'zipwrite.js', 'media.js', 'engine.js', 'scrub.js', 'ui.js'];
const modules = MODULE_FILES.map((f) => `/* ===== ${f} ===== */\n` + strip(read(f)));

// Concatenation only works while the modules' top-level names stay disjoint.
// Two files declaring `const td` produce a bundle that dies on load, so the
// build refuses to write one.
const owners = new Map();
const clashes = [];
MODULE_FILES.forEach((f, i) => {
  for (const m of modules[i].matchAll(/^(?:const|let|var|function|async function|class)\s+([A-Za-z_$][\w$]*)/gm)) {
    const name = m[1] === 'function' ? null : m[1];
    if (!name) continue;
    if (owners.has(name)) clashes.push(`${name}: ${owners.get(name)} and ${f}`);
    else owners.set(name, f);
  }
});
if (clashes.length) {
  console.error('duplicate top-level names would break the bundle:\n  ' + clashes.join('\n  '));
  process.exit(1);
}

// optional: embed sample documents as base64 so the single file demos itself
const sampleDir = join(here, 'samples');
let samples = [];
if (existsSync(sampleDir)) {
  samples = readdirSync(sampleDir)
    .filter((f) => f.endsWith('.docx'))
    .sort()
    .map((f) => {
      const meta = existsSync(join(sampleDir, f + '.json')) ? JSON.parse(readFileSync(join(sampleDir, f + '.json'), 'utf8')) : {};
      return { file: f, label: meta.label || f, note: meta.note || '', pair: meta.pair || null, hidden: !!meta.hidden, b64: readFileSync(join(sampleDir, f)).toString('base64') };
    });
}

const bootstrap = `
const SAMPLES = ${JSON.stringify(samples.map(({ file, label, note, pair, b64 }) => ({ file, label, note, pair, b64 })))};
function b64bytes(b64) {
  const bin = atob(b64);
  const u8 = new Uint8Array(bin.length);
  for (let i = 0; i < bin.length; i++) u8[i] = bin.charCodeAt(i);
  return u8;
}
document.querySelectorAll('.samples button[data-sample]').forEach((b) => {
  b.onclick = async (e) => {
    e.stopPropagation();
    const s = SAMPLES[+b.dataset.sample];
    await loadFile(new File([b64bytes(s.b64)], s.file));
    if (s.pair) {
      const other = SAMPLES.find((x) => x.file === s.pair);
      if (other) await compareWith(new File([b64bytes(other.b64)], other.file));
    }
  };
});
wireUp();
`;

const sampleHtml = samples
  .map((s, i) => (s.hidden ? '' : `<button data-sample="${i}" title="${s.note.replace(/"/g, '&quot;')}">${s.label}</button>`))
  .filter(Boolean)
  .join('\n      ');

const html = read('index.html')
  .replace('<!--CSS-->', '<style>\n' + read('app.css') + '\n</style>')
  .replace('<!--SAMPLES-->', sampleHtml)
  .replace('<!--JS-->', '<script type="module">\n' + modules.join('\n') + '\n' + bootstrap + '\n</script>');

mkdirSync(join(here, 'dist'), { recursive: true });
const out = join(here, 'dist', 'palimpsest.html');
writeFileSync(out, html);
// and the bundle has to parse before it ships
const probe = join(tmpdir(), 'palimpsest-bundle-check.mjs');
writeFileSync(probe, modules.join('\n'));
const check = spawnSync(process.execPath, ['--check', probe], { encoding: 'utf8' });
if (check.status !== 0) {
  console.error('bundle does not parse:\n' + (check.stderr || '').split('\n').slice(0, 6).join('\n'));
  process.exit(1);
}

console.log(`${out}  ${(html.length / 1024).toFixed(0)} KB${samples.length ? `  (${samples.length} embedded samples)` : ''}`);
