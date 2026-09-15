#!/usr/bin/env node
// validate-lineage.js — ground-truth test of the lineage inference.
// Several corpus documents are dated *in their filenames* (successive drafts of
// the same federal form). The engine never sees those dates: it only sees the
// sets of editing-session ids. So the filename order is an independent answer key.
import { readFileSync, readdirSync } from 'node:fs';
import zlib from 'node:zlib';
import { analyze, compareDocs } from './src/engine.js';

const inflate = async (b) => new Uint8Array(zlib.inflateRawSync(b));
const load = async (f) => analyze(new Uint8Array(readFileSync('corpus/' + f)), { inflate, name: f });

const groups = [
  {
    name: 'NIH eSubmission validations drafts (date in filename)',
    order: ['eSub-validations-draft-1-12-2016', 'eSub-validations-draft-2-5-2016', 'eSub-validations-draft-3-1-2016', 'eSub-validations-draft-3-31-2016', 'eSub-validations-draft-7-13-2016'],
  },
  {
    name: 'NIST Baldrige eligibility form (year in filename)',
    order: ['2014_Baldrige_Eligibility_Form', '2015_Baldrige_Eligibility_Form', '2016_Baldrige_Eligibility_Form'],
  },
  {
    name: 'NIST Baldrige award application form (year in filename)',
    order: ['2014_Baldrige_Award_Application_Form', '2015_Baldrige_Award_Application_Form'],
  },
];

const files = readdirSync('corpus');
let pass = 0, fail = 0, skipped = 0;
for (const g of groups) {
  const resolved = g.order.map((stem) => files.find((f) => f.includes(stem))).filter(Boolean);
  if (resolved.length < 2) {
    console.log(`\n${g.name}: not enough files in corpus, skipping`);
    skipped++;
    continue;
  }
  console.log(`\n\x1b[1m${g.name}\x1b[0m  (${resolved.length} files)`);
  const reps = {};
  for (const f of resolved) reps[f] = await load(f);
  for (let i = 0; i < resolved.length; i++) {
    for (let j = 0; j < resolved.length; j++) {
      if (i === j) continue;
      const c = compareDocs(reps[resolved[i]], reps[resolved[j]], 'B');
      const want = i < j ? 'descendant' : 'ancestor';
      const good = c.relation === want;
      if (good) pass++; else fail++;
      const arrow = i < j ? '→' : '←';
      console.log(
        `  ${good ? '\x1b[38;5;108m ok \x1b[0m' : '\x1b[1;31mFAIL\x1b[0m'} ${g.order[i].slice(-14)} ${arrow} ${g.order[j].slice(-14)}  ` +
          `expected ${want}, got ${c.relation} (shared ${c.shared.length}, onlyA ${c.onlyA.length}, onlyB ${c.onlyB.length})`,
      );
    }
  }
}
console.log(`\n${fail ? '\x1b[1;31m' : '\x1b[1;32m'}${pass} directed pairs correct, ${fail} wrong${skipped ? `, ${skipped} groups skipped` : ''}\x1b[0m\n`);
process.exit(fail ? 1 : 0);
