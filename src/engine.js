// engine.js — Palimpsest: reconstruct the editing history of a .docx from OOXML internals.
//
// Everything here is inference from bytes that Word itself wrote into the file:
//   * w:rsid* attributes   — "revision save ID": a random id minted per editing session.
//   * w14:paraId/textId    — stable per-paragraph ids (survive edits; enable version matching).
//   * docProps/*           — author names, edit minutes, save counts, producing app version.
//   * tracked changes      — dated, attributed, and often *not* accepted (deleted text is still there).
//   * a:srcRect crops      — Word keeps the whole image and just displays a window of it.
//
// No network, no dependencies, no upload.

import { readZip, parseXML, walk, findAll, find, child, children, textOf, DEFAULT_INFLATE } from './zipxml.js';
import { imageInfo, sniffMime } from './media.js';

const A = (n, k) => (n && n.attrs ? n.attrs[k] : undefined);
const RSID_ATTRS = ['w:rsidR', 'w:rsidRPr', 'w:rsidP', 'w:rsidDel', 'w:rsidRDefault', 'w:rsidSect', 'w:rsidTr'];

const APP_VERSIONS = {
  '11': 'Word 2003', '12': 'Word 2007', '14': 'Word 2010', '15': 'Word 2013',
  '16': 'Word 2016 / 2019 / 365',
};

/* ============================================================ helpers ==== */

// a fill that is present and not white: white glyphs on it are visible, not hidden
const isDark = (fill) => !!fill && fill !== 'auto' && !/^(FFFFFF|ffffff)$/.test(fill);
const pct = (a, b) => (b ? Math.round((a / b) * 1000) / 10 : 0);
function iso(s) {
  if (!s) return null;
  const d = new Date(s);
  return isNaN(+d) ? null : d.toISOString();
}
function minutesBetween(a, b) {
  const x = new Date(a), y = new Date(b);
  return isNaN(+x) || isNaN(+y) ? null : Math.round((y - x) / 60000);
}
function hhmm(s) {
  const d = new Date(s);
  return isNaN(+d) ? null : d.getUTCHours() + d.getUTCMinutes() / 60;
}
function bump(map, k, n = 1) {
  if (k === undefined || k === null || k === '') return;
  map.set(k, (map.get(k) || 0) + n);
}
const sortedCounts = (map) => [...map.entries()].sort((a, b) => b[1] - a[1]).map(([k, n]) => ({ name: k, n }));

/* ==================================================== document model ==== */
// Flatten the body into paragraphs of runs, keeping every provenance attribute
// we can see on the way down.

const TEXTY = new Set(['w:t', 'w:delText', 'w:instrText', 'w:delInstrText']);
// block-level wrappers we walk through to reach paragraphs, wherever they live
const DESCEND = new Set([
  'w:tr', 'w:tc', 'w:sdt', 'w:sdtContent', 'w:customXml', 'w:txbxContent',
  'mc:AlternateContent', 'mc:Choice', 'mc:Fallback',
  'w:glossaryDocument', 'w:docParts', 'w:docPart', 'w:docPartBody',
  'w:hdr', 'w:ftr', 'w:footnotes', 'w:footnote', 'w:endnotes', 'w:endnote',
  'w:comments', 'w:comment', 'w:ins', 'w:del',
]);
const CONTAINERS = new Set(['w:hyperlink', 'w:smartTag', 'w:sdtContent', 'w:sdt', 'w:bdo', 'w:dir', 'w:customXml', 'w:fldSimple']);

function runProps(rPr) {
  const p = {};
  if (!rPr) return p;
  const rf = child(rPr, 'w:rFonts');
  if (rf) {
    p.font = A(rf, 'w:ascii') || A(rf, 'w:hAnsi') || A(rf, 'w:cs') || null;
    p.fontEA = A(rf, 'w:eastAsia') || null;
    p.fontTheme = A(rf, 'w:asciiTheme') || A(rf, 'w:cstheme') || null;
  }
  const lang = child(rPr, 'w:lang');
  if (lang) {
    p.lang = A(lang, 'w:val') || null;
    p.langEA = A(lang, 'w:eastAsia') || null;
    p.langBidi = A(lang, 'w:bidi') || null;
  }
  for (const [tag, key] of [['w:b', 'b'], ['w:i', 'i'], ['w:u', 'u'], ['w:strike', 'strike'], ['w:vanish', 'vanish'], ['w:webHidden', 'webHidden'], ['w:caps', 'caps'], ['w:smallCaps', 'smallCaps']]) {
    const c = child(rPr, tag);
    if (c) p[key] = A(c, 'w:val') !== '0' && A(c, 'w:val') !== 'false';
  }
  const col = child(rPr, 'w:color');
  if (col) p.color = A(col, 'w:val');
  const hl = child(rPr, 'w:highlight');
  if (hl) p.highlight = A(hl, 'w:val');
  const shd = child(rPr, 'w:shd');
  if (shd) p.shading = A(shd, 'w:fill');
  const sz = child(rPr, 'w:sz');
  if (sz) p.sz = +A(sz, 'w:val') / 2;
  const style = child(rPr, 'w:rStyle');
  if (style) p.rStyle = A(style, 'w:val');
  return p;
}

function buildModel(doc, styleNames) {
  const paragraphs = [];
  const drawings = [];
  const fields = [];
  const bookmarks = [];
  let order = 0;

  const collectRuns = (node, ctx, out, depth = 0) => {
    if (depth > 60) return;
    for (const k of node.kids) {
      if (typeof k !== 'object') continue;
      const tag = k.tag;
      if (tag === 'w:r') {
        let text = '';
        let kind = 'text';
        for (const t of k.kids) {
          if (typeof t !== 'object') continue;
          if (TEXTY.has(t.tag)) {
            text += textOf(t);
            if (t.tag === 'w:delText') kind = 'delText';
            if (t.tag === 'w:instrText' || t.tag === 'w:delInstrText') kind = 'field';
          } else if (t.tag === 'w:tab') text += '\t';
          else if (t.tag === 'w:br' || t.tag === 'w:cr') text += '\n';
          else if (t.tag === 'w:noBreakHyphen') text += '-';
          else if (t.tag === 'w:softHyphen') text += '';
          else if (t.tag === 'w:sym') text += '\u25a1';
          else if (t.tag === 'w:drawing' || t.tag === 'w:pict' || t.tag === 'w:object') {
            const d = { paraIndex: paragraphs.length, rsid: A(k, 'w:rsidR') || ctx.paraRsid, node: t };
            drawings.push(d);
            text += '\u2b1c';
          }
        }
        const rp = runProps(child(k, 'w:rPr'));
        if (kind === 'field') fields.push({ instr: text.trim(), paraIndex: paragraphs.length });
        out.push({
          text,
          kind,
          rsid: A(k, 'w:rsidR') || ctx.paraRsidDefault || null,
          rsidRPr: A(k, 'w:rsidRPr') || null,
          rsidDel: A(k, 'w:rsidDel') || null,
          rev: ctx.rev || null,
          props: rp,
          order: order++,
        });
      } else if (tag === 'w:ins' || tag === 'w:del' || tag === 'w:moveFrom' || tag === 'w:moveTo') {
        const rev = {
          type: tag.slice(2),
          author: A(k, 'w:author') || null,
          date: A(k, 'w:date') || A(k, 'w16du:dateUtc') || null,
          id: A(k, 'w:id') || null,
        };
        // A removal context is sticky: text inside <w:moveFrom><w:ins> was
        // inserted and then moved away, so accepting the changes deletes it.
        // Letting the inner mark win would report it as surviving text.
        const outer = ctx.rev;
        const sticky = outer && (outer.type === 'del' || outer.type === 'moveFrom') ? outer : rev;
        collectRuns(k, { ...ctx, rev: sticky }, out, depth + 1);
      } else if (CONTAINERS.has(tag)) {
        collectRuns(k, ctx, out, depth + 1);
      } else if (tag === 'w:commentRangeStart' || tag === 'w:commentRangeEnd') {
        out.push({ text: '', kind: 'commentMark', commentId: A(k, 'w:id'), which: tag.endsWith('Start') ? 'start' : 'end', rsid: null, props: {}, order: order++ });
      } else if (tag === 'w:bookmarkStart') {
        bookmarks.push({ name: A(k, 'w:name'), paraIndex: paragraphs.length });
      } else if (tag === 'm:oMath' || tag === 'm:oMathPara') {
        out.push({ text: textOf(k), kind: 'math', rsid: ctx.paraRsidDefault || null, props: {}, order: order++ });
      }
    }
  };

  const visitBlocks = (node, tableDepth, depth = 0, bg = null) => {
    if (depth > 120) return; // pathological nesting (fuzzed or generated files)
    for (const k of node.kids) {
      if (typeof k !== 'object') continue;
      if (k.tag === 'w:p') {
        const pPr = child(k, 'w:pPr');
        const styleId = pPr ? A(child(pPr, 'w:pStyle'), 'w:val') : null;
        const ctx = {
          paraRsid: A(k, 'w:rsidR'),
          paraRsidDefault: A(k, 'w:rsidRDefault') || A(k, 'w:rsidR'),
        };
        const runs = [];
        collectRuns(k, ctx, runs);
        const numPr = pPr ? child(pPr, 'w:numPr') : null;
        const pMarkRPr = pPr ? runProps(child(pPr, 'w:rPr')) : {};
        const rPrChange = pPr ? find(pPr, 'w:pPrChange') : null;
        paragraphs.push({
          index: paragraphs.length,
          paraId: A(k, 'w14:paraId') || null,
          textId: A(k, 'w14:textId') || null,
          rsidR: A(k, 'w:rsidR') || null,
          rsidRPr: A(k, 'w:rsidRPr') || null,
          rsidP: A(k, 'w:rsidP') || null,
          rsidDel: A(k, 'w:rsidDel') || null,
          rsidRDefault: A(k, 'w:rsidRDefault') || null,
          style: styleId ? styleNames.get(styleId) || styleId : null,
          styleId: styleId || null,
          listLevel: numPr ? +(A(child(numPr, 'w:ilvl'), 'w:val') || 0) : null,
          tableDepth,
          markVanish: !!pMarkRPr.vanish,
          bgFill: (pPr ? A(child(pPr, 'w:shd'), 'w:fill') : null) || bg || null,
          pPrChange: rPrChange ? { author: A(rPrChange, 'w:author'), date: A(rPrChange, 'w:date') } : null,
          runs,
        });
      } else if (k.tag === 'w:tbl') {
        visitBlocks(k, tableDepth + 1, depth + 1, bg);
      } else if (DESCEND.has(k.tag)) {
        // a shaded table cell makes white text perfectly visible: remember the fill
        const cellShd = k.tag === 'w:tc' ? A(find(child(k, 'w:tcPr') || { kids: [] }, 'w:shd'), 'w:fill') : null;
        visitBlocks(k, tableDepth, depth + 1, cellShd && cellShd !== 'auto' && cellShd !== 'FFFFFF' ? cellShd : bg);
      }
    }
  };

  const body = find(doc, 'w:body') || doc;
  visitBlocks(body, 0);
  // text boxes live inside drawings; pick up their paragraphs too
  for (const tb of findAll(body, 'w:txbxContent')) visitBlocks(tb, 0);
  return { paragraphs, drawings, fields, bookmarks, body };
}

/* ========================================================== sessions ==== */

// Empirically (verified against a corpus of real documents): rsid values are
// uniform random 24-bit numbers, and settings.xml lists them sorted, so the
// values themselves carry NO chronology. What they do carry is *partition*:
// which text belongs to the same save. Two further facts do a lot of work:
//   * w:rsidRoot is by definition the session that created the file.
//   * an rsid that appears in the body but NOT in settings.xml's session list
//     was not minted by this document — it rode in with pasted content.
function sessionAnalysis(model, auxModels, settingsRsids, rootRsid) {
  const sessions = new Map(); // rsid -> stats
  const touch = (rsid, f) => {
    if (!rsid || rsid === '00000000') return null; // converters emit a null rsid
    let s = sessions.get(rsid);
    if (!s) {
      s = { rsid, chars: 0, runs: 0, paraMarks: 0, propEdits: 0, deletes: 0, paraIdx: [], langs: new Map(), fonts: new Map(), dates: [], authors: new Map(), native: settingsRsids.has(rsid) || rsid === rootRsid };
      sessions.set(rsid, s);
    }
    if (f) f(s);
    return s;
  };
  const edges = new Set(); // "a>b" == a happened no later than b
  const foreignParas = new Map(); // paragraph index -> foreign rsids seen there

  touch(rootRsid, (s) => { s.isRoot = true; });
  for (const r of settingsRsids) touch(r);

  const noteForeign = (rsid, paraIndex) => {
    if (!rsid || settingsRsids.has(rsid) || rsid === rootRsid) return;
    if (!foreignParas.has(paraIndex)) foreignParas.set(paraIndex, new Set());
    foreignParas.get(paraIndex).add(rsid);
  };

  const scan = (paragraphs, isBody) => {
    for (const p of paragraphs) {
      const pi = isBody ? p.index : -1;
      touch(p.rsidR, (s) => { s.paraMarks++; if (isBody && !s.paraIdx.includes(p.index)) s.paraIdx.push(p.index); });
      // only sound constraint: you cannot delete something before it exists
      if (p.rsidDel && p.rsidR && p.rsidDel !== p.rsidR) edges.add(p.rsidR + '>' + p.rsidDel);
      touch(p.rsidP, (s) => s.propEdits++);
      touch(p.rsidRPr, (s) => s.propEdits++);
      touch(p.rsidDel, (s) => s.deletes++);
      for (const a of [p.rsidRPr, p.rsidP]) noteForeign(a, pi);
      for (const r of p.runs) {
        const n = (r.text || '').length;
        const s = touch(r.rsid, (s) => {
          s.runs++;
          s.chars += n;
          if (isBody && !s.paraIdx.includes(p.index)) s.paraIdx.push(p.index);
          if (r.props.lang) bump(s.langs, r.props.lang);
          if (r.props.font) bump(s.fonts, r.props.font, n || 1);
        });
        if (r.rev && s) {
          if (r.rev.date) s.dates.push(r.rev.date);
          if (r.rev.author) bump(s.authors, r.rev.author);
        }
        if (r.rsidDel && r.rsid && r.rsidDel !== r.rsid) edges.add(r.rsid + '>' + r.rsidDel);
        touch(r.rsidRPr, (s) => s.propEdits++);
        touch(r.rsidDel, (s) => s.deletes++);
        noteForeign(r.rsidRPr, pi);
        if (n) noteForeign(r.rsid, pi);
      }
    }
  };
  scan(model.paragraphs, true);
  for (const a of auxModels) scan(a.model.paragraphs, false);

  // the creating session precedes every other session of this document
  if (rootRsid) for (const [k, s] of sessions) if (k !== rootRsid && s.native) edges.add(rootRsid + '>' + k);

  const list = [...sessions.values()];
  const totalChars = list.reduce((a, s) => a + s.chars, 0);
  for (const s of list) {
    s.share = pct(s.chars, totalChars);
    s.langs = sortedCounts(s.langs);
    s.fonts = sortedCounts(s.fonts);
    s.authors = sortedCounts(s.authors);
    s.dates = [...new Set(s.dates)].sort();
    s.span = s.paraIdx.length ? [Math.min(...s.paraIdx), Math.max(...s.paraIdx)] : null;
    s.contiguous = s.span ? s.span[1] - s.span[0] + 1 === s.paraIdx.length : false;
  }
  const order = layerSessions(list.filter((s) => s.native).map((s) => s.rsid), edges);
  for (const s of list) s.layer = s.native ? order.layer.get(s.rsid) ?? null : null;
  list.sort((a, b) => (a.native === b.native ? 0 : a.native ? -1 : 1) || (a.layer ?? 99) - (b.layer ?? 99) || b.chars - a.chars);
  return {
    sessions: list,
    totalChars,
    edges: [...edges],
    layers: order.layers,
    cycles: order.cycles,
    foreign: list.filter((s) => !s.native),
    foreignParas: [...foreignParas.entries()].map(([paraIndex, set]) => ({ paraIndex, rsids: [...set] })).filter((x) => x.paraIndex >= 0),
  };
}

// Longest-path layering over the "no later than" DAG (cycles are collapsed:
// mutually-constraining rsids cannot be separated in time by this evidence).
function layerSessions(nodes, edgeSet) {
  const adj = new Map(nodes.map((n) => [n, []]));
  const indeg = new Map(nodes.map((n) => [n, 0]));
  for (const e of edgeSet) {
    const [a, b] = e.split('>');
    if (!adj.has(a) || !adj.has(b)) continue;
    adj.get(a).push(b);
    indeg.set(b, indeg.get(b) + 1);
  }
  const layer = new Map();
  let q = nodes.filter((n) => indeg.get(n) === 0);
  for (const n of q) layer.set(n, 0);
  let seen = 0;
  while (q.length) {
    const next = [];
    for (const n of q) {
      seen++;
      for (const m of adj.get(n)) {
        layer.set(m, Math.max(layer.get(m) ?? 0, (layer.get(n) ?? 0) + 1));
        indeg.set(m, indeg.get(m) - 1);
        if (indeg.get(m) === 0) next.push(m);
      }
    }
    q = next;
  }
  const cycles = seen < nodes.length;
  if (cycles) for (const n of nodes) if (!layer.has(n)) layer.set(n, 1);
  const layers = new Map();
  for (const [n, l] of layer) {
    if (!layers.has(l)) layers.set(l, []);
    layers.get(l).push(n);
  }
  return { layer, layers: [...layers.entries()].sort((a, b) => (a[0] ?? 99) - (b[0] ?? 99)).map(([l, ns]) => ({ layer: l, rsids: ns })), cycles };
}

/* ============================================================ analyze ==== */

export async function analyze(bytes, opts = {}) {
  const inflate = opts.inflate || DEFAULT_INFLATE;
  const zip = await readZip(bytes, inflate);
  const parts = zip.order.map((n) => {
    const e = zip.entries.get(n);
    return { name: n, size: e.size, compressedSize: e.compressedSize, method: e.method, mtime: e.mtime, crc: e.crc };
  });
  const rep = {
    file: { name: opts.name || 'document.docx', bytes: bytes.length },
    container: { parts, zipComment: zip.comment, partCount: parts.length },
    findings: [],
    identities: [],
    media: [],
    embedded: [],
    hiddenText: [],
    deletedText: [],
    comments: [],
    fields: [],
    external: [],
    customProps: [],
    labels: [],
  };
  const T = async (n) => (await zip.text(n)) || '';
  const has = (n) => zip.has(n);
  const add = (level, title, detail, extra = {}) => rep.findings.push({ level, title, detail, ...extra });

  if (!has('word/document.xml')) throw new Error('not a Word .docx (word/document.xml missing)');

  /* ---- properties ------------------------------------------------------ */
  const coreXml = await T('docProps/core.xml');
  const appXml = await T('docProps/app.xml');
  const customXml = await T('docProps/custom.xml');
  const core = parseXML(coreXml);
  const app = parseXML(appXml);
  const gx = (tree, tag) => {
    const n = find(tree, tag);
    return n ? textOf(n).trim() : null;
  };
  const props = {
    title: gx(core, 'dc:title'),
    subject: gx(core, 'dc:subject'),
    creator: gx(core, 'dc:creator'),
    keywords: gx(core, 'cp:keywords'),
    description: gx(core, 'dc:description'),
    lastModifiedBy: gx(core, 'cp:lastModifiedBy'),
    revision: gx(core, 'cp:revision'),
    created: iso(gx(core, 'dcterms:created')),
    modified: iso(gx(core, 'dcterms:modified')),
    lastPrinted: iso(gx(core, 'cp:lastPrinted')),
    category: gx(core, 'cp:category'),
    contentStatus: gx(core, 'cp:contentStatus'),
    template: gx(app, 'Template'),
    totalTime: gx(app, 'TotalTime') ? +gx(app, 'TotalTime') : null,
    pages: gx(app, 'Pages') ? +gx(app, 'Pages') : null,
    words: gx(app, 'Words') ? +gx(app, 'Words') : null,
    characters: gx(app, 'Characters') ? +gx(app, 'Characters') : null,
    paragraphsCount: gx(app, 'Paragraphs') ? +gx(app, 'Paragraphs') : null,
    lines: gx(app, 'Lines') ? +gx(app, 'Lines') : null,
    application: gx(app, 'Application'),
    appVersion: gx(app, 'AppVersion'),
    company: gx(app, 'Company'),
    manager: gx(app, 'Manager'),
    docSecurity: gx(app, 'DocSecurity'),
    hyperlinkBase: gx(app, 'HyperlinkBase'),
  };
  rep.props = props;
  if (props.appVersion) {
    const major = props.appVersion.split('.')[0].replace(/^0+/, '');
    props.appVersionLabel = APP_VERSIONS[major] || null;
  }
  for (const p of findAll(parseXML(customXml), 'property')) {
    const name = A(p, 'name');
    const val = textOf(p).trim();
    if (!name) continue;
    rep.customProps.push({ name, value: val.slice(0, 400) });
  }

  /* ---- styles / settings ---------------------------------------------- */
  const styleNames = new Map();
  const stylesXml = await T('word/styles.xml');
  for (const s of findAll(parseXML(stylesXml), 'w:style')) {
    const id = A(s, 'w:styleId');
    const nm = A(child(s, 'w:name'), 'w:val');
    if (id) styleNames.set(id, nm || id);
  }
  const settingsXml = await T('word/settings.xml');
  const settings = parseXML(settingsXml);
  const settingsRsids = new Set();
  const rsidsNode = find(settings, 'w:rsids');
  if (rsidsNode) for (const r of children(rsidsNode, 'w:rsid')) settingsRsids.add(A(r, 'w:val'));
  const rootRsid = A(find(settings, 'w:rsidRoot'), 'w:val') || null;
  rep.settings = {
    rsidRoot: rootRsid,
    rsidCount: settingsRsids.size,
    trackChangesOn: !!find(settings, 'w:trackChanges'),
    documentProtection: (() => {
      const dp = find(settings, 'w:documentProtection');
      return dp ? { edit: A(dp, 'w:edit'), enforced: A(dp, 'w:enforcement'), hasHash: !!A(dp, 'w:hash'), algo: A(dp, 'w:cryptAlgorithmSid') || A(dp, 'w:cryptProviderType') } : null;
    })(),
    proofState: (() => {
      const ps = find(settings, 'w:proofState');
      return ps ? { spelling: A(ps, 'w:spelling'), grammar: A(ps, 'w:grammar') } : null;
    })(),
    themeFontLang: (() => {
      const t = find(settings, 'w:themeFontLang');
      return t ? [A(t, 'w:val'), A(t, 'w:eastAsia'), A(t, 'w:bidi')].filter(Boolean).join(' / ') : null;
    })(),
    removePersonalInformation: !!find(settings, 'w:removePersonalInformation'),
    attachedTemplateId: A(find(settings, 'w:attachedTemplate'), 'r:id') || null,
    defaultTabStop: A(find(settings, 'w:defaultTabStop'), 'w:val') || null,
    zoom: A(find(settings, 'w:zoom'), 'w:percent') || null,
  };

  /* ---- relationships (external targets, template path) ---------------- */
  const relsFor = async (partName) => {
    const rp = partName.replace(/([^/]+)$/, '_rels/$1.rels');
    const x = await T(rp);
    const out = new Map();
    if (!x) return out;
    for (const r of findAll(parseXML(x), 'Relationship')) {
      out.set(A(r, 'Id'), { type: (A(r, 'Type') || '').split('/').pop(), target: A(r, 'Target'), mode: A(r, 'TargetMode') || 'Internal' });
    }
    return out;
  };
  const docRels = await relsFor('word/document.xml');
  const settingsRels = await relsFor('word/settings.xml');
  for (const [id, r] of [...docRels, ...settingsRels]) {
    if (r.mode === 'External' || /^(file:|\\\\|[A-Za-z]:\\)/.test(r.target || '')) {
      rep.external.push({ id, type: r.type, target: r.target, local: /^(file:|\\\\|[A-Za-z]:\\)/.test(r.target || '') });
    }
  }
  if (rep.settings.attachedTemplateId) {
    const t = settingsRels.get(rep.settings.attachedTemplateId);
    if (t) rep.settings.attachedTemplate = t.target;
  }

  /* ---- document model ------------------------------------------------- */
  const docXml = await T('word/document.xml');
  const doc = parseXML(docXml);
  const model = buildModel(doc, styleNames);
  rep.text = { paragraphs: model.paragraphs.length };

  // headers/footers/footnotes also carry rsids and text worth mining
  const auxParts = zip.order.filter((n) => /^word\/(header|footer|footnotes|endnotes|comments|glossary)/.test(n) && n.endsWith('.xml'));
  const auxModels = [];
  for (const p of auxParts) {
    const x = await T(p);
    if (!x) continue;
    auxModels.push({ part: p, model: buildModel(parseXML(x), styleNames) });
  }

  // Physical counts straight off the XML text. The semantic model above can
  // legitimately see nothing (a converter that writes every rsid as 00000000,
  // an insertion that only touches a paragraph mark) while the bookkeeping is
  // still sitting in the file — and a remover has to work from what is there.
  const raw = { rsidAttrs: 0, paraIds: 0, ins: 0, del: 0, revMarks: 0, vanish: 0, lang: 0, parts: 0 };
  for (const n of ['word/document.xml', ...auxParts, 'word/styles.xml', 'word/numbering.xml', 'word/settings.xml']) {
    const x = n === 'word/document.xml' ? docXml : await T(n);
    if (!x) continue;
    raw.parts++;
    raw.rsidAttrs += (x.match(/\sw:rsid[A-Za-z]*="/g) || []).length;
    raw.paraIds += (x.match(/\sw1?4?:?(?:paraId|textId)="/g) || []).length;
    raw.ins += (x.match(/<w:(?:ins|moveTo)\b/g) || []).length;
    raw.del += (x.match(/<w:(?:del|moveFrom)\b/g) || []).length;
    raw.revMarks += (x.match(/<w:(?:pPrChange|rPrChange|tblPrChange|trPrChange|tcPrChange|sectPrChange|cellIns|cellDel|cellMerge)\b/g) || []).length;
    raw.vanish += (x.match(/<w:(?:vanish|webHidden)\b/g) || []).length;
    raw.lang += (x.match(/<w:lang\b/g) || []).length;
  }
  rep.raw = raw;

  /* ---- parts nothing declares or points at ----------------------------- */
  // Every legitimate part of a .docx is declared in [Content_Types].xml and
  // reachable from a relationship. Leftovers — Word's own [trash]/ dumps,
  // parts orphaned by a converter — are content that no reader will ever show.
  {
    const ct = parseXML(await T('[Content_Types].xml'));
    const declared = new Set(findAll(ct, 'Override').map((o) => (A(o, 'PartName') || '').replace(/^\//, '')));
    const exts = new Set(findAll(ct, 'Default').map((d) => (A(d, 'Extension') || '').toLowerCase()));
    const referenced = new Set(['[Content_Types].xml', '_rels/.rels']);
    for (const n of zip.order) {
      if (!/_rels\/[^/]*\.rels$/.test(n)) continue;
      referenced.add(n);
      const dir = n.replace(/_rels\/[^/]*$/, '');
      for (const r of findAll(parseXML(await T(n)), 'Relationship')) {
        const t = A(r, 'Target') || '';
        if (A(r, 'TargetMode') === 'External' || /^[a-z]+:/i.test(t) || t.startsWith('//') || t.startsWith('\\\\')) continue;
        const segs = [];
        for (const s of (t.startsWith('/') ? t.slice(1) : dir + t.replace(/^\.\//, '')).split('/')) {
          if (s === '..') segs.pop();
          else if (s && s !== '.') segs.push(s);
        }
        referenced.add(segs.join('/'));
      }
    }
    // Deliberately conservative: a part is only junk when *nothing* declares a
    // type for it and *nothing* relates to it. Either signal alone is too weak
    // — plenty of real documents forget to relate settings.xml, and a remover
    // that deletes core parts on a technicality is worse than useless.
    const CORE = /^(docProps\/|word\/(document|settings|styles|stylesWithEffects|numbering|fontTable|webSettings|theme\/|footnotes|endnotes|header\d*|footer\d*)\.?)/;
    rep.orphanParts = zip.order
      .filter((n) => !n.endsWith('/') && n !== '[Content_Types].xml' && !/^word\/media\//.test(n) && !/(^|\/)_rels\//.test(n) && !CORE.test(n))
      .map((n) => ({ part: n, size: (zip.entries.get(n) || {}).size || 0, typed: declared.has(n) || exts.has((n.split('.').pop() || '').toLowerCase()), referenced: referenced.has(n) }))
      .filter((o) => !o.typed && !o.referenced);
  }

  const S = sessionAnalysis(model, auxModels, settingsRsids, rootRsid);
  rep.sessions = S.sessions;
  rep.metrics = compositionMetrics(S);
  rep.sessionOrder = { layers: S.layers, edges: S.edges.length, ambiguous: S.cycles, foreignParas: S.foreignParas };

  /* ---- identities ------------------------------------------------------ */
  const idents = new Map();
  const noteIdent = (name, role) => {
    if (!name || !name.trim()) return;
    const k = name.trim();
    if (!idents.has(k)) idents.set(k, new Set());
    idents.get(k).add(role);
  };
  noteIdent(props.creator, 'created the file');
  noteIdent(props.lastModifiedBy, 'saved it last');
  noteIdent(props.manager, 'listed as manager');

  /* ---- tracked changes + hidden + deleted ----------------------------- */
  const revDates = [];
  let insertions = 0;
  const scanForRevisions = (m, partName) => {
    for (const p of m.paragraphs) {
      if (p.pPrChange) {
        noteIdent(p.pPrChange.author, 'reformatted a paragraph');
        if (p.pPrChange.date) revDates.push(p.pPrChange.date);
      }
      for (const r of p.runs) {
        if (r.rev) {
          if (r.rev.author) noteIdent(r.rev.author, r.rev.type === 'del' ? 'deleted text (tracked)' : 'inserted text (tracked)');
          if (r.rev.date) revDates.push(r.rev.date);
          if (r.rev.type === 'ins' || r.rev.type === 'moveTo') insertions++;
          if ((r.rev.type === 'del' || r.rev.type === 'moveFrom') && r.text.trim()) {
            rep.deletedText.push({ text: r.text, author: r.rev.author, date: r.rev.date, paraIndex: p.index, part: partName, type: r.rev.type });
          }
        }
        // A run whose only "text" is the placeholder for a picture has no
        // words to hide: white type around an image is not invisible text.
        const visibleText = r.text.replace(/\u2b1c/g, '').trim();
        if (r.props.vanish && visibleText) rep.hiddenText.push({ text: r.text, kind: 'hidden (w:vanish)', paraIndex: p.index, part: partName });
        else if (
          r.props.color && /^(FFFFFF|ffffff)$/.test(r.props.color) && visibleText &&
          r.props.highlight === undefined && isDark(r.props.shading) === false && isDark(p.bgFill) === false
        ) {
          rep.hiddenText.push({ text: r.text, kind: 'white text on white', paraIndex: p.index, part: partName });
        } else if (r.props.sz !== undefined && r.props.sz <= 1.5 && visibleText) {
          rep.hiddenText.push({ text: r.text, kind: `${r.props.sz}pt text`, paraIndex: p.index, part: partName });
        }
      }
    }
    for (const f of m.fields) if (f.instr) rep.fields.push({ instr: f.instr.slice(0, 300), part: partName });
  };
  scanForRevisions(model, 'word/document.xml');
  for (const a of auxModels) scanForRevisions(a.model, a.part);
  rep.trackedInsertions = insertions;
  rep.paraIdCount =
    model.paragraphs.filter((p) => p.paraId).length + auxModels.reduce((a, x) => a + x.model.paragraphs.filter((p) => p.paraId).length, 0);

  /* ---- comments -------------------------------------------------------- */
  if (has('word/comments.xml')) {
    const cx = parseXML(await T('word/comments.xml'));
    const exDone = new Map();
    if (has('word/commentsExtended.xml')) {
      const ex = parseXML(await T('word/commentsExtended.xml'));
      for (const e of findAll(ex, 'w15:commentEx')) exDone.set(A(e, 'w15:paraId'), A(e, 'w15:done') === '1');
    }
    for (const c of findAll(cx, 'w:comment')) {
      const paraIds = findAll(c, 'w:p').map((p) => A(p, 'w14:paraId'));
      const author = A(c, 'w:author');
      noteIdent(author, 'left a comment');
      const date = A(c, 'w:date');
      if (date) revDates.push(date);
      rep.comments.push({
        id: A(c, 'w:id'),
        author,
        initials: A(c, 'w:initials'),
        date,
        text: textOf(c).trim().slice(0, 1200),
        resolved: paraIds.some((p) => exDone.get(p)),
      });
    }
  }
  if (has('word/people.xml')) {
    const px = parseXML(await T('word/people.xml'));
    for (const p of findAll(px, 'w15:person')) {
      const nm = A(p, 'w15:author');
      const pres = child(p, 'w15:presenceInfo');
      const uid = A(pres, 'w15:userId') || '';
      const provider = A(pres, 'w15:providerId') || '';
      const email = (uid.match(/[\w.+-]+@[\w.-]+\.\w+/) || [])[0] || null;
      noteIdent(nm, 'co-authored via ' + (provider || 'Word'));
      if (email) noteIdent(email, 'account id in people.xml');
      rep.identityIds = rep.identityIds || [];
      rep.identityIds.push({ author: nm, provider, userId: uid, email });
    }
  }

  /* ---- sensitivity labels & DMS traces -------------------------------- */
  for (const p of rep.customProps) {
    if (/^MSIP_Label_/.test(p.name)) {
      const m = p.name.match(/^MSIP_Label_([0-9a-fA-F-]+)_(\w+)$/);
      if (m) rep.labels.push({ guid: m[1], field: m[2], value: p.value });
    }
  }
  for (const n of zip.order) {
    if (/^customXml\/item\d*\.xml$/.test(n)) {
      const x = await T(n);
      const tags = [...new Set([...x.matchAll(/<(?:\w+:)?(\w{3,40})[ >\/]/g)].map((m) => m[1]))].slice(0, 12);
      const urls = [...new Set([...x.matchAll(/https?:\/\/[^\s"'<>]{6,160}/g)].map((m) => m[0]))].slice(0, 6);
      rep.embedded.push({ part: n, kind: 'customXml', tags, urls, size: x.length, preview: x.replace(/\s+/g, ' ').slice(0, 400) });
    } else if (/^word\/embeddings\//.test(n)) {
      const b = await zip.bytes(n);
      rep.embedded.push({ part: n, kind: 'embedded object', size: b ? b.length : 0 });
    } else if (n === 'word/vbaProject.bin') {
      const b = await zip.bytes(n);
      rep.embedded.push({ part: n, kind: 'VBA macro project', size: b ? b.length : 0 });
    } else if (/^word\/glossary\//.test(n) && n.endsWith('document.xml')) {
      const gm = buildModel(parseXML(await T(n)), styleNames);
      const txt = gm.paragraphs.map((p) => p.runs.map((r) => r.text).join('')).filter((s) => s.trim());
      if (txt.length) rep.embedded.push({ part: n, kind: 'AutoText / building blocks', size: txt.join('').length, preview: txt.join(' \u00b7 ').slice(0, 600) });
    }
  }

  /* ---- images: crop excavation + metadata ----------------------------- */
  const mediaNames = zip.order.filter((n) => /^(word\/media\/|docProps\/thumbnail)/.test(n));
  const cropsByRel = new Map();
  for (const d of model.drawings) {
    const blip = find(d.node, 'a:blip');
    const embed = A(blip, 'r:embed') || A(blip, 'r:link');
    const sr = find(d.node, 'a:srcRect');
    const ext = find(d.node, 'wp:extent') || find(d.node, 'a:ext');
    let crop = sr ? { l: +(A(sr, 'l') || 0) / 1000, t: +(A(sr, 't') || 0) / 1000, r: +(A(sr, 'r') || 0) / 1000, b: +(A(sr, 'b') || 0) / 1000 } : null;
    if (crop && !(crop.l || crop.t || crop.r || crop.b)) crop = null; // <a:srcRect/> with no sides = no crop
    const descr = find(d.node, 'wp:docPr');
    if (!embed) continue;
    const rel = docRels.get(embed);
    if (!rel) continue;
    const target = 'word/' + rel.target.replace(/^\.?\//, '');
    const prev = cropsByRel.get(target) || [];
    prev.push({
      crop,
      paraIndex: d.paraIndex,
      rsid: d.rsid,
      displayEmu: ext ? { cx: +(A(ext, 'cx') || 0), cy: +(A(ext, 'cy') || 0) } : null,
      name: A(descr, 'name') || null,
      descr: A(descr, 'descr') || null,
    });
    cropsByRel.set(target, prev);
  }
  for (const n of mediaNames) {
    const b = await zip.bytes(n);
    if (!b) continue;
    const info = imageInfo(b);
    const uses = cropsByRel.get(n) || [];
    const worst = uses.reduce((acc, u) => {
      if (!u.crop) return acc;
      const vis = (1 - (u.crop.l + u.crop.r) / 100) * (1 - (u.crop.t + u.crop.b) / 100);
      return acc === null || vis < acc.vis ? { vis, crop: u.crop } : acc;
    }, null);
    rep.media.push({
      part: n,
      size: b.length,
      format: info.format,
      width: info.width,
      height: info.height,
      meta: info.meta,
      mime: sniffMime(b),
      uses: uses.map((u) => ({ ...u, displayPx: u.displayEmu ? { w: Math.round(u.displayEmu.cx / 9525), h: Math.round(u.displayEmu.cy / 9525) } : null })),
      cropped: !!worst,
      visibleFraction: worst ? Math.round(worst.vis * 1000) / 10 : 100,
      crop: worst ? worst.crop : null,
      unused: uses.length === 0 && !n.startsWith('docProps/'),
      bytes: opts.keepBytes ? b : undefined,
    });
  }

  /* ---- locale / font islands ------------------------------------------ */
  const langCounts = new Map();
  const fontCounts = new Map();
  for (const p of model.paragraphs) {
    for (const r of p.runs) {
      const n = (r.text || '').length;
      if (!n) continue;
      if (r.props.lang) bump(langCounts, r.props.lang, n);
      if (r.props.font) bump(fontCounts, r.props.font, n);
    }
  }
  rep.locales = sortedCounts(langCounts);
  rep.fonts = sortedCounts(fontCounts);
  const fontTable = parseXML(await T('word/fontTable.xml'));
  rep.fontTable = findAll(fontTable, 'w:font').map((f) => A(f, 'w:name')).filter(Boolean);

  /* ---- plain text + counts -------------------------------------------- */
  const plain = model.paragraphs
    .map((p) => p.runs.filter((r) => r.kind !== 'delText' && r.kind !== 'field').map((r) => r.text).join(''))
    .join('\n');
  rep.text.chars = plain.length;
  rep.text.words = (plain.match(/[\p{L}\p{N}][\p{L}\p{N}'’\-]*/gu) || []).length;
  rep.text.preview = plain.slice(0, 400);
  rep._plain = plain; // full visible text, for the scrubber's preservation check (stripped from JSON)

  /* ---- timing ---------------------------------------------------------- */
  const allDates = [...revDates.map(iso).filter(Boolean)].sort();
  rep.timeline = {
    created: props.created,
    modified: props.modified,
    lastPrinted: props.lastPrinted,
    revisions: props.revision ? +props.revision : null,
    totalEditMinutes: props.totalTime,
    wallClockMinutes: props.created && props.modified ? minutesBetween(props.created, props.modified) : null,
    revisionDates: allDates,
    zipMtimes: [...new Set(parts.map((p) => p.mtime).filter(Boolean))].sort(),
  };
  const hours = allDates.map(hhmm).filter((h) => h !== null);
  if (hours.length) {
    rep.timeline.hourHistogram = Array.from({ length: 24 }, (_, h) => hours.filter((x) => Math.floor(x) === h).length);
  }

  // anchor sessions to real timestamps where tracked changes name them
  for (const s of rep.sessions) if (s.dates.length) s.anchoredAt = s.dates[0];
  const maxLayer = Math.max(...rep.sessions.map((s) => s.layer ?? -1));
  const lastOnes = rep.sessions.filter((s) => s.layer === maxLayer);
  if (lastOnes.length === 1 && props.modified && !lastOnes[0].anchoredAt) {
    lastOnes[0].anchoredAt = props.modified;
    lastOnes[0].anchorNote = 'last session in the partial order → the save recorded in docProps';
  }
  const rootSession = rep.sessions.find((s) => s.isRoot);
  if (rootSession && props.created && !rootSession.anchoredAt) {
    rootSession.anchoredAt = props.created;
    rootSession.anchorNote = 'rsidRoot → the session that created the file';
  }

  rep.identities = [...idents.entries()].map(([name, roles]) => ({ name, roles: [...roles] }));

  buildFindings(rep, model, S, add);
  rep.verdict = verdict(rep);
  rep._model = model; // kept for the UI (stripped from JSON output by callers)
  return rep;
}

// How many editing sessions is the visible text *effectively* spread over?
// exp(Shannon entropy of the per-session character shares): 1.0 means one save
// produced everything, 8.0 means the text is genuinely spread over ~8 sittings.
function compositionMetrics(S) {
  const withText = S.sessions.filter((s) => s.native && s.chars > 0);
  const total = withText.reduce((a, s) => a + s.chars, 0);
  if (!total) return { effectiveSessions: 0, entropy: 0, topShare: 0, sessionsWithText: 0 };
  let H = 0;
  for (const s of withText) {
    const p = s.chars / total;
    if (p > 0) H -= p * Math.log(p);
  }
  return {
    effectiveSessions: Math.round(Math.exp(H) * 100) / 100,
    entropy: Math.round((withText.length > 1 ? H / Math.log(withText.length) : 0) * 1000) / 1000,
    topShare: Math.max(...withText.map((s) => s.share)),
    sessionsWithText: withText.length,
  };
}

/* =========================================================== findings ==== */

function buildFindings(rep, model, S, add) {
  const p = rep.props;
  // Word's own last-save count is the honest denominator; fall back to ours
  const words = Math.max(rep.text.words || 0, p.words || 0);

  // --- has this file already been through a scrubber? ------------------
  // Sparse metadata can be consistent with scrubbing, disabled recording,
  // or a producer/converter that never emitted these fields. This flag is
  // a heuristic, not evidence that a particular cleaning action occurred.
  const scrubbed =
    words > 20 &&
    !rep.settings.rsidCount &&
    rep.sessions.length === 0 &&
    rep.identities.length === 0 &&
    !p.totalTime &&
    (rep.settings.removePersonalInformation || !p.revision || p.revision === '1');
  rep.scrubbed = scrubbed;
  if (scrubbed) {
    add('info', 'Little editing metadata remains', `No revision-save ids, no authorship and no positive editing time. This is consistent with cleaning, disabled recording, or a producer that omitted these fields${rep.settings.removePersonalInformation ? '. A personal-information removal setting is enabled' : ''}.`);
  }

  // --- how the text got here ------------------------------------------
  if (!scrubbed && p.totalTime !== null && p.totalTime !== undefined && words > 150) {
    const wpm = p.totalTime > 0 ? words / p.totalTime : Infinity;
    if (wpm > 120) {
      add(
        'key',
        p.totalTime === 0 ? 'Zero minutes of editing recorded' : `${Math.round(wpm)} words per minute of recorded editing`,
        p.totalTime === 0
          ? `The file records 0 editing minutes and ${words.toLocaleString()} words. Pasting, generation, conversion, or a reset or incomplete timer can produce this combination; it does not establish how the text was written.`
          : `${words.toLocaleString()} words against ${p.totalTime} minute${p.totalTime === 1 ? '' : 's'} of recorded editing time. This ratio can suggest imported text, but the recorded timer may be reset or incomplete and is not a typing-speed measurement.`,
        { metric: { words, minutes: p.totalTime, wpm: p.totalTime ? Math.round(wpm) : null } },
      );
    } else if (p.totalTime > 5) {
      add('info', `${p.totalTime} minutes of editing recorded`, `Word accumulated ${p.totalTime} minutes with this file open across ${p.revision || '?'} save${p.revision === '1' ? '' : 's'} — consistent with ${words.toLocaleString()} words being written here (${Math.round(wpm)} words per editing minute).`);
    }
  }

  // --- sessions -------------------------------------------------------
  const visible = rep.sessions.filter((s) => s.chars > 0 && s.native);
  const ghosts = rep.sessions.filter((s) => s.chars === 0 && s.native);
  if (rep.settings.rsidCount || visible.length) {
    add('info', `${rep.settings.rsidCount || visible.length} editing sessions recorded`,
      `settings.xml lists ${rep.settings.rsidCount} revision-save IDs; ${visible.length} of them tag visible text. ${ghosts.length ? `${ghosts.length} listed ID${ghosts.length === 1 ? '' : 's'} have no associated visible text in the analyzed content. They may relate to overwritten text, formatting, or other changes.` : ''}`);
  }
  if (rep.metrics.sessionsWithText >= 4 && rep.metrics.effectiveSessions < 1.6 && words > 200) {
    add('key', `The text effectively comes from ${rep.metrics.effectiveSessions} of its ${rep.settings.rsidCount || rep.metrics.sessionsWithText} sessions`,
      `Character volume per session gives an effective count of ${rep.metrics.effectiveSessions} (exp of the share entropy). Documents written incrementally spread their text over many sessions; this one has a long tail of tiny touch-ups on top of a single large deposit.`);
  }
  // rsids in the body that this document never minted: they came in with the text
  const foreign = rep.sessions.filter((s) => !s.native);
  if (foreign.length >= 3 && rep.settings.rsidCount) {
    const paras = S.foreignParas.length;
    add('key', `${foreign.length} revision ids in the body are absent from the session list`,
      `The file lists ${rep.settings.rsidCount} revision-save IDs, while ${paras} paragraph${paras === 1 ? '' : 's'} carry ${foreign.length} additional IDs. Copied content can retain IDs, but an incomplete or rewritten session list can also explain this. An ID alone does not identify a source document.`,
      { foreign: foreign.slice(0, 40).map((s) => s.rsid) });
  }
  if (visible.length) {
    const top = visible[0].share >= visible[visible.length - 1].share ? [...visible].sort((a, b) => b.share - a.share)[0] : visible[0];
    if (top.share >= 85 && words > 150) {
      add('key', `${top.share}% of the text shares one revision ID`,
        `Revision ${top.rsid} tags ${top.chars.toLocaleString()} of ${S.totalChars.toLocaleString()} characters. This concentration is consistent with bulk insertion or editing; it does not distinguish pasting from other ways of producing the text.`);
    }
  }
  // paste islands: a session that owns one contiguous block and nothing else
  const islands = visible.filter((s) => s.contiguous && s.paraIdx.length <= 6 && s.chars > 80 && s.share < 60);
  if (islands.length) {
    add('key', `${islands.length} block${islands.length === 1 ? '' : 's'} of text has a distinct revision ID`,
      `These contiguous paragraph runs share a revision ID not found elsewhere in the analyzed text: ${islands.slice(0, 6).map((s) => `¶${s.span[0] + 1}${s.span[1] !== s.span[0] ? `–${s.span[1] + 1}` : ''}`).join(', ')}. This is a grouping clue, not proof of pasting or a separate writing session.`,
      { islands: islands.map((s) => ({ rsid: s.rsid, span: s.span, chars: s.chars })) });
  }

  // --- recoverable content --------------------------------------------
  if (rep.deletedText.length) {
    const chars = rep.deletedText.reduce((a, d) => a + d.text.length, 0);
    add('key', `${chars.toLocaleString()} characters of deleted text are still in the file`,
      `${rep.deletedText.length} tracked deletion${rep.deletedText.length === 1 ? '' : 's'} were never accepted, so Word kept the original wording inside the document. Anyone who opens the file can read what was cut.`);
  }
  if (rep.hiddenText.length) {
    const kinds = [...new Set(rep.hiddenText.map((h) => h.kind))];
    add('key', `${rep.hiddenText.length} ${rep.hiddenText.length === 1 ? 'passage is' : 'passages are'} invisible on screen`,
      `Text present in the file but formatted not to be seen (${kinds.join(', ')}). This is how keyword stuffing, stale drafts and leftover instructions survive a "final" version.`);
  }
  if (rep.comments.length) {
    const open = rep.comments.filter((c) => !c.resolved).length;
    add(rep.comments.length ? 'key' : 'info', `${rep.comments.length} comment${rep.comments.length === 1 ? '' : 's'} shipped with the document`,
      `${open} unresolved. Comments carry author names and timestamps, and often the internal discussion the document was meant to hide.`);
  }

  // --- images ---------------------------------------------------------
  const cropped = rep.media.filter((m) => m.cropped && m.visibleFraction < 99);
  if (cropped.length) {
    const worst = cropped.reduce((a, b) => (a.visibleFraction < b.visibleFraction ? a : b));
    add('key', `${cropped.length} cropped image${cropped.length === 1 ? '' : 's'} — and the cropped-away parts are still here`,
      `Word stores the whole picture and only displays a window onto it. The tightest crop shows ${worst.visibleFraction}% of ${worst.part.split('/').pop()} — the other ${Math.round(100 - worst.visibleFraction)}% is recoverable from this file.`);
  }
  const withMeta = rep.media.filter((m) => Object.keys(m.meta || {}).length);
  if (withMeta.length) {
    const gps = withMeta.find((m) => m.meta.GPS);
    const cam = withMeta.find((m) => m.meta.Make || m.meta.Model);
    if (gps) add('key', 'An embedded image carries GPS coordinates', `${gps.part.split('/').pop()} still has its EXIF location: ${gps.meta.GPS}.`);
    else if (cam) add('info', 'Camera metadata survives inside an embedded image', `${cam.part.split('/').pop()}: ${[cam.meta.Make, cam.meta.Model, cam.meta.DateTimeOriginal].filter(Boolean).join(' · ')}`);
    else add('info', `${withMeta.length} embedded image${withMeta.length === 1 ? ' carries' : 's carry'} their own metadata`, withMeta.slice(0, 3).map((m) => `${m.part.split('/').pop()}: ${Object.entries(m.meta).slice(0, 3).map(([k, v]) => `${k}=${v}`).join(', ')}`).join(' — '));
  }
  const unused = rep.media.filter((m) => m.unused);
  if (unused.length) add('info', `${unused.length} image${unused.length === 1 ? ' is' : 's are'} in the file but displayed nowhere`, `Left over from earlier drafts: ${unused.map((m) => m.part.split('/').pop()).slice(0, 5).join(', ')}.`);
  if (rep.media.some((m) => m.part.startsWith('docProps/thumbnail'))) {
    add('info', 'The file carries a page-1 thumbnail', 'The preview image is rendered at save time and can show an earlier state of the first page than the text you read.');
  }

  // --- identity -------------------------------------------------------
  if (rep.identities.length) {
    add(rep.identities.length > 1 ? 'key' : 'info', `${rep.identities.length} name${rep.identities.length === 1 ? '' : 's'} attached to this file`,
      rep.identities.map((i) => `${i.name} (${i.roles.join('; ')})`).join(' — '));
  }
  if (rep.identityIds && rep.identityIds.some((i) => i.email)) {
    add('key', 'Account identifiers (emails) are embedded in people.xml',
      rep.identityIds.filter((i) => i.email).map((i) => `${i.author} → ${i.email}`).join(', '));
  }
  if (rep.settings.removePersonalInformation) add('info', 'Word was asked to strip personal info on save', 'The removePersonalInformation flag is set, which is why authorship may be blank while session structure survives.');
  if (p.company) add('info', 'Organisation recorded in the file', `Company: ${p.company}`);
  if (p.template && p.template !== 'Normal.dotm' && p.template !== 'Normal') add('info', 'Built from a named template', `Template: ${p.template}${rep.settings.attachedTemplate ? ` (attached: ${rep.settings.attachedTemplate})` : ''}`);
  else if (rep.settings.attachedTemplate) add('info', 'Attached template path', rep.settings.attachedTemplate);
  if (rep.labels.length) {
    const byGuid = new Map();
    for (const l of rep.labels) {
      if (!byGuid.has(l.guid)) byGuid.set(l.guid, {});
      byGuid.get(l.guid)[l.field] = l.value;
    }
    const desc = [...byGuid.entries()].map(([g, f]) => `${f.Name || g}${f.Owner ? ` (applied by ${f.Owner}` : ''}${f.SetDate ? ` on ${f.SetDate.slice(0, 10)})` : f.Owner ? ')' : ''}`).join('; ');
    add('key', 'Microsoft sensitivity label attached', `${desc}. Labels record who classified the document and when — and the label survives even after the text changes.`);
  }
  const localRefs = rep.external.filter((e) => e.local);
  const webRefs = rep.external.filter((e) => !e.local);
  if (localRefs.length) {
    add('key', `${localRefs.length} reference${localRefs.length === 1 ? '' : 's'} to a local or network path`,
      `${localRefs.slice(0, 6).map((e) => `${e.type}: ${e.target}`).join(' — ')}. Paths like these expose drive letters, server names and folder structure from the author's machine.`);
  }
  if (webRefs.length > 2) {
    add('info', `${webRefs.length} external links`, `Ordinary hyperlinks, listed in the metadata panel. Worth a glance for staging or intranet hosts: ${[...new Set(webRefs.map((e) => { try { return new URL(e.target).host; } catch { return e.target.slice(0, 30); } }))].slice(0, 6).join(', ')}`);
  }
  const dmsHints = rep.embedded.filter((e) => e.kind === 'customXml' && (e.urls.length || e.tags.some((t) => /contenttype|documentmanagement|sharepoint|properties/i.test(t))));
  if (dmsHints.length) {
    add('info', 'Traces of a document-management system', `customXml parts describe the library this file lived in: ${dmsHints.flatMap((d) => d.urls).slice(0, 3).join(', ') || dmsHints[0].tags.join(', ')}`);
  }
  if (rep.embedded.some((e) => e.kind === 'VBA macro project')) add('key', 'The file contains a VBA macro project', 'Macros are code; treat this document as executable.');
  if ((rep.orphanParts || []).length) {
    const trash = rep.orphanParts.filter((o) => /^\[trash\]|\/trash/i.test(o.part));
    const bytes = rep.orphanParts.reduce((a, o) => a + o.size, 0);
    add(trash.length ? 'key' : 'info',
      trash.length ? `${trash.length} part${trash.length === 1 ? '' : 's'} of this file live in Word's [trash] folder` : `${rep.orphanParts.length} part${rep.orphanParts.length === 1 ? '' : 's'} nothing in the document points at`,
      `${rep.orphanParts.map((o) => `${o.part} (${(o.size / 1024).toFixed(1)} KB${o.typed ? '' : ', not even declared in [Content_Types].xml'})`).join(', ')}. ${(bytes / 1024).toFixed(0)} KB that no reader will render and no Word command will show you.`);
  }

  const glossary = rep.embedded.find((e) => e.kind === 'AutoText / building blocks');
  if (glossary) add('key', 'AutoText entries are stored inside the document', `Saved building blocks travel with the file and often contain unrelated boilerplate from the author\'s machine: “${glossary.preview.slice(0, 160)}…”`);
  const embObjs = rep.embedded.filter((e) => e.kind === 'embedded object');
  if (embObjs.length) add('key', `${embObjs.length} embedded object${embObjs.length === 1 ? ' (a whole file) is' : 's (whole files) are'} inside this document`, `${embObjs.map((e) => `${e.part.split('/').pop()} (${Math.round(e.size / 1024)} KB)`).join(', ')}. Embedded spreadsheets keep every cell, not just the chart you can see.`);

  // --- time anomalies --------------------------------------------------
  if (p.created && p.modified) {
    const drift = minutesBetween(p.created, p.modified);
    if (drift < 0) add('key', 'The file was saved before it was created', `created=${p.created}, modified=${p.modified}. Either the clock was wrong or the metadata was edited.`);
    else if (drift === 0 && words > 400) add('key', 'Created and last saved in the same minute', `${words.toLocaleString()} words appeared and the file was finished inside one minute (${p.created}).`);
  }
  if (p.lastPrinted) add('info', 'The document was printed', `Last printed ${p.lastPrinted}${p.modified ? ` — ${p.lastPrinted > p.modified ? 'after' : 'before'} the last save.` : '.'}`);
  if (rep.timeline.revisionDates.length > 1) {
    const f = rep.timeline.revisionDates[0], l = rep.timeline.revisionDates.at(-1);
    add('info', 'Tracked changes give real timestamps', `Dated edits run ${f.slice(0, 16).replace('T', ' ')} → ${l.slice(0, 16).replace('T', ' ')} (${rep.timeline.revisionDates.length} dated events), which pins named editing sessions to the calendar.`);
  }
  if (rep.timeline.zipMtimes.length > 1) {
    add('info', 'The parts inside the zip have different timestamps', `${rep.timeline.zipMtimes.length} distinct part timestamps (${rep.timeline.zipMtimes[0]} → ${rep.timeline.zipMtimes.at(-1)}) — a sign the container was assembled or rewritten by a tool, not saved in one go by Word.`);
  }

  // --- locale ----------------------------------------------------------
  if (rep.locales.length > 1) {
    const total = rep.locales.reduce((a, l) => a + l.n, 0);
    const minor = rep.locales.slice(1).filter((l) => l.n / total > 0.002);
    if (minor.length) {
      add('key', `Text carries ${rep.locales.length} different language settings`,
        `Dominant: ${rep.locales[0].name} (${pct(rep.locales[0].n, total)}%). Also present: ${minor.map((l) => `${l.name} (${pct(l.n, total)}%)`).join(', ')}. Language tags may reflect proofing choices, templates, copied formatting, or input settings; they do not establish the writer or location.`);
    }
  }
  if (rep.fonts.length > 3) {
    const total = rep.fonts.reduce((a, f) => a + f.n, 0);
    add('info', `${rep.fonts.length} explicit fonts in the body text`, rep.fonts.slice(0, 6).map((f) => `${f.name} ${pct(f.n, total)}%`).join(', ') + '. Font islands often mark pasted passages.');
  }

  // --- producer --------------------------------------------------------
  if (!rep.settings.rsidCount && !rep.sessions.some((s) => s.chars)) {
    if (!scrubbed) add('info', 'No session ids in this file', `No revision-save IDs were found. Recording may have been disabled, stripped, or omitted by producer “${p.application || 'unknown'}”. Metadata, tracked changes and image analysis still apply.`);
  }
  if (p.application && !/Microsoft/i.test(p.application)) {
    add('info', `Written by ${p.application}`, 'Not Microsoft Word — the file was produced or rewritten by another tool, which usually strips Word\'s own history.');
  }
  if (rep.settings.documentProtection) {
    const dp = rep.settings.documentProtection;
    add('info', `Editing restrictions are enabled (${dp.edit || 'unknown mode'})`, `Enforced with a ${dp.hasHash ? 'password hash stored in settings.xml' : 'flag'} — this is a UI lock, not encryption; deleting one XML element removes it.`);
  }
  if (rep.settings.trackChangesOn) add('info', 'Track changes is currently switched on', 'Whoever opens this next will have their edits recorded and attributed.');
  if (rep.fields.length) {
    const interesting = rep.fields.filter((f) => /MERGEFIELD|INCLUDE|LINK|DDE|HYPERLINK|DOCPROPERTY|FILENAME|AUTHOR|USERNAME|MACROBUTTON/i.test(f.instr));
    if (interesting.length) add('info', `${interesting.length} field code${interesting.length === 1 ? '' : 's'} reference outside data`, interesting.slice(0, 5).map((f) => f.instr.slice(0, 120)).join(' | '));
  }
  rep.findings.sort((a, b) => (a.level === b.level ? 0 : a.level === 'key' ? -1 : 1));
}

function verdict(rep) {
  const p = rep.props;
  const visible = rep.sessions.filter((s) => s.chars > 0 && s.native);
  const words = Math.max(rep.text.words || 0, rep.props.words || 0);
  const lines = [];
  const sessionCount = rep.settings.rsidCount || visible.length;
  if (sessionCount > 1) {
    lines.push(`${sessionCount} editing sessions are recorded inside this file${visible.length && visible.length !== sessionCount ? `, ${visible.length} of which still ${visible.length === 1 ? 'owns' : 'own'} text you can see` : ''}.`);
  }
  const top = [...visible].sort((a, b) => b.share - a.share)[0];
  if (top && top.share >= 85 && words > 150) lines.push(`One revision ID tags ${top.share}% of the text; this is a concentration of recorded metadata, not proof of how the text was produced.`);
  else if (top && visible.length > 3) lines.push(`The text is spread across sessions (largest owns ${top.share}%), which is what incremental writing looks like.`);
  if (rep.scrubbed) lines.push('No save IDs, authorship, or positive editing clock were found. Cleaning or incomplete metadata could explain their absence.');
  else if (p.totalTime === 0 && words > 150) lines.push(`Word recorded zero editing minutes for ${words.toLocaleString()} words.`);
  else if (p.totalTime) lines.push(`${p.totalTime} minute${p.totalTime === 1 ? '' : 's'} of editing across ${p.revision || '?'} save${p.revision === '1' ? '' : 's'}.`);
  if (rep.deletedText.length) lines.push(`${rep.deletedText.length} deleted passage${rep.deletedText.length === 1 ? ' is' : 's are'} still readable.`);
  if (rep.hiddenText.length) lines.push(`${rep.hiddenText.length} passage${rep.hiddenText.length === 1 ? ' is' : 's are'} invisible on screen but present in the file.`);
  const names = rep.identities.map((i) => i.name);
  if (names.length) lines.push(`Names in the file: ${names.slice(0, 4).join(', ')}${names.length > 4 ? ` +${names.length - 4} more` : ''}.`);
  const crop = rep.media.filter((m) => m.cropped && m.visibleFraction < 99);
  if (crop.length) lines.push(`${crop.length} image${crop.length === 1 ? '' : 's'} hide cropped-out content.`);
  return {
    headline: headline(rep, top, words, sessionCount),
    lines,
    score: { key: rep.findings.filter((f) => f.level === 'key').length, info: rep.findings.filter((f) => f.level === 'info').length },
  };
}

function headline(rep, top, words, sessionCount) {
  const p = rep.props;
  const foreign = rep.sessions.filter((s) => !s.native).length;
  const wpm = p.totalTime ? words / p.totalTime : null;
  if (rep.deletedText.length) return 'Text that was deleted is still inside this file';
  if (rep.hiddenText.length > 2) return `${rep.hiddenText.length} passages are hidden in plain sight`;
  // recoverable content outranks this: a stripped history is a fact about the
  // file, but text someone can still read is a fact about its contents
  if (rep.scrubbed) return 'This document has sparse editing metadata';
  if (p.totalTime === 0 && words > 300) return 'This document records zero editing minutes';
  if (foreign >= 10) return `${foreign} revision IDs are absent from the session list`;
  if (rep.media.some((m) => m.cropped && m.visibleFraction < 90)) return 'A cropped image is hiding recoverable content';
  if (top && top.share >= 90 && words > 300) {
    return wpm !== null && wpm < 120 && p.totalTime >= 10
      ? `One dominant revision ID; ${p.totalTime} editing minutes recorded`
      : 'Most text shares one revision ID';
  }
  if (rep.comments.length) return `${rep.comments.length} internal comment${rep.comments.length === 1 ? '' : 's'} shipped with this file`;
  if (sessionCount >= 15) return `Assembled over ${sessionCount} separate editing sessions`;
  if (rep.identities.length > 2) return `${rep.identities.length} people left their names in this file`;
  if (!sessionCount) return 'No Word session history — but the metadata still talks';
  return `${sessionCount} editing session${sessionCount === 1 ? '' : 's'} recorded`;
}

/* ============================================================ lineage ==== */

/**
 * Compare two analysed documents. Two independent signals:
 *  - the set of editing-session ids (a file's session list is inherited by
 *    every Save As, so shared ids prove shared ancestry);
 *  - w14:paraId, a per-paragraph id that survives rewriting, which tells you
 *    which paragraphs are *the same paragraph* even after the words changed.
 */
export function compareDocs(a, b, bname = 'the second file') {
  const aN = new Set(a.sessions.filter((s) => s.native).map((s) => s.rsid));
  const bN = new Set(b.sessions.filter((s) => s.native).map((s) => s.rsid));
  const shared = [...aN].filter((x) => bN.has(x));
  const onlyA = [...aN].filter((x) => bN.has(x) === false);
  const onlyB = [...bN].filter((x) => aN.has(x) === false);
  const sameRoot = !!a.settings.rsidRoot && a.settings.rsidRoot === b.settings.rsidRoot;
  // Word occasionally prunes a session id whose text no longer exists, so
  // "later" cannot require perfect inclusion. Direction = near-containment plus
  // the fact that a save can only add ids, never invent old ones.
  const inAinB = aN.size ? shared.length / aN.size : 0;
  const inBinA = bN.size ? shared.length / bN.size : 0;
  const overlap = shared.length / Math.max(1, aN.size + bN.size - shared.length);
  const dropNote = (n) => (n ? ` (${n} id${n === 1 ? '' : 's'} from the earlier file were pruned along the way)` : '');
  let verdict;
  let relation;
  if (inAinB >= 0.97 && bN.size > aN.size) {
    relation = 'descendant';
    verdict = `${bname} is the later file: it contains ${shared.length} of this document's ${aN.size} editing sessions and adds ${onlyB.length} more${dropNote(onlyA.length)}.`;
  } else if (inBinA >= 0.97 && aN.size > bN.size) {
    relation = 'ancestor';
    verdict = `${bname} is the earlier file: the open document contains its ${shared.length} sessions and adds ${onlyA.length} more${dropNote(onlyB.length)}.`;
  } else if (overlap > 0.35 || (sameRoot && shared.length)) {
    relation = 'forked';
    verdict = `Same lineage, but the two files forked: ${onlyA.length} session${onlyA.length === 1 ? '' : 's'} exist only in the open file, ${onlyB.length} only in ${bname}, on top of ${shared.length} shared. Two copies were edited in parallel.`;
  } else if (shared.length) {
    relation = 'shared-text';
    verdict = `${shared.length} editing session${shared.length === 1 ? '' : 's'} in common out of ${aN.size} and ${bN.size} — text was copied between these files, but they are not versions of one document.`;
  } else {
    relation = 'unrelated';
    verdict = 'No shared editing sessions: these documents have unrelated histories.';
  }

  const pa = new Map(), pb = new Map();
  for (const p of (a._model ? a._model.paragraphs : [])) if (p.paraId) pa.set(p.paraId, p.runs.map((r) => r.text).join('').trim());
  for (const p of (b._model ? b._model.paragraphs : [])) if (p.paraId) pb.set(p.paraId, p.runs.map((r) => r.text).join('').trim());
  let paragraphs = null;
  if (pa.size && pb.size) {
    const kept = [...pa.keys()].filter((x) => pb.has(x));
    paragraphs = {
      survived: kept.length,
      rewritten: kept.filter((x) => pa.get(x) !== pb.get(x) && (pa.get(x) || pb.get(x))).map((x) => ({ paraId: x, before: pa.get(x), after: pb.get(x) })),
      removed: [...pa.keys()].filter((x) => pb.has(x) === false && pa.get(x)).map((x) => ({ paraId: x, text: pa.get(x) })),
      added: [...pb.keys()].filter((x) => pa.has(x) === false && pb.get(x)).map((x) => ({ paraId: x, text: pb.get(x) })),
    };
  }
  return { relation, verdict, shared, onlyA, onlyB, sameRoot, paragraphs };
}

/* ============================================================= export ==== */

/** Plain-text evidence report, for pasting into a note or a story. */
export function textReport(rep) {
  const L = [];
  L.push(`PALIMPSEST REPORT — ${rep.file.name}`);
  L.push(`${(rep.file.bytes / 1024).toFixed(0)} KB · ${rep.container.partCount} parts · ${rep.props.application || 'unknown producer'}${rep.props.appVersionLabel ? ` (${rep.props.appVersionLabel})` : ''}`);
  L.push('');
  L.push(rep.verdict.headline.toUpperCase());
  for (const l of rep.verdict.lines) L.push('  · ' + l);
  L.push('');
  L.push('FINDINGS');
  for (const f of rep.findings) L.push(`${f.level === 'key' ? '[!]' : '[i]'} ${f.title}\n    ${f.detail}`);
  const S = rep.sessions.filter((s) => s.chars > 0);
  if (S.length) {
    L.push('');
    L.push('EDITING SESSIONS (rsid · chars · share · paragraphs)');
    for (const s of S) L.push(`  ${s.rsid}${s.isRoot ? ' [created the file]' : s.native ? '' : ' [foreign]'}  ${s.chars} · ${s.share}% · ¶${s.span ? s.span[0] + 1 + '-' + (s.span[1] + 1) : '-'}${s.anchoredAt ? ' · ' + s.anchoredAt : ''}`);
  }
  if (rep.deletedText.length) {
    L.push('');
    L.push('DELETED BUT STILL PRESENT');
    for (const d of rep.deletedText) L.push(`  [${d.author || '?'} ${d.date || ''}] ${JSON.stringify(d.text)}`);
  }
  if (rep.hiddenText.length) {
    L.push('');
    L.push('INVISIBLE ON SCREEN');
    for (const d of rep.hiddenText) L.push(`  [${d.kind}] ${JSON.stringify(d.text)}`);
  }
  if (rep.comments.length) {
    L.push('');
    L.push('COMMENTS');
    for (const c of rep.comments) L.push(`  [${c.author || '?'} ${c.date || ''}${c.resolved ? ' resolved' : ''}] ${JSON.stringify(c.text)}`);
  }
  if (rep.media.length) {
    L.push('');
    L.push('IMAGES');
    for (const m of rep.media) L.push(`  ${m.part} ${m.format} ${m.width}x${m.height}${m.cropped ? ` cropped→${m.visibleFraction}% visible` : ''}${Object.keys(m.meta || {}).length ? ' ' + JSON.stringify(m.meta) : ''}`);
  }
  L.push('');
  L.push('Generated locally by Palimpsest. Every line above is read from bytes inside the file.');
  return L.join('\n');
}
