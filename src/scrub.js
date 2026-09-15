// scrub.js — the other half of Palimpsest: remove what the analyzer found and
// write the document back out. Same engine, same zero dependencies; the crop
// baker is the only part that needs a host callback (a canvas, in the browser).
//
// Everything here is string surgery on the OOXML parts rather than a tree
// rewrite: it preserves byte-for-byte everything it is not deliberately
// touching, which is exactly the property you want in a redaction tool.

import { readZip, DEFAULT_INFLATE } from './zipxml.js';
import { writeZip, DEFAULT_DEFLATE } from './zipwrite.js';

const DEC = new TextDecoder('utf-8');
const ENC = new TextEncoder();

/* ------------------------------------------------- xml string surgery ---- */

// Index of the '>' that closes the tag starting at i, respecting quoted attrs.
function tagEnd(s, i) {
  let q = 0;
  for (let j = i + 1; j < s.length; j++) {
    const c = s[j];
    if (q) {
      if (c === q) q = 0;
      continue;
    }
    if (c === '"' || c === "'") {
      q = c;
      continue;
    }
    if (c === '>') return j;
  }
  return -1;
}

function isNameEnd(c) {
  return c === undefined || c === ' ' || c === '\t' || c === '\n' || c === '\r' || c === '>' || c === '/';
}

/**
 * Cut or unwrap every `<tag>` element in an XML string.
 * @param {string} xml
 * @param {string} tag
 * @param {'cut'|'unwrap'} mode
 * @param {(inner: string, start: string) => boolean} [pred] inner text + start tag
 * @returns {{out: string, n: number}}
 */
export function editElements(xml, tag, mode, pred) {
  const open = '<' + tag;
  const close = '</' + tag;
  // `<w:del` must not match `<w:delText`, and `</w:del` must not match
  // `</w:delText` — the whole element name has to end where the needle does.
  const seek = (needle, from) => {
    let at = xml.indexOf(needle, from);
    while (at >= 0 && !isNameEnd(xml[at + needle.length])) at = xml.indexOf(needle, at + 1);
    return at;
  };
  let out = '';
  let pos = 0;
  let n = 0;
  for (;;) {
    const i = seek(open, pos);
    if (i < 0) break;
    const se = tagEnd(xml, i);
    if (se < 0) break;
    const selfClosing = xml[se - 1] === '/';
    const startTag = xml.slice(i, se + 1);
    let inner = '';
    let end = se + 1; // index just past the element
    if (!selfClosing) {
      let depth = 1;
      let k = se + 1;
      let broke = false;
      while (depth > 0) {
        const o = seek(open, k);
        const c = seek(close, k);
        if (c < 0) {
          broke = true;
          break;
        }
        if (o >= 0 && o < c) {
          const oe = tagEnd(xml, o);
          if (oe < 0) {
            broke = true;
            break;
          }
          if (xml[oe - 1] !== '/') depth++;
          k = oe + 1;
        } else {
          depth--;
          k = xml.indexOf('>', c) + 1;
          if (depth === 0) {
            inner = xml.slice(se + 1, c);
            end = k;
          }
        }
      }
      if (broke) break; // malformed; leave the remainder untouched
    }
    if (!pred || pred(inner, startTag)) {
      out += xml.slice(pos, i);
      if (mode === 'unwrap') out += inner;
      n++;
      pos = end;
    } else {
      // Keep this element but carry on scanning *inside* it: a w:r that holds a
      // drawing holds the runs of the text box too, and those are separate
      // candidates. Advancing past the whole element would hide them.
      out += xml.slice(pos, se + 1);
      pos = se + 1;
    }
  }
  return { out: out + xml.slice(pos), n };
}

export function stripAttrs(xml, re) {
  let n = 0;
  const out = xml.replace(re, () => {
    n++;
    return '';
  });
  return { out, n };
}

const xmlText = (frag) => frag.replace(/<[^>]*>/g, '');

function setElementText(xml, tag, value) {
  const re = new RegExp(`<${tag}([^>]*)>[\\s\\S]*?</${tag}>|<${tag}([^>]*)/>`, 'g');
  let found = false;
  const out = xml.replace(re, (m, a1, a2) => {
    found = true;
    const attrs = (a1 || a2 || '').replace(/\s*\/$/, '');
    return value === null ? '' : `<${tag}${attrs}>${value}</${tag}>`;
  });
  return { out, found };
}

/* ------------------------------------------------------------ helpers ---- */

const IS_XML = (n) => /\.(xml|rels)$/i.test(n);
const WORD_PARTS = (n) =>
  /^word\/(document|header\d*|footer\d*|footnotes|endnotes|comments|commentsExtended|commentsIds|commentsExtensible|styles|numbering|settings|glossary\/.*)\.xml$/.test(n);

function relDir(relsName) {
  // word/_rels/document.xml.rels -> word/
  return relsName.replace(/_rels\/[^/]*$/, '');
}

function resolveTarget(relsName, target) {
  if (/^[a-z]+:/i.test(target) || target.startsWith('//') || target.startsWith('\\\\')) return null; // external
  // a leading slash means package-root-relative, not "inside this folder"
  let path = target.startsWith('/') ? target.slice(1) : relDir(relsName) + target.replace(/^\.\//, '');
  const segs = [];
  for (const s of path.split('/')) {
    if (s === '..') segs.pop();
    else if (s !== '.' && s !== '') segs.push(s);
  }
  return segs.join('/');
}

/* --------------------------------------------------------- image bytes --- */

// JPEG: drop APP1..APP15 (EXIF, XMP, IPTC, ICC-in-APP2 is a tradeoff we take
// for APP2 only when it is not an ICC profile) and COM comments.
function stripJpeg(b) {
  if (b[0] !== 0xff || b[1] !== 0xd8) return null;
  const keep = [0xff, 0xd8];
  let i = 2;
  let dropped = 0;
  const emit = () => {
    // whatever we stopped at (entropy-coded scan data, or a truncated tail)
    // travels verbatim: this function only ever removes whole segments.
    const rest = b.subarray(i);
    const out = new Uint8Array(keep.length + rest.length);
    out.set(keep, 0);
    out.set(rest, keep.length);
    return dropped ? { bytes: out, dropped } : null;
  };
  while (i + 3 < b.length) {
    if (b[i] !== 0xff) break;
    const marker = b[i + 1];
    if (marker === 0xd8 || (marker >= 0xd0 && marker <= 0xd9) || marker === 0x01) {
      keep.push(b[i], marker);
      i += 2;
      continue;
    }
    const len = (b[i + 2] << 8) | b[i + 3];
    if (len < 2 || i + 2 + len > b.length) break;
    const isApp = marker >= 0xe1 && marker <= 0xef;
    const isIcc = marker === 0xe2 && DEC.decode(b.subarray(i + 4, i + 15)).startsWith('ICC_PROFILE');
    const isAdobe = marker === 0xee; // colour transform, needed for some CMYK
    const drop = (isApp && !isIcc && !isAdobe) || marker === 0xfe;
    if (drop) dropped += 2 + len;
    else for (let k = i; k < i + 2 + len; k++) keep.push(b[k]);
    i += 2 + len;
    if (marker === 0xda) return emit(); // start of scan
  }
  return emit();
}

const PNG_KEEP = new Set(['IHDR', 'PLTE', 'IDAT', 'IEND', 'tRNS', 'gAMA', 'cHRM', 'sRGB', 'iCCP', 'bKGD', 'pHYs', 'sBIT', 'hIST']);

function stripPng(b) {
  const sig = [137, 80, 78, 71, 13, 10, 26, 10];
  for (let i = 0; i < 8; i++) if (b[i] !== sig[i]) return null;
  const out = [b.subarray(0, 8)];
  let p = 8;
  let dropped = 0;
  const dv = new DataView(b.buffer, b.byteOffset, b.byteLength);
  while (p + 8 <= b.length) {
    const len = dv.getUint32(p);
    const type = String.fromCharCode(b[p + 4], b[p + 5], b[p + 6], b[p + 7]);
    const total = 12 + len;
    if (p + total > b.length) break;
    if (PNG_KEEP.has(type)) out.push(b.subarray(p, p + total));
    else dropped += total;
    p += total;
    if (type === 'IEND') break;
  }
  if (!dropped) return null;
  const size = out.reduce((a, c) => a + c.length, 0);
  const res = new Uint8Array(size);
  let q = 0;
  for (const c of out) {
    res.set(c, q);
    q += c.length;
  }
  return { bytes: res, dropped };
}

function stripImageMeta(bytes) {
  return stripJpeg(bytes) || stripPng(bytes);
}

/* --------------------------------------------------------------- plan ---- */

const yes = (n) => n > 0;

/**
 * What can be removed from this document, given its report.
 * @returns {Array<{id,label,detail,count,unit,on,risk}>}
 */
export function plan(rep) {
  const p = rep.props || {};
  const s = rep.sessions || [];
  const fixes = [];
  const push = (id, label, detail, count, unit, on, risk = 'safe') => {
    if (count) fixes.push({ id, label, detail, count, unit, on, risk });
  };

  const raw = rep.raw || {};
  // counts come off the XML text, not the analysis: a file can carry rsids the
  // session model ignores (every value 00000000) and revision records that
  // produced no run (an inserted paragraph mark), and those still have to go
  push('sessions', 'Remove editing-session ids', 'Deletes every w:rsid* stamp plus the settings.xml session list — the whole save history.', raw.rsidAttrs || s.length, 'stamps', true);
  push('paraIds', 'Remove stable paragraph ids', 'w14:paraId survives rewording and lets anyone match your paragraphs against another copy of the document.', raw.paraIds || rep.paraIdCount || 0, 'ids', true);
  const idents = (rep.identities || []).length;
  push('identity', 'Remove names, employer and template path', 'dc:creator, cp:lastModifiedBy, Company, Manager, HyperlinkBase, the attached-template path and word/people.xml.', idents || (p.company ? 1 : 0), 'identities', true);
  push('dates', 'Remove the creation and last-saved timestamps', 'dcterms:created and dcterms:modified — the pair that shows a report was written at 3am, or in the minute before it was sent.', (p.created ? 1 : 0) + (p.modified ? 1 : 0), 'timestamps', true);
  push('locales', 'Remove run language tags', 'w:lang on individual runs records the keyboard of whoever typed them; a document with three of them was assembled from three places. Removing them changes no pixel — only which dictionary Word spell-checks with.', raw.lang || 0, 'tags', true);
  push('timing', 'Reset the editing clock', `Zeroes TotalTime${p.totalTime ? ` (${p.totalTime} min)` : ''}, revision count${p.revision ? ` (${p.revision})` : ''} and the last-printed date.`, (p.totalTime ? 1 : 0) + (p.revision ? 1 : 0) + (p.lastPrinted ? 1 : 0), 'fields', true);
  push('tracked', 'Accept all tracked changes', 'Insertions become ordinary text; deleted text is actually deleted. Formatting-revision records go too.', (raw.ins || 0) + (raw.del || 0) + (raw.revMarks || 0) || (rep.deletedText || []).length, 'revisions', true, 'content');
  push('comments', 'Delete comments', 'Removes comments.xml, the extended/ids parts and the in-text anchors.', (rep.comments || []).length, 'comments', true, 'content');
  push('invisible', 'Delete invisible text', 'Runs that are hidden, white-on-white or 1pt — text a reader never sees but a parser does.', (rep.hiddenText || []).length, 'runs', true, 'content');
  const orphans = rep.orphanParts || [];
  push('orphans', 'Delete parts nothing points at', 'Leftover streams — including Word’s own [trash] folder — that travel with the file but are never rendered.', orphans.length, 'parts', true, 'content');
  const metaImgs = (rep.media || []).filter((m) => m.meta && Object.keys(m.meta).length);
  push('imgmeta', 'Strip image metadata', 'EXIF, XMP and comment blocks inside the pictures: camera, software, GPS, original filename.', metaImgs.length, 'images', true);
  const cropped = (rep.media || []).filter((m) => m.cropped);
  push('crops', 'Bake image crops', 'Re-encodes each cropped picture so the off-camera pixels stop travelling with the file.', cropped.length, 'images', true, 'reencode');
  const thumb = (rep.media || []).filter((m) => /^docProps\/thumbnail/.test(m.part));
  push('thumbnail', 'Delete the saved page thumbnail', 'A rendered picture of the document as it looked at some earlier save — it can show text that has since been edited away.', thumb.length, 'images', true);
  const unused = (rep.media || []).filter((m) => m.unused);
  push('unused', 'Delete images that are never displayed', 'Media parts nothing in the document refers to — usually deleted screenshots.', unused.length, 'images', true, 'content');
  const gloss = (rep.embedded || []).filter((e) => e.kind === 'AutoText / building blocks');
  push('glossary', 'Delete AutoText / building blocks', 'The glossary part carries saved boilerplate from the author’s machine.', gloss.length, 'parts', true, 'content');
  const cx = (rep.embedded || []).filter((e) => e.kind === 'customXml');
  push('customxml', 'Delete customXml parts', 'Sensitivity labels, document-management ids and whatever else a server stapled on.', cx.length + (rep.labels || []).length, 'parts', true);
  const local = (rep.external || []).filter((e) => e.local);
  push('refs', 'Remove local and network path references', 'file:// and \\\\server\\share targets that name your machine and drive layout.', local.length, 'links', true);
  const embeds = (rep.embedded || []).filter((e) => e.kind === 'embedded object');
  push('embeds', 'Delete embedded files', 'Whole workbooks and documents stored inside this one. Removing them removes the objects that display them.', embeds.length, 'files', false, 'content');
  const vba = (rep.embedded || []).filter((e) => e.kind === 'VBA macro project');
  push('macros', 'Delete the macro project', 'vbaProject.bin. The document stops being able to run its macros.', vba.length, 'projects', false, 'content');
  fixes.push({ id: 'future', label: 'Tell Word to stop recording', detail: 'Writes removePersonalInformation and removeDateAndTime into settings.xml so the next author’s saves stay clean too.', count: 1, unit: 'flags', on: true, risk: 'safe' });
  return fixes;
}

/* -------------------------------------------------------------- scrub ---- */

/**
 * @param {Uint8Array} bytes original .docx
 * @param {object} rep report from analyze()
 * @param {object} opts { inflate, deflate, fixes: string[]|null, bakeImage }
 * @returns {Promise<{bytes: Uint8Array, applied: Array, notes: string[]}>}
 */
export async function scrub(bytes, rep, opts = {}) {
  const inflate = opts.inflate || DEFAULT_INFLATE;
  const deflate = opts.deflate === undefined ? DEFAULT_DEFLATE : opts.deflate;
  const zip = await readZip(bytes, inflate);
  const want = opts.fixes ? new Set(opts.fixes) : new Set(plan(rep).filter((f) => f.on).map((f) => f.id));
  const notes = [];
  const counts = new Map();
  const did = (id, n = 1) => counts.set(id, (counts.get(id) || 0) + n);

  /** @type {Map<string, {bytes?: Uint8Array, text?: string}>} */
  const files = new Map();
  for (const name of zip.order) {
    if (name.endsWith('/')) continue;
    const b = await zip.bytes(name);
    files.set(name, IS_XML(name) ? { text: DEC.decode(b) } : { bytes: b });
  }
  const dead = new Set();
  const killPart = (name) => {
    if (files.has(name)) {
      files.delete(name);
      dead.add(name);
      return true;
    }
    return false;
  };
  const xmlParts = () => [...files.keys()].filter((n) => files.get(n).text !== undefined);
  const editXml = (name, fn) => {
    const f = files.get(name);
    if (!f || f.text === undefined) return;
    f.text = fn(f.text);
  };

  /* ---- content-bearing removals first (they reference parts) ---------- */

  // tracked changes: accept
  if (want.has('tracked')) {
    for (const n of xmlParts()) {
      if (!WORD_PARTS(n)) continue;
      editXml(n, (x) => {
        let t = x;
        for (const tag of ['w:del', 'w:moveFrom']) {
          const r = editElements(t, tag, 'cut');
          t = r.out;
          did('tracked', r.n);
        }
        for (const tag of ['w:ins', 'w:moveTo']) {
          const r = editElements(t, tag, 'unwrap');
          t = r.out;
          did('tracked', r.n);
        }
        // paragraph-mark and formatting revision records
        for (const tag of ['w:pPrChange', 'w:rPrChange', 'w:tblPrChange', 'w:trPrChange', 'w:tcPrChange', 'w:sectPrChange', 'w:tblGridChange', 'w:cellIns', 'w:cellDel', 'w:cellMerge', 'w:customXmlDelRangeStart', 'w:customXmlDelRangeEnd', 'w:customXmlInsRangeStart', 'w:customXmlInsRangeEnd', 'w:moveFromRangeStart', 'w:moveFromRangeEnd', 'w:moveToRangeStart', 'w:moveToRangeEnd']) {
          const r = editElements(t, tag, 'cut');
          t = r.out;
          did('tracked', r.n);
        }
        // a w:del inside w:rPr/w:pPr would have taken the mark with it above;
        // any leftover w:delText outside a w:del is now ordinary text
        t = t.replace(/<w:delText(\s[^>]*)?>/g, '<w:t$1>').replace(/<\/w:delText>/g, '</w:t>');
        return t;
      });
    }
    editXml('word/settings.xml', (x) => editElements(x, 'w:trackChanges', 'cut').out);
  }

  // comments
  if (want.has('comments')) {
    for (const n of ['word/comments.xml', 'word/commentsExtended.xml', 'word/commentsIds.xml', 'word/commentsExtensible.xml', 'word/threadedComments.xml']) {
      if (killPart(n)) did('comments');
    }
    for (const n of xmlParts()) {
      if (!WORD_PARTS(n)) continue;
      editXml(n, (x) => {
        let t = x;
        for (const tag of ['w:commentRangeStart', 'w:commentRangeEnd', 'w:commentReference']) t = editElements(t, tag, 'cut').out;
        // the run that only carried the reference mark is now empty
        t = editElements(t, 'w:r', 'cut', (inner) => /^(<w:rPr>[\s\S]*<\/w:rPr>)?$/.test(inner.trim())).out;
        return t;
      });
    }
  }

  // invisible text: cut the exact runs the analyzer flagged
  if (want.has('invisible') && (rep.hiddenText || []).length) {
    const flagged = new Set(rep.hiddenText.map((h) => h.text));
    // Only this run's own properties count. A run that anchors a drawing
    // contains the text box's runs, so testing the whole subtree would let a
    // 1pt caption condemn the entire WordArt around it.
    const ownRPr = (inner) => {
      if (!inner.startsWith('<w:rPr>') && !inner.startsWith('<w:rPr ')) return '';
      const at = inner.indexOf('</w:rPr>');
      return at < 0 ? '' : inner.slice(0, at);
    };
    const invisible = (rPr) =>
      /<w:vanish\s*\/>|<w:vanish[^>]*w:val="(1|true|on)"|<w:webHidden\s*\/>/.test(rPr) ||
      /<w:sz[^>]*w:val="([0-2])"/.test(rPr) ||
      /<w:color[^>]*w:val="FFFFFF"/i.test(rPr);
    for (const n of xmlParts()) {
      if (!WORD_PARTS(n)) continue;
      editXml(n, (x) => {
        const r = editElements(x, 'w:r', 'cut', (inner) => {
          const rPr = ownRPr(inner);
          if (!rPr || !invisible(rPr)) return false;
          const body = inner.slice(rPr.length + '</w:rPr>'.length);
          if (/<w:r[ >]/.test(body)) return false; // holds other runs: not a leaf
          return flagged.has(xmlText(body));
        });
        did('invisible', r.n);
        return r.out;
      });
    }
  }

  // images: crops, metadata, unused
  const mediaByPart = new Map((rep.media || []).map((m) => [m.part, m]));
  const rebuilt = new Set(); // images the baker re-encoded: their metadata went with the old bytes
  if (want.has('crops') && opts.bakeImage) {
    for (const m of rep.media || []) {
      if (!m.cropped || !m.crop) continue;
      const f = files.get(m.part);
      if (!f || !f.bytes) continue;
      try {
        const baked = await opts.bakeImage({ bytes: f.bytes, mime: m.mime, crop: m.crop, part: m.part });
        if (baked && baked.length) {
          f.bytes = baked;
          rebuilt.add(m.part);
          did('crops');
        }
      } catch (e) {
        notes.push(`could not re-encode ${m.part}: ${e.message}`);
      }
    }
    if (counts.get('crops')) {
      // the source no longer needs a crop window
      for (const n of xmlParts()) {
        if (WORD_PARTS(n)) editXml(n, (x) => editElements(x, 'a:srcRect', 'cut').out);
      }
    }
  } else if (want.has('crops')) {
    notes.push('crop baking needs an image encoder — the browser build does this with a canvas; the CLI leaves cropped images alone');
  }

  if (want.has('imgmeta')) {
    for (const [name, f] of files) {
      if (!f.bytes || !/^(word\/media|word\/embeddings\/.*media|docProps\/thumbnail)/.test(name)) continue;
      const stripped = stripImageMeta(f.bytes);
      if (stripped) {
        f.bytes = stripped.bytes;
        did('imgmeta');
        did('imgmetaBytes', stripped.dropped);
      } else if (!rebuilt.has(name) && mediaByPart.get(name)?.meta && Object.keys(mediaByPart.get(name).meta).length) {
        notes.push(`${name}: ${mediaByPart.get(name).format} metadata cannot be stripped without re-encoding`);
      }
    }
  }

  if (want.has('orphans')) {
    for (const o of rep.orphanParts || []) if (killPart(o.part)) did('orphans');
  }

  if (want.has('thumbnail')) {
    for (const n of [...files.keys()]) if (/^docProps\/thumbnail\./.test(n) && killPart(n)) did('thumbnail');
  }

  if (want.has('unused')) {
    for (const m of rep.media || []) if (m.unused && killPart(m.part)) did('unused');
  }

  // embedded parts
  if (want.has('glossary')) {
    for (const n of [...files.keys()]) if (/^word\/glossary\//.test(n) && killPart(n)) did('glossary');
  }
  if (want.has('customxml')) {
    for (const n of [...files.keys()]) if (/^customXml\//.test(n) && killPart(n)) did('customxml');
    editXml('docProps/custom.xml', (x) => {
      const r = editElements(x, 'property', 'cut', (inner, start) => /MSIP_Label|_dlc_|ContentTypeId|TemplateUrl|Order|_ip_/.test(start + inner));
      did('customxml', r.n);
      return r.out;
    });
  }
  if (want.has('embeds')) {
    const ids = [];
    for (const n of [...files.keys()]) {
      if (/^word\/embeddings\//.test(n)) {
        for (const [rn, rf] of files) {
          if (!/_rels\/.*\.rels$/.test(rn) || rf.text === undefined) continue;
          const re = /<Relationship\b[^>]*>/g;
          let m;
          while ((m = re.exec(rf.text))) {
            const tgt = (m[0].match(/Target="([^"]*)"/) || [])[1];
            const id = (m[0].match(/Id="([^"]*)"/) || [])[1];
            if (tgt && id && resolveTarget(rn, tgt) === n) ids.push(id);
          }
        }
        if (killPart(n)) did('embeds');
      }
    }
    if (ids.length) {
      for (const n of xmlParts()) {
        if (!WORD_PARTS(n)) continue;
        editXml(n, (x) => {
          let t = x;
          for (const tag of ['w:object', 'w:pict']) {
            t = editElements(t, tag, 'cut', (inner) => ids.some((id) => inner.includes(`"${id}"`))).out;
          }
          return t;
        });
      }
    }
  }
  if (want.has('macros')) {
    for (const n of [...files.keys()]) {
      if (/vbaProject\.bin$|vbaData\.xml$/.test(n) && killPart(n)) did('macros');
    }
    editXml('word/settings.xml', (x) => editElements(x, 'w:attachedTemplate', 'cut').out);
  }

  /* ---- metadata ------------------------------------------------------- */

  if (want.has('identity')) {
    editXml('docProps/core.xml', (x) => {
      let t = x;
      for (const tag of ['dc:creator', 'cp:lastModifiedBy', 'cp:category', 'cp:contentStatus']) {
        const r = setElementText(t, tag, '');
        t = r.out;
        if (r.found) did('identity');
      }
      return t;
    });
    editXml('docProps/app.xml', (x) => {
      let t = x;
      for (const tag of ['Company', 'Manager', 'HyperlinkBase']) {
        const r = setElementText(t, tag, '');
        t = r.out;
        if (r.found) did('identity');
      }
      const tpl = setElementText(t, 'Template', 'Normal.dotm');
      t = tpl.out;
      return t;
    });
    if (killPart('word/people.xml')) did('identity');
    editXml('word/settings.xml', (x) => editElements(x, 'w:attachedTemplate', 'cut').out);
    // any author name still attached to a revision or comment we kept
    for (const n of xmlParts()) {
      if (!WORD_PARTS(n)) continue;
      editXml(n, (x) => {
        // w:author is mandatory on revision and comment elements, so the name is
        // replaced rather than dropped
        const r = stripAttrs(x, / w:author="[^"]*"/g);
        if (r.n) did('identity', 1);
        return r.out.replace(/<w:(?:ins|del|comment|moveFrom|moveTo|pPrChange|rPrChange|tblPrChange|trPrChange|tcPrChange|sectPrChange)\b/g, (m) => `${m} w:author="Author"`);
      });
    }
  }

  if (want.has('timing')) {
    editXml('docProps/app.xml', (x) => {
      const r = setElementText(x, 'TotalTime', '0');
      if (r.found) did('timing');
      return r.out;
    });
    editXml('docProps/core.xml', (x) => {
      let t = x;
      const rev = setElementText(t, 'cp:revision', '1');
      t = rev.out;
      if (rev.found) did('timing');
      const pr = editElements(t, 'cp:lastPrinted', 'cut');
      t = pr.out;
      if (pr.n) did('timing');
      return t;
    });
  }

  if (want.has('refs')) {
    const isLocal = (t) => /^(file:|\\\\|[A-Za-z]:\\)/.test(t || '');
    const deadIds = new Map(); // rels part -> Set(id)
    for (const [rn, rf] of files) {
      if (!/_rels\/.*\.rels$/.test(rn) || rf.text === undefined) continue;
      const kill = new Set();
      rf.text = editElements(rf.text, 'Relationship', 'cut', (inner, start) => {
        const tgt = (start.match(/Target="([^"]*)"/) || [])[1];
        const id = (start.match(/Id="([^"]*)"/) || [])[1];
        if (!isLocal(tgt)) return false;
        if (id) kill.add(id);
        did('refs');
        return true;
      }).out;
      if (kill.size) deadIds.set(rn, kill);
    }
    for (const [rn, kill] of deadIds) {
      const owner = rn.replace('_rels/', '').replace(/\.rels$/, '');
      editXml(owner, (x) => {
        let t = x;
        // unwrap the hyperlink but keep its text
        t = editElements(t, 'w:hyperlink', 'unwrap', (inner, start) => [...kill].some((id) => start.includes(`"${id}"`))).out;
        for (const id of kill) t = t.split(` r:id="${id}"`).join('').split(` r:embed="${id}"`).join('');
        return t;
      });
    }
  }

  /* ---- the save history itself ---------------------------------------- */

  if (want.has('dates')) {
    editXml('docProps/core.xml', (x) => {
      let d = x;
      for (const tag of ['dcterms:created', 'dcterms:modified']) {
        const r = editElements(d, tag, 'cut');
        d = r.out;
        did('dates', r.n);
      }
      return d;
    });
  }

  if (want.has('locales')) {
    for (const n of xmlParts()) {
      if (!WORD_PARTS(n)) continue;
      editXml(n, (x) => {
        const r = editElements(x, 'w:lang', 'cut');
        did('locales', r.n);
        return r.out;
      });
    }
  }

  if (want.has('sessions')) {
    for (const n of xmlParts()) {
      if (!WORD_PARTS(n)) continue;
      editXml(n, (x) => {
        const r = stripAttrs(x, / w:rsid[A-Za-z]*="[0-9A-Fa-f]*"/g);
        did('sessions', r.n);
        return r.out;
      });
    }
    editXml('word/settings.xml', (x) => editElements(x, 'w:rsids', 'cut').out);
  }

  if (want.has('paraIds')) {
    for (const n of xmlParts()) {
      if (!WORD_PARTS(n)) continue;
      editXml(n, (x) => {
        const r = stripAttrs(x, / w1?4?:?(paraId|textId)="[0-9A-Fa-f]*"/g);
        did('paraIds', r.n);
        return r.out;
      });
    }
  }

  if (want.has('future')) {
    editXml('word/settings.xml', (x) => {
      let t = x;
      t = editElements(t, 'w:removePersonalInformation', 'cut').out;
      t = editElements(t, 'w:removeDateAndTime', 'cut').out;
      const flags = '<w:removePersonalInformation/><w:removeDateAndTime/>';
      // CT_Settings is an ordered sequence: these two belong after zoom/view/
      // writeProtection and before everything else.
      let at = -1;
      for (const tag of ['w:writeProtection', 'w:view', 'w:zoom']) {
        const m = t.match(new RegExp(`<${tag}\\b[^>]*?/>|<${tag}\\b[^>]*>[\\s\\S]*?</${tag}>`));
        if (m) at = Math.max(at, m.index + m[0].length);
      }
      if (at < 0) {
        const m = t.match(/<w:settings\b[^>]*>/);
        at = m ? m.index + m[0].length : -1;
      }
      if (at >= 0) {
        t = t.slice(0, at) + flags + t.slice(at);
        did('future', 2);
      }
      return t;
    });
  }

  /* ---- repair the container ------------------------------------------- */

  if (dead.size) {
    // content types: drop overrides for parts that no longer exist
    editXml('[Content_Types].xml', (x) =>
      editElements(x, 'Override', 'cut', (inner, start) => {
        const pn = (start.match(/PartName="\/?([^"]*)"/) || [])[1];
        return pn && dead.has(pn);
      }).out
    );
    // relationships: drop rels pointing at dead parts, then the r:id uses
    for (const [rn, rf] of [...files]) {
      if (!/_rels\/.*\.rels$/.test(rn) || rf.text === undefined) continue;
      const kill = new Set();
      rf.text = editElements(rf.text, 'Relationship', 'cut', (inner, start) => {
        const tgt = (start.match(/Target="([^"]*)"/) || [])[1];
        const resolved = tgt ? resolveTarget(rn, tgt) : null;
        if (resolved && dead.has(resolved)) {
          const id = (start.match(/Id="([^"]*)"/) || [])[1];
          if (id) kill.add(id);
          return true;
        }
        return false;
      }).out;
      const owner = rn.replace('_rels/', '').replace(/\.rels$/, '');
      if (kill.size && files.has(owner) && files.get(owner).text !== undefined) {
        editXml(owner, (x) => {
          let t = x;
          for (const id of kill) {
            t = t.split(` r:id="${id}"`).join('').split(` r:embed="${id}"`).join('').split(` r:link="${id}"`).join('');
          }
          return t;
        });
      }
      // a rels part with nothing left is fine, but an empty glossary dir is not
      if (dead.has(owner)) killPart(rn);
    }
  }

  /* ---- write ----------------------------------------------------------- */

  // [Content_Types].xml must be first in the archive
  const names = [...files.keys()].sort((a, b) => (a === '[Content_Types].xml' ? -1 : b === '[Content_Types].xml' ? 1 : 0));
  const out = [];
  for (const n of names) {
    const f = files.get(n);
    out.push({ name: n, data: f.text !== undefined ? ENC.encode(f.text) : f.bytes });
  }
  const zipped = await writeZip(out, { deflate });

  const applied = [...counts].filter(([k]) => !k.endsWith('Bytes')).map(([id, n]) => ({ id, n }));
  if (counts.get('imgmetaBytes')) applied.push({ id: 'imgmetaBytes', n: counts.get('imgmetaBytes') });
  return { bytes: zipped, applied, notes, removedParts: [...dead] };
}

const norm = (s) => (s || '').replace(/\s+/g, ' ').trim();
const wordsOf = (s) => norm(s).match(/[\p{L}\p{N}][\p{L}\p{N}'\u2019\-]*/gu) || [];

function bag(list) {
  const m = new Map();
  for (const w of list) m.set(w, (m.get(w) || 0) + 1);
  return m;
}

/**
 * Did the scrub remove exactly what it claimed and nothing else?
 *
 * Word multisets rather than string surgery: the flagged fragments are often a
 * single character ("N" in a 1pt checkbox) which naive substring removal would
 * strike out of the wrong place. The invariants are (1) the scrubbed document
 * invents no words, and (2) every word it lost belongs to something the
 * analyzer flagged as deliberately removable.
 */
export function checkTextPreserved(before, after) {
  if (before._plain === undefined || after._plain === undefined) return null;
  const A = bag(wordsOf(before._plain));
  const B = bag(wordsOf(after._plain));
  const allowed = bag(
    [...(before.deletedText || []), ...(before.hiddenText || [])].flatMap((x) => wordsOf(x.text)),
  );
  const invented = [];
  const lost = [];
  for (const [w, n] of B) {
    const extra = n - (A.get(w) || 0);
    if (extra > 0) invented.push([w, extra]);
  }
  for (const [w, n] of A) {
    const missing = n - (B.get(w) || 0);
    if (missing > 0 && missing > (allowed.get(w) || 0)) lost.push([w, missing - (allowed.get(w) || 0)]);
  }
  return {
    ok: invented.length === 0 && lost.length === 0,
    words: { before: [...A.values()].reduce((a, b) => a + b, 0), after: [...B.values()].reduce((a, b) => a + b, 0) },
    invented: invented.slice(0, 8),
    lost: lost.slice(0, 8),
  };
}

/** Compare two reports and describe what the scrub actually achieved. */
export function verifyScrub(before, after) {
  const key = (f) => f.title;
  const gone = before.findings.filter((f) => !after.findings.some((g) => key(g) === key(f)));
  const left = after.findings.filter((f) => f.level !== 'info');
  const num = (r) => ({
    findings: r.findings.filter((f) => f.level !== 'info').length,
    sessions: (r.sessions || []).length,
    identities: (r.identities || []).length,
    hidden: (r.hiddenText || []).length,
    deleted: (r.deletedText || []).length,
    comments: (r.comments || []).length,
    imgMeta: (r.media || []).filter((m) => m.meta && Object.keys(m.meta).length).length,
    cropped: (r.media || []).filter((m) => m.cropped).length,
    unusedImages: (r.media || []).filter((m) => m.unused).length,
    words: r.text ? r.text.words : 0,
    chars: r.text ? r.text.chars : 0,
  });
  return { before: num(before), after: num(after), resolved: gone.map(key), remaining: left.map(key), text: checkTextPreserved(before, after) };
}
