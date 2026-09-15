#!/usr/bin/env node
// test.js — builds a .docx by hand that contains one of every leak Palimpsest
// claims to find, then asserts that the engine finds them. Run: node test.js
import zlib from 'node:zlib';
import { analyze, compareDocs, textReport } from './src/engine.js';
import { readZip, parseXML, findAll } from './src/zipxml.js';
import { imageInfo } from './src/media.js';
import { plan, scrub, verifyScrub, checkTextPreserved, editElements } from './src/scrub.js';

const inflate = async (b) => new Uint8Array(zlib.inflateRawSync(b));
const enc = new TextEncoder();

/* ------------------------------------------------------- zip writing --- */
function crc32(buf) {
  let c, crc = 0xffffffff;
  for (let i = 0; i < buf.length; i++) {
    c = (crc ^ buf[i]) & 0xff;
    for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
    crc = (crc >>> 8) ^ c;
  }
  return (crc ^ 0xffffffff) >>> 0;
}
function makeZip(files) {
  // files: [{name, data: Uint8Array, store?: boolean}]
  const chunks = [];
  const central = [];
  let offset = 0;
  const push = (arr) => {
    chunks.push(arr);
    offset += arr.length;
  };
  for (const f of files) {
    const raw = f.data;
    const comp = f.store ? raw : new Uint8Array(zlib.deflateRawSync(raw, { level: 6 }));
    const nameB = enc.encode(f.name);
    const crc = crc32(raw);
    const lh = new Uint8Array(30 + nameB.length);
    const dv = new DataView(lh.buffer);
    dv.setUint32(0, 0x04034b50, true);
    dv.setUint16(4, 20, true);
    dv.setUint16(8, f.store ? 0 : 8, true);
    dv.setUint16(10, 0, true); // time
    dv.setUint16(12, 0x0021, true); // date 1980-01-01, exactly what Word writes
    dv.setUint32(14, crc, true);
    dv.setUint32(18, comp.length, true);
    dv.setUint32(22, raw.length, true);
    dv.setUint16(26, nameB.length, true);
    lh.set(nameB, 30);
    const localOffset = offset;
    push(lh);
    push(comp);
    const ch = new Uint8Array(46 + nameB.length);
    const cdv = new DataView(ch.buffer);
    cdv.setUint32(0, 0x02014b50, true);
    cdv.setUint16(4, 20, true);
    cdv.setUint16(6, 20, true);
    cdv.setUint16(10, f.store ? 0 : 8, true);
    cdv.setUint16(14, 0x0021, true);
    cdv.setUint32(16, crc, true);
    cdv.setUint32(20, comp.length, true);
    cdv.setUint32(24, raw.length, true);
    cdv.setUint16(28, nameB.length, true);
    cdv.setUint32(42, localOffset, true);
    ch.set(nameB, 46);
    central.push(ch);
  }
  const cdStart = offset;
  for (const c of central) push(c);
  const eocd = new Uint8Array(22);
  const edv = new DataView(eocd.buffer);
  edv.setUint32(0, 0x06054b50, true);
  edv.setUint16(8, central.length, true);
  edv.setUint16(10, central.length, true);
  edv.setUint32(12, offset - cdStart, true);
  edv.setUint32(16, cdStart, true);
  push(eocd);
  const total = chunks.reduce((a, c) => a + c.length, 0);
  const out = new Uint8Array(total);
  let p = 0;
  for (const c of chunks) {
    out.set(c, p);
    p += c.length;
  }
  return out;
}

/* ----------------------------------------------- synthetic JPEG+EXIF --- */
function exifTIFF({ make, model, gps, software }) {
  // minimal TIFF/Exif: IFD0 with Make, Model, Software + GPS IFD pointer
  const parts = [];
  const strings = [];
  let strOff = 0;
  const addStr = (s) => {
    const b = enc.encode(s + '\0');
    strings.push(b);
    const o = strOff;
    strOff += b.length;
    return o;
  };
  const entries = [];
  if (make) entries.push({ tag: 271, type: 2, val: addStr(make), count: make.length + 1 });
  if (model) entries.push({ tag: 272, type: 2, val: addStr(model), count: model.length + 1 });
  if (software) entries.push({ tag: 305, type: 2, val: addStr(software), count: software.length + 1 });
  const nIFD0 = entries.length + (gps ? 1 : 0);
  const ifd0Size = 2 + nIFD0 * 12 + 4;
  const strBase = 8 + ifd0Size;
  const gpsBase = strBase + strOff;
  // GPS IFD: LatRef(ascii), Lat(3 rationals), LonRef, Lon
  const gpsEntries = 4;
  const gpsIfdSize = 2 + gpsEntries * 12 + 4;
  const ratBase = gpsBase + gpsIfdSize;
  const dms = (v) => {
    const a = Math.abs(v);
    const d = Math.floor(a), m = Math.floor((a - d) * 60), s = Math.round(((a - d) * 60 - m) * 60 * 100);
    return [[d, 1], [m, 1], [s, 100]];
  };
  const total = ratBase + (gps ? 48 : 0);
  const buf = new Uint8Array(total);
  const dv = new DataView(buf.buffer);
  buf[0] = 0x49; buf[1] = 0x49; // little endian
  dv.setUint16(2, 42, true);
  dv.setUint32(4, 8, true);
  dv.setUint16(8, nIFD0, true);
  let e = 10;
  for (const en of entries) {
    dv.setUint16(e, en.tag, true);
    dv.setUint16(e + 2, en.type, true);
    dv.setUint32(e + 4, en.count, true);
    dv.setUint32(e + 8, strBase + en.val, true);
    e += 12;
  }
  if (gps) {
    dv.setUint16(e, 34853, true);
    dv.setUint16(e + 2, 4, true);
    dv.setUint32(e + 4, 1, true);
    dv.setUint32(e + 8, gpsBase, true);
    e += 12;
  }
  dv.setUint32(e, 0, true);
  let sp = strBase;
  for (const s of strings) {
    buf.set(s, sp);
    sp += s.length;
  }
  if (gps) {
    dv.setUint16(gpsBase, gpsEntries, true);
    let g = gpsBase + 2;
    const put = (tag, type, count, valOrOff, inline) => {
      dv.setUint16(g, tag, true);
      dv.setUint16(g + 2, type, true);
      dv.setUint32(g + 4, count, true);
      if (inline) buf[g + 8] = valOrOff;
      else dv.setUint32(g + 8, valOrOff, true);
      g += 12;
    };
    put(1, 2, 2, gps[0] >= 0 ? 78 : 83, true); // 'N' / 'S'
    put(2, 5, 3, ratBase, false);
    put(3, 2, 2, gps[1] >= 0 ? 69 : 87, true); // 'E' / 'W'
    put(4, 5, 3, ratBase + 24, false);
    dv.setUint32(g, 0, true);
    const lat = dms(gps[0]), lon = dms(gps[1]);
    lat.forEach(([n, d], i) => {
      dv.setUint32(ratBase + i * 8, n, true);
      dv.setUint32(ratBase + i * 8 + 4, d, true);
    });
    lon.forEach(([n, d], i) => {
      dv.setUint32(ratBase + 24 + i * 8, n, true);
      dv.setUint32(ratBase + 24 + i * 8 + 4, d, true);
    });
  }
  return buf;
}

function jpegWithExif(opts) {
  const buf = exifTIFF(opts);
  const app1 = new Uint8Array(2 + 2 + 6 + buf.length);
  const adv = new DataView(app1.buffer);
  adv.setUint16(0, 0xffe1);
  adv.setUint16(2, 2 + 6 + buf.length);
  app1.set(enc.encode('Exif\0\0'), 4);
  app1.set(buf, 10);
  const sof = new Uint8Array([0xff, 0xc0, 0x00, 0x11, 0x08, 0x02, 0x58, 0x03, 0x20, 0x03, 0x01, 0x11, 0x00, 0x02, 0x11, 0x01, 0x03, 0x11, 0x01]);
  const out = new Uint8Array(2 + app1.length + sof.length + 2);
  out.set([0xff, 0xd8], 0);
  out.set(app1, 2);
  out.set(sof, 2 + app1.length);
  out.set([0xff, 0xd9], 2 + app1.length + sof.length);
  return out;
}


/* ----------------------------------------------- synthetic PNG image --- */
// A real, renderable PNG: an innocuous gradient inside the crop window and a
// loud red band in the region Word crops away, so "excavate" has something to show.
function pngWithExif(w, h, exif) {
  const raw = new Uint8Array((w * 3 + 1) * h);
  let o = 0;
  for (let y = 0; y < h; y++) {
    raw[o++] = 0; // filter: none
    for (let x = 0; x < w; x++) {
      const inWindow = x >= w * 0.25 && x <= w * 0.85 && y >= h * 0.1 && y <= h * 0.95;
      let r, g, b;
      if (inWindow) {
        r = 40 + Math.round((x / w) * 60);
        g = 70 + Math.round((y / h) * 70);
        b = 90 + Math.round(((x + y) / (w + h)) * 90);
        if ((x + y) % 90 < 3) { r = 120; g = 150; b = 170; }
      } else {
        const band = y > h * 0.1 && y < h * 0.62;
        r = band ? 200 : 60;
        g = band ? 40 : 20;
        b = band ? 40 : 20;
        if (band && ((x % 40 < 22 && y % 60 < 34) || y % 60 > 52)) { r = 250; g = 240; b = 230; }
      }
      raw[o++] = r; raw[o++] = g; raw[o++] = b;
    }
  }
  const chunk = (type, data) => {
    const out = new Uint8Array(12 + data.length);
    const dv = new DataView(out.buffer);
    dv.setUint32(0, data.length);
    out.set(enc.encode(type), 4);
    out.set(data, 8);
    const crcBuf = new Uint8Array(4 + data.length);
    crcBuf.set(enc.encode(type), 0);
    crcBuf.set(data, 4);
    dv.setUint32(8 + data.length, crc32(crcBuf));
    return out;
  };
  const ihdr = new Uint8Array(13);
  const idv = new DataView(ihdr.buffer);
  idv.setUint32(0, w);
  idv.setUint32(4, h);
  ihdr[8] = 8; ihdr[9] = 2; // 8-bit RGB
  const parts = [
    new Uint8Array([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]),
    chunk('IHDR', ihdr),
    chunk('eXIf', exifTIFF(exif)),
    chunk('IDAT', new Uint8Array(zlib.deflateSync(raw, { level: 6 }))),
    chunk('IEND', new Uint8Array(0)),
  ];
  const total = parts.reduce((a, c) => a + c.length, 0);
  const out = new Uint8Array(total);
  let q = 0;
  for (const c of parts) { out.set(c, q); q += c.length; }
  return out;
}

/* ------------------------------------------------------ the document --- */
const R_ROOT = '00A1B2C3'; // creating session
const R_2 = '00D4E5F6'; // second session
const R_3 = '0011AA22'; // third session (pasted block)
const R_FOREIGN = '00FFEE01'; // id that is NOT in this file's session list
const NS =
  'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ' +
  'xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" ' +
  'xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml" ' +
  'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ' +
  'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" ' +
  'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"';

const para = (rsid, text, extra = '') =>
  `<w:p w14:paraId="${(Math.random() * 0xffffffff >>> 0).toString(16).toUpperCase().padStart(8, '0')}" w:rsidR="${rsid}" w:rsidRDefault="${rsid}">${extra}<w:r w:rsidR="${rsid}"><w:t xml:space="preserve">${text}</w:t></w:r></w:p>`;

const documentXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document ${NS}><w:body>
${para(R_ROOT, 'Quarterly Review of the Northern Division')}
${para(R_ROOT, 'This paragraph was typed when the file was created.')}
${para(R_2, 'This sentence was added in a later session.')}
<w:p w14:paraId="AAAA0001" w:rsidR="${R_2}" w:rsidRDefault="${R_2}">
  <w:r w:rsidR="${R_2}"><w:t xml:space="preserve">The decision was </w:t></w:r>
  <w:del w:id="9" w:author="Dana Whitlock" w:date="2024-03-11T09:14:00Z"><w:r w:rsidR="${R_2}" w:rsidDel="${R_3}"><w:delText xml:space="preserve">rejected outright by legal</w:delText></w:r></w:del>
  <w:ins w:id="10" w:author="Dana Whitlock" w:date="2024-03-11T09:15:00Z"><w:r w:rsidR="${R_3}"><w:t xml:space="preserve">approved with conditions</w:t></w:r></w:ins>
  <w:r w:rsidR="${R_2}"><w:t>.</w:t></w:r>
</w:p>
<w:p w14:paraId="AAAA0002" w:rsidR="${R_3}" w:rsidRDefault="${R_3}">
  <w:commentRangeStart w:id="1"/>
  <w:r w:rsidR="${R_3}" w:rsidRPr="${R_FOREIGN}"><w:rPr><w:rFonts w:ascii="Times New Roman"/><w:lang w:val="de-DE"/></w:rPr><w:t xml:space="preserve">Dieser Absatz wurde aus einem anderen Dokument eingefügt und behält seine Formatierungs-ID.</w:t></w:r>
  <w:commentRangeEnd w:id="1"/>
</w:p>
${para(R_3, 'Pasted block, second paragraph, same session as the one above.')}
<w:p w14:paraId="AAAA0006" w:rsidR="${R_2}" w:rsidRDefault="${R_2}"><w:r w:rsidR="${R_2}"><w:drawing><wp:inline><wp:extent cx="1143000" cy="857250"/><wp:docPr id="2" name="Picture 2" descr="phone snap of the whiteboard"/><a:graphic><a:graphicData><a:pic><a:blipFill><a:blip r:embed="rId5"/></a:blipFill></a:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>
<w:p w14:paraId="AAAA0003" w:rsidR="${R_2}" w:rsidRDefault="${R_2}">
  <w:r w:rsidR="${R_2}"><w:rPr><w:vanish/></w:rPr><w:t xml:space="preserve">INTERNAL ONLY: do not send to the client before Friday.</w:t></w:r>
  <w:r w:rsidR="${R_2}"><w:rPr><w:color w:val="FFFFFF"/></w:rPr><w:t xml:space="preserve">strategy consultant senior partner mckinsey keyword keyword</w:t></w:r>
</w:p>
<w:p w14:paraId="AAAA0004" w:rsidR="${R_2}" w:rsidRDefault="${R_2}"><w:r w:rsidR="${R_2}"><w:drawing><wp:inline><wp:extent cx="2286000" cy="1143000"/><wp:docPr id="1" name="Picture 1" descr="site photo"/><a:graphic><a:graphicData><a:pic><a:blipFill><a:blip r:embed="rId8"/><a:srcRect l="25000" t="10000" r="15000" b="5000"/></a:blipFill></a:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>
<w:p w14:paraId="AAAA0005" w:rsidR="${R_2}" w:rsidRDefault="${R_2}"><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> MERGEFIELD  ClientName  \\* MERGEFORMAT </w:instrText></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>
<w:sectPr w:rsidR="${R_2}" w:rsidSect="${R_ROOT}"/>
</w:body></w:document>`;

const settingsXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:settings ${NS}>
<w:proofState w:spelling="clean" w:grammar="dirty"/>
<w:attachedTemplate r:id="rId1"/>
<w:documentProtection w:edit="readOnly" w:enforcement="1" w:cryptProviderType="rsaAES" w:hash="Zm9vYmFy"/>
<w:trackChanges/>
<w:themeFontLang w:val="en-GB"/>
<w:rsids><w:rsidRoot w:val="${R_ROOT}"/><w:rsid w:val="${R_ROOT}"/><w:rsid w:val="${R_2}"/><w:rsid w:val="${R_3}"/><w:rsid w:val="0099CC44"/><w:rsid w:val="00778899"/></w:rsids>
</w:settings>`;

const files = [
  {
    name: '[Content_Types].xml',
    data: enc.encode(
      `<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="jpeg" ContentType="image/jpeg"/><Default Extension="png" ContentType="image/png"/><Default Extension="xlsx" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/comments.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"/><Override PartName="/word/glossary/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.glossary+xml"/><Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/><Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/><Override PartName="/docProps/custom.xml" ContentType="application/vnd.openxmlformats-officedocument.custom-properties+xml"/><Override PartName="/docProps/thumbnail.jpeg" ContentType="image/jpeg"/></Types>`,
    ),
  },
  {
    name: '_rels/.rels',
    data: enc.encode(
      `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/><Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties" Target="docProps/custom.xml"/><Relationship Id="rId5" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/thumbnail" Target="docProps/thumbnail.jpeg"/></Relationships>`,
    ),
  },
  { name: 'word/document.xml', data: enc.encode(documentXml) },
  {
    name: 'word/_rels/document.xml.rels',
    data: enc.encode(`<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
      <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>
      <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
      <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/fontTable" Target="fontTable.xml"/>
      <Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="comments.xml"/>
      <Relationship Id="rId5" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image1.jpeg"/>
      <Relationship Id="rId6" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="file:///Y:/Legal/Matters/2024/NorthernDivision/final-draft.docx" TargetMode="External"/>
      <Relationship Id="rId7" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/package" Target="embeddings/Microsoft_Excel_Worksheet.xlsx"/>
      <Relationship Id="rId8" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image2.png"/>
      <Relationship Id="rId9" Type="http://schemas.microsoft.com/office/2011/relationships/people" Target="people.xml"/>
      <Relationship Id="rId10" Type="http://schemas.microsoft.com/office/2011/relationships/commentsExtended" Target="commentsExtended.xml"/>
      <Relationship Id="rId11" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/glossaryDocument" Target="glossary/document.xml"/>
      </Relationships>`),
  },
  { name: 'word/settings.xml', data: enc.encode(settingsXml) },
  {
    name: 'word/_rels/settings.xml.rels',
    data: enc.encode(`<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/attachedTemplate" Target="\\\\fileserver01\\templates\\Client-Memo-v4.dotx" TargetMode="External"/></Relationships>`),
  },
  { name: 'word/styles.xml', data: enc.encode(`<?xml version="1.0"?><w:styles ${NS}><w:style w:styleId="Heading1"><w:name w:val="heading 1"/></w:style></w:styles>`) },
  { name: 'word/fontTable.xml', data: enc.encode(`<?xml version="1.0"?><w:fonts ${NS}><w:font w:name="Calibri"/><w:font w:name="Times New Roman"/><w:font w:name="SimSun"/></w:fonts>`) },
  {
    name: 'word/comments.xml',
    data: enc.encode(`<?xml version="1.0"?><w:comments ${NS}><w:comment w:id="1" w:author="Priya Raman" w:initials="PR" w:date="2024-03-12T16:02:00Z"><w:p w14:paraId="BBBB0001"><w:r><w:t>Do we really want to admit this in writing?</w:t></w:r></w:p></w:comment></w:comments>`),
  },
  { name: 'word/commentsExtended.xml', data: enc.encode(`<?xml version="1.0"?><w15:commentsEx ${NS}><w15:commentEx w15:paraId="BBBB0001" w15:done="0"/></w15:commentsEx>`) },
  {
    name: 'word/people.xml',
    data: enc.encode(`<?xml version="1.0"?><w15:people ${NS}><w15:person w15:author="Priya Raman"><w15:presenceInfo w15:providerId="AD" w15:userId="S::priya.raman@northern-division.example.com::a1b2c3"/></w15:person></w15:people>`),
  },
  {
    name: 'word/glossary/document.xml',
    data: enc.encode(`<?xml version="1.0"?><w:glossaryDocument ${NS}><w:docParts><w:docPart><w:docPartBody><w:p><w:r><w:t>Confidential – Attorney Work Product – Ferris &amp; Vale LLP</w:t></w:r></w:p></w:docPartBody></w:docPart></w:docParts></w:glossaryDocument>`),
  },
  { name: 'word/media/image1.jpeg', data: jpegWithExif({ make: 'Apple', model: 'iPhone 13 Pro', software: '16.5.1', gps: [37.8267, -122.4233] }) },
  { name: 'word/media/image2.png', data: pngWithExif(900, 620, { software: 'Screenshot', make: 'Contoso' }) },
  { name: 'word/embeddings/Microsoft_Excel_Worksheet.xlsx', data: enc.encode('PK\x03\x04 pretend workbook with every cell intact ' + 'x'.repeat(3000)) },
  {
    name: 'docProps/core.xml',
    data: enc.encode(`<?xml version="1.0"?><cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/"><dc:title>Quarterly Review</dc:title><dc:creator>Dana Whitlock</dc:creator><cp:lastModifiedBy>Priya Raman</cp:lastModifiedBy><cp:revision>4</cp:revision><dcterms:created>2024-03-10T08:00:00Z</dcterms:created><dcterms:modified>2024-03-12T16:30:00Z</dcterms:modified><cp:lastPrinted>2024-03-12T09:00:00Z</cp:lastPrinted></cp:coreProperties>`),
  },
  {
    name: 'docProps/app.xml',
    data: enc.encode(`<?xml version="1.0"?><Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"><Template>Client-Memo-v4.dotx</Template><TotalTime>2</TotalTime><Pages>3</Pages><Words>820</Words><Characters>4800</Characters><Application>Microsoft Office Word</Application><AppVersion>16.0000</AppVersion><Company>Ferris &amp; Vale LLP</Company><Manager>Alan Ferris</Manager></Properties>`),
  },
  {
    name: 'docProps/custom.xml',
    data: enc.encode(`<?xml version="1.0"?><Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/custom-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
      <property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="2" name="MSIP_Label_2f1c3e70_Name"><vt:lpwstr>Confidential \\ Legal Privileged</vt:lpwstr></property>
      <property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="3" name="MSIP_Label_2f1c3e70_Owner"><vt:lpwstr>dana.whitlock@ferrisvale.example.com</vt:lpwstr></property>
      <property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="4" name="MSIP_Label_2f1c3e70_SetDate"><vt:lpwstr>2024-03-10T08:04:19Z</vt:lpwstr></property>
      </Properties>`),
  },
  { name: 'docProps/thumbnail.jpeg', data: jpegWithExif({ software: 'thumbnail' }) },
];

/* ---------------------------------------------------------- assertions --- */
let pass = 0, fail = 0;
const ok = (name, cond, extra) => {
  if (cond) {
    pass++;
    console.log(`\x1b[38;5;108m  ok  \x1b[0m ${name}`);
  } else {
    fail++;
    console.log(`\x1b[1;31m FAIL \x1b[0m ${name}${extra !== undefined ? '  ← ' + JSON.stringify(extra) : ''}`);
  }
};

const bytes = makeZip(files);
if (process.argv.includes('--emit')) {
  const { writeFileSync } = await import('node:fs');
  const target = process.argv[process.argv.indexOf('--emit') + 1] || 'samples/synthetic-everything.docx';
  writeFileSync(target, bytes);
  console.log('wrote ' + target + ' (' + bytes.length + ' bytes)');
}
const rep = await analyze(bytes, { inflate, name: 'synthetic.docx', keepBytes: true });
const titles = rep.findings.map((f) => f.title).join(' | ');
const has = (re) => re.test(titles);

console.log('\n\x1b[1msynthetic document with one of everything\x1b[0m');
ok('zip is parsed without a library', rep.container.partCount === files.length, rep.container.partCount);
ok('document text is extracted in order', rep.text.preview.startsWith('Quarterly Review of the Northern Division'), rep.text.preview.slice(0, 40));
ok('deleted text excluded from the visible word count', !rep.text.preview.includes('rejected outright'));

ok('rsidRoot identified as the creating session', rep.sessions.find((s) => s.isRoot)?.rsid === R_ROOT);
ok('native sessions counted from settings.xml', rep.settings.rsidCount === 5, rep.settings.rsidCount);
ok('sessions with no surviving text are still listed', rep.sessions.filter((s) => s.native && !s.chars).length === 2, rep.sessions.filter((s) => s.native && !s.chars).map((s) => s.rsid));
ok('three sessions own visible text', rep.sessions.filter((s) => s.native && s.chars > 0).length === 3);
ok('foreign formatting id detected', rep.sessions.some((s) => s.rsid === R_FOREIGN && !s.native));
ok('root session anchored to the creation timestamp', rep.sessions.find((s) => s.isRoot)?.anchoredAt === '2024-03-10T08:00:00.000Z');
ok('deleting session ordered after the inserting session', (() => {
  const a = rep.sessions.find((s) => s.rsid === R_2), b = rep.sessions.find((s) => s.rsid === R_3);
  return a && b && a.layer < b.layer;
})(), rep.sessions.filter((s) => s.native).map((s) => [s.rsid, s.layer]));
ok('contiguous block reported as an ID group without asserting how it was written', has(/has a distinct revision ID/), titles);

ok('tracked deletion text recovered', rep.deletedText.some((d) => d.text.includes('rejected outright by legal') && d.author === 'Dana Whitlock'));
ok('deletion is dated', rep.deletedText[0]?.date === '2024-03-11T09:14:00Z');
ok('hidden (w:vanish) text recovered', rep.hiddenText.some((h) => /INTERNAL ONLY/.test(h.text) && /vanish/.test(h.kind)));
ok('white-on-white keyword stuffing recovered', rep.hiddenText.some((h) => /mckinsey/.test(h.text) && /white/.test(h.kind)));
ok('comment body, author and resolution state read', rep.comments[0]?.author === 'Priya Raman' && /admit this in writing/.test(rep.comments[0].text) && rep.comments[0].resolved === false);
ok('co-author email lifted from people.xml', rep.identities.some((i) => i.name === 'priya.raman@northern-division.example.com'));
ok('every name is collected with its role', ['Dana Whitlock', 'Priya Raman', 'Alan Ferris'].every((n) => rep.identities.some((i) => i.name === n)), rep.identities.map((i) => i.name));

ok('image crop geometry computed', (() => {
  const m = rep.media.find((x) => x.part.endsWith('image2.png'));
  return m && m.cropped && Math.abs(m.visibleFraction - 51) < 1.5;
})(), rep.media.map((m) => [m.part, m.visibleFraction]));
ok('EXIF camera make/model parsed from embedded image', (() => {
  const m = rep.media.find((x) => x.part.endsWith('image1.jpeg'));
  return m?.meta.Make === 'Apple' && m.meta.Model === 'iPhone 13 Pro';
})(), rep.media[0]?.meta);
ok('EXIF GPS converted to decimal degrees', (() => {
  const m = rep.media.find((x) => x.part.endsWith('image1.jpeg'));
  if (!m?.meta.GPS) return false;
  const [lat, lon] = m.meta.GPS.split(',').map(Number);
  return Math.abs(lat - 37.8267) < 0.001 && Math.abs(lon + 122.4233) < 0.001;
})(), rep.media.find((x) => x.part.endsWith('image1.jpeg'))?.meta);
ok('JPEG dimensions read from SOF0', (() => {
  const m = rep.media.find((x) => x.part.endsWith('image1.jpeg'));
  return m?.width === 800 && m?.height === 600;
})(), [rep.media[0]?.width, rep.media[0]?.height]);

ok('sensitivity label decoded with owner and date', has(/sensitivity label/i) && rep.labels.some((l) => l.field === 'Owner' && /dana.whitlock/.test(l.value)));
ok('UNC template path surfaced', rep.settings.attachedTemplate?.includes('fileserver01'), rep.settings.attachedTemplate);
ok('external file:// reference surfaced', rep.external.some((e) => /Y:\/Legal/.test(e.target)));
ok('embedded workbook reported', has(/embedded object/i) && rep.embedded.some((e) => e.kind === 'embedded object'));
ok('AutoText/glossary content recovered', rep.embedded.some((e) => e.kind === 'AutoText / building blocks' && /Attorney Work Product/.test(e.preview)));
ok('MERGEFIELD field code captured', rep.fields.some((f) => /MERGEFIELD\s+ClientName/.test(f.instr)));
ok('document protection reported as cosmetic', has(/Editing restrictions/i) && rep.settings.documentProtection.hasHash === true);
ok('track-changes-left-on warned', has(/Track changes is currently switched on/i));
ok('mixed run languages detected', rep.locales.some((l) => l.name === 'de-DE'), rep.locales);
ok('typing rate implausibility computed', has(/words per minute of recorded editing/), titles);
ok('printing recorded', has(/was printed/));
ok('producer version mapped to a product name', rep.props.appVersionLabel === 'Word 2016 / 2019 / 365');
ok('verdict headline prefers the recoverable deletion', rep.verdict.headline === 'Text that was deleted is still inside this file', rep.verdict.headline);

/* ------------------------------------------------- lineage comparison --- */
// Save the same document again with one paragraph rewritten and a new session.
const R_4 = '00445566';
const docV2 = documentXml
  .replace('This sentence was added in a later session.', 'This sentence was quietly rewritten before signing.')
  .replace('<w:sectPr', `${para(R_4, 'A paragraph appended in a fourth session.')}<w:sectPr`);
const setV2 = settingsXml.replace('<w:rsid w:val="00778899"/>', `<w:rsid w:val="00778899"/><w:rsid w:val="${R_4}"/>`);
const files2 = files.map((f) =>
  f.name === 'word/document.xml' ? { ...f, data: enc.encode(docV2) } : f.name === 'word/settings.xml' ? { ...f, data: enc.encode(setV2) } : f,
);
const rep2 = await analyze(makeZip(files2), { inflate, name: 'synthetic-v2.docx' });
const L = compareDocs(rep, rep2, 'v2.docx');
console.log('\n\x1b[1mlineage: the same file saved again\x1b[0m');
ok('descendant detected from the session list', L.relation === 'descendant', L);
ok('the added session is the only difference', L.onlyB.length === 1 && L.onlyB[0] === R_4 && L.onlyA.length === 0, { onlyA: L.onlyA, onlyB: L.onlyB });
ok('paragraph identity survives a rewrite', L.paragraphs.rewritten.length === 1 && /quietly rewritten/.test(L.paragraphs.rewritten[0].after) && /added in a later session/.test(L.paragraphs.rewritten[0].before), L.paragraphs && L.paragraphs.rewritten);
ok('appended paragraph reported as new', L.paragraphs.added.length === 1 && /fourth session/.test(L.paragraphs.added[0].text), L.paragraphs && L.paragraphs.added);
ok('unrelated documents are not linked', compareDocs(rep, { sessions: [{ rsid: '00ABCDEF', native: true }], settings: { rsidRoot: '00ABCDEF' }, _model: { paragraphs: [] } }).relation === 'unrelated');
ok('text report includes the recovered deletion', /rejected outright by legal/.test(textReport(rep)));

console.log('\n\x1b[1mparsers\x1b[0m');
const x = parseXML('<a b="1" c=\'2 > 3\'><d/>te&amp;xt<!-- c --><e f="&#65;"/><![CDATA[<raw>]]></a>');
ok('attributes with > inside quotes parse', x.kids[0].attrs.c === '2 > 3');
ok('numeric entities decode', findAll(x, 'e')[0].attrs.f === 'A');
ok('CDATA and entities in text decode', JSON.stringify(x.kids[0].kids.filter((k) => typeof k === 'string')) === '["te&xt","<raw>"]', x.kids[0].kids.filter((k) => typeof k === 'string'));
const z = await readZip(bytes, inflate);
ok('stored and deflated entries both read back', (await z.text('word/settings.xml')).includes(R_ROOT));
ok('zip records Word\'s zeroed 1980 timestamps', z.entries.get('word/document.xml').mtime === '1980-01-01T00:00:00');
const png = new Uint8Array([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 0, 0, 0, 13, 73, 72, 68, 82, 0, 0, 1, 44, 0, 0, 0, 200, 8, 6, 0, 0, 0, 0, 0, 0, 0]);
ok('PNG geometry read', imageInfo(png).width === 300 && imageInfo(png).height === 200, imageInfo(png));
ok('PNG eXIf chunk parsed', rep.media.find((m) => m.part.endsWith('image2.png'))?.meta.Software === 'Screenshot', rep.media.find((m) => m.part.endsWith('image2.png'))?.meta);
ok('non-image bytes degrade quietly', imageInfo(enc.encode('not an image at all')).format === null);

/* ============================================================ scrubbing == */
console.log('\n\x1b[1mscrub: removing what was found\x1b[0m');
const deflate = async (b) => new Uint8Array(zlib.deflateRawSync(b, { level: 6 }));
const fixIds = plan(rep).map((f) => f.id);
ok('every leak in the fixture has a fix offered', ['sessions', 'paraIds', 'identity', 'timing', 'dates', 'tracked', 'comments', 'invisible', 'imgmeta', 'crops', 'thumbnail', 'glossary', 'customxml', 'refs', 'embeds', 'locales'].every((id) => fixIds.includes(id)), fixIds);

const cleaned = await scrub(bytes, rep, { inflate, deflate });
const repC = await analyze(cleaned.bytes, { inflate, name: 'clean.docx' });
const zc = await readZip(cleaned.bytes, inflate);
const docC = await zc.text('word/document.xml');

ok('the cleaned file still parses as a .docx', repC.text.paragraphs === rep.text.paragraphs, { before: rep.text.paragraphs, after: repC.text.paragraphs });
ok('no rsid stamps survive anywhere', !/\sw:rsid[A-Za-z]*="/.test(docC) && !/<w:rsids>/.test(await zc.text('word/settings.xml')));
ok('no session is left to analyse', repC.sessions.length === 0 && repC.settings.rsidCount === 0, { sessions: repC.sessions.length, rsidCount: repC.settings.rsidCount });
ok('paragraph ids are gone', !/w14:paraId=/.test(docC) && repC.paraIdCount === 0);
ok('no names remain', repC.identities.length === 0, repC.identities);
ok('tracked deletion is applied, not just hidden', repC.deletedText.length === 0 && !/rejected outright by legal/.test(docC));
ok('the insertion survived as ordinary text', /approved with conditions/.test(docC) && !/<w:ins\b/.test(docC));
ok('hidden and white text are gone', repC.hiddenText.length === 0 && !/do not send to the client/.test(docC), repC.hiddenText);
ok('comments are gone, anchors and all', repC.comments.length === 0 && !zc.has('word/comments.xml') && !/commentRangeStart|commentReference/.test(docC));
ok('AutoText part is gone', !zc.has('word/glossary/document.xml'));
ok('customXml and the sensitivity label are gone', repC.labels.length === 0 && ![...zc.order].some((n) => n.startsWith('customXml/')), { labels: repC.labels, parts: [...zc.order].filter((n) => n.startsWith('customXml/')) });
ok('people.xml (account ids) is gone', !zc.has('word/people.xml'));
ok('the UNC path reference is gone', repC.external.filter((e) => e.local).length === 0, repC.external);
ok('EXIF, GPS and XMP are stripped from the images', repC.media.every((m) => Object.keys(m.meta).length === 0), repC.media.map((m) => [m.part, m.meta]));
ok('the editing clock is reset', repC.props.totalTime === 0 && repC.props.revision === '1' && !repC.props.lastPrinted);
ok('timestamps are gone', !repC.props.created && !repC.props.modified);
ok('run language tags are gone', repC.locales.length === 0, repC.locales);
ok('the page thumbnail is gone', !zc.has('docProps/thumbnail.jpeg'));
ok('Word is told to keep stripping on save', repC.settings.removePersonalInformation === true);
ok('the scrubbed fixture has sparse metadata without claiming its history is proven', repC.scrubbed === true && /sparse editing metadata/.test(repC.verdict.headline), repC.verdict.headline);

const tp = checkTextPreserved(rep, repC);
ok('visible text survives except the flagged passages', tp.ok, tp);
ok('nothing was invented', tp.invented.length === 0, tp.invented);

// the container has to stay legal: content types, relationships, both directions
const ct = await zc.text('[Content_Types].xml');
ok('[Content_Types].xml is still the first entry', zc.order[0] === '[Content_Types].xml', zc.order[0]);
ok('no content type points at a part that was removed', findAll(parseXML(ct), 'Override').every((o) => zc.has((o.attrs.PartName || '').replace(/^\//, ''))), findAll(parseXML(ct), 'Override').map((o) => o.attrs.PartName).filter((n) => !zc.has(n.replace(/^\//, ''))));
const relIds = new Set(findAll(parseXML(await zc.text('word/_rels/document.xml.rels')), 'Relationship').map((r) => r.attrs.Id));
const usedIds = [...docC.matchAll(/\sr:(?:id|embed|link)="([^"]+)"/g)].map((m) => m[1]);
ok('every r:id in the document still resolves', usedIds.every((id) => relIds.has(id)), usedIds.filter((id) => !relIds.has(id)));
const relTargets = findAll(parseXML(await zc.text('word/_rels/document.xml.rels')), 'Relationship')
  .filter((r) => r.attrs.TargetMode !== 'External' && !/^[a-z]+:|^\\\\/i.test(r.attrs.Target || ''))
  .map((r) => 'word/' + (r.attrs.Target || '').replace(/^\.?\//, ''));
ok('no relationship dangles', relTargets.every((t) => zc.has(t)), relTargets.filter((t) => !zc.has(t)));

// crop baking is the browser's job; a Node scrub must say so rather than pretend
ok('the CLI reports that it cannot bake crops', cleaned.notes.some((n) => /crop baking needs an image encoder/.test(n)), cleaned.notes);
ok('crop windows are untouched without an encoder', repC.media.some((m) => m.cropped), repC.media.map((m) => m.cropped));

// and with an encoder, the crop is applied and the window removed
const baked = await scrub(bytes, rep, {
  inflate,
  deflate,
  bakeImage: async ({ bytes: b, crop }) => {
    // stand in for the browser's canvas: hand back a smaller valid PNG
    const info = imageInfo(b);
    const w = Math.max(1, Math.round(info.width * (1 - (crop.l + crop.r) / 100)));
    const h = Math.max(1, Math.round(info.height * (1 - (crop.t + crop.b) / 100)));
    return pngWithExif(w, h, {});
  },
});
const repB = await analyze(baked.bytes, { inflate, name: 'baked.docx' });
ok('with an encoder the crop is baked in', repB.media.every((m) => !m.cropped), repB.media.map((m) => [m.part, m.cropped]));
const cropped = rep.media.find((m) => m.cropped);
const bakedImg = repB.media.find((m) => m.part === cropped.part);
ok('the baked image is the visible window only', bakedImg.width < cropped.width && bakedImg.height < cropped.height, { before: [cropped.width, cropped.height], after: [bakedImg.width, bakedImg.height] });

// opting out has to actually opt out
const minimal = await scrub(bytes, rep, { inflate, deflate, fixes: ['sessions'] });
const repM = await analyze(minimal.bytes, { inflate, name: 'minimal.docx' });
ok('a single-fix scrub touches only that fix', repM.sessions.length === 0 && repM.identities.length === rep.identities.length && repM.comments.length === rep.comments.length, { identities: repM.identities.length, comments: repM.comments.length });
ok('opting into embedded-file removal removes them', (await analyze((await scrub(bytes, rep, { inflate, deflate, fixes: ['embeds'] })).bytes, { inflate, name: 'e.docx' })).embedded.filter((e) => e.kind === 'embedded object').length === 0);

// element surgery: the part that would quietly corrupt a document
ok('editElements does not confuse w:del with w:delText', editElements('<w:p><w:del><w:r><w:delText>gone</w:delText></w:r></w:del><w:r><w:t>kept</w:t></w:r></w:p>', 'w:del', 'cut').out === '<w:p><w:r><w:t>kept</w:t></w:r></w:p>');
ok('editElements handles nesting of the same tag', editElements('<a><x><y/><x><y/></x></x></a>', 'x', 'cut').out === '<a></a>');
ok('editElements unwraps without eating siblings', editElements('<p><i attr="1">in</i>out</p>', 'i', 'unwrap').out === '<p>inout</p>');
ok('editElements descends into elements it rejects', editElements('<a><r><r hit="1">x</r></r></a>', 'r', 'cut', (inner, start) => /hit/.test(start)).out === '<a><r></r></a>', editElements('<a><r><r hit="1">x</r></r></a>', 'r', 'cut', (inner, start) => /hit/.test(start)).out);
ok('editElements tolerates > inside attribute values', editElements('<a><b c="2 > 3">x</b>y</a>', 'b', 'cut').out === '<a>y</a>');
ok('a malformed part is left alone rather than mangled', editElements('<a><b>unclosed</a>', 'b', 'cut').out === '<a><b>unclosed</a>');

console.log(`\n${fail ? '\x1b[1;31m' : '\x1b[1;32m'}${pass} passed, ${fail} failed\x1b[0m\n`);
process.exit(fail ? 1 : 0);
