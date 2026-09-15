// media.js — image geometry + embedded-metadata extraction (JPEG/PNG/GIF/BMP/TIFF/WMF/EMF).
// Dependency-free. Returns only what is actually present in the bytes.

const dec = new TextDecoder('latin1');

export function imageInfo(bytes) {
  const out = { format: null, width: null, height: null, meta: {}, warnings: [] };
  if (!bytes || bytes.length < 16) return out;
  const b = bytes;
  const u16 = (i, le) => (le ? b[i] | (b[i + 1] << 8) : (b[i] << 8) | b[i + 1]);
  const u32 = (i, le) => (le ? (b[i] | (b[i + 1] << 8) | (b[i + 2] << 16) | (b[i + 3] << 24)) >>> 0 : ((b[i] << 24) | (b[i + 1] << 16) | (b[i + 2] << 8) | b[i + 3]) >>> 0);

  if (b[0] === 0x89 && b[1] === 0x50) {
    out.format = 'png';
    out.width = u32(16, false);
    out.height = u32(20, false);
    // chunks: look for text metadata + physical density
    let i = 8;
    while (i + 8 < b.length) {
      const len = u32(i, false);
      const type = dec.decode(b.subarray(i + 4, i + 8));
      const data = b.subarray(i + 8, i + 8 + len);
      if (type === 'tEXt' || type === 'iTXt') {
        const s = dec.decode(data);
        const nul = s.indexOf('\0');
        const key = nul < 0 ? type : s.slice(0, nul);
        let val = nul < 0 ? s : s.slice(nul + 1).replace(/^[\0-\x08]+/, '');
        val = val.replace(/\0/g, ' ').trim();
        if (val) out.meta[key] = val.slice(0, 2000);
      } else if (type === 'pHYs') {
        const px = u32(i + 8, false), py = u32(i + 12, false);
        if (data[8] === 1) out.meta['Density'] = `${Math.round(px * 0.0254)}x${Math.round(py * 0.0254)} dpi`;
      } else if (type === 'eXIf') {
        Object.assign(out.meta, parseTIFF(data, 0));
      } else if (type === 'IDAT' && Object.keys(out.meta).length) break;
      if (type === 'IEND') break;
      i += 12 + len;
      if (len < 0 || len > b.length) break;
    }
    return out;
  }
  if (b[0] === 0xff && b[1] === 0xd8) {
    out.format = 'jpeg';
    let i = 2;
    while (i + 4 < b.length) {
      if (b[i] !== 0xff) { i++; continue; }
      const marker = b[i + 1];
      if (marker === 0xd8 || marker === 0x01 || (marker >= 0xd0 && marker <= 0xd7)) { i += 2; continue; }
      const len = u16(i + 2, false);
      const seg = b.subarray(i + 4, i + 2 + len);
      if (marker >= 0xc0 && marker <= 0xcf && marker !== 0xc4 && marker !== 0xc8 && marker !== 0xcc) {
        out.height = u16(i + 5, false);
        out.width = u16(i + 7, false);
      } else if (marker === 0xe1) {
        const head = dec.decode(seg.subarray(0, 6));
        if (head.startsWith('Exif')) Object.assign(out.meta, parseTIFF(seg, 6));
        else if (dec.decode(seg.subarray(0, 28)).includes('ns.adobe.com/xap')) {
          const xmp = dec.decode(seg);
          for (const [k, re] of [
            ['CreatorTool', /CreatorTool>([^<]{1,120})</],
            ['XMP:Software', /xmp:CreatorTool="([^"]{1,120})"/],
            ['XMP:CreateDate', /CreateDate>([^<]{1,60})</],
            ['XMP:DocumentID', /DocumentID>([^<]{1,80})</],
          ]) {
            const m = xmp.match(re);
            if (m) out.meta[k] = m[1];
          }
        }
      } else if (marker === 0xed) {
        const s = dec.decode(seg);
        if (s.startsWith('Photoshop')) out.meta['Photoshop'] = 'IRB present';
      } else if (marker === 0xee) {
        out.meta['Adobe'] = 'APP14 Adobe marker';
      } else if (marker === 0xfe) {
        const s = dec.decode(seg).trim();
        if (s) out.meta['Comment'] = s.slice(0, 300);
      } else if (marker === 0xda) break;
      if (!len) break;
      i += 2 + len;
    }
    return out;
  }
  if (b[0] === 0x47 && b[1] === 0x49) {
    out.format = 'gif';
    out.width = u16(6, true);
    out.height = u16(8, true);
    return out;
  }
  if (b[0] === 0x42 && b[1] === 0x4d) {
    out.format = 'bmp';
    out.width = u32(18, true);
    out.height = u32(22, true);
    return out;
  }
  if ((b[0] === 0x49 && b[1] === 0x49) || (b[0] === 0x4d && b[1] === 0x4d)) {
    out.format = 'tiff';
    const tags = parseTIFF(b, 0);
    Object.assign(out.meta, tags);
    if (tags._width) out.width = tags._width;
    if (tags._height) out.height = tags._height;
    delete out.meta._width;
    delete out.meta._height;
    return out;
  }
  if (u32(0, true) === 0x9ac6cdd7) {
    out.format = 'wmf';
    return out;
  }
  if (u32(0, true) === 0x00000001 && dec.decode(b.subarray(40, 44)) === ' EMF') {
    out.format = 'emf';
    return out;
  }
  if (dec.decode(b.subarray(0, 4)) === 'RIFF') { out.format = 'webp'; return out; }
  if (dec.decode(b.subarray(4, 12)).includes('ftyp')) { out.format = 'heif/avif'; return out; }
  return out;
}

const TIFF_TAGS = {
  256: '_width', 257: '_height',
  271: 'Make', 272: 'Model', 305: 'Software', 306: 'DateTime', 315: 'Artist', 33432: 'Copyright',
  36867: 'DateTimeOriginal', 36868: 'DateTimeDigitized', 42016: 'ImageUniqueID',
  0x9c9b: 'XPTitle', 0x9c9c: 'XPComment', 0x9c9d: 'XPAuthor', 0x9c9e: 'XPKeywords', 0x9c9f: 'XPSubject',
  37510: 'UserComment', 42032: 'CameraOwnerName', 42033: 'BodySerialNumber', 42036: 'LensModel',
  33434: 'ExposureTime', 34855: 'ISO', 37377: 'ShutterSpeed',
};
const GPS_TAGS = { 1: 'GPSLatRef', 2: 'GPSLat', 3: 'GPSLonRef', 4: 'GPSLon', 5: 'GPSAltRef', 6: 'GPSAlt', 29: 'GPSDate' };

function parseTIFF(b, base) {
  const out = {};
  try {
    const le = b[base] === 0x49;
    const g16 = (i) => (le ? b[i] | (b[i + 1] << 8) : (b[i] << 8) | b[i + 1]);
    const g32 = (i) => (le ? (b[i] | (b[i + 1] << 8) | (b[i + 2] << 16) | (b[i + 3] << 24)) >>> 0 : ((b[i] << 24) | (b[i + 1] << 16) | (b[i + 2] << 8) | b[i + 3]) >>> 0);
    if (g16(base + 2) !== 42) return out;
    const readIFD = (off, tagMap, prefix) => {
      if (off + 2 > b.length) return 0;
      const n = g16(base + off);
      if (n > 512) return 0;
      for (let k = 0; k < n; k++) {
        const e = base + off + 2 + k * 12;
        if (e + 12 > b.length) break;
        const tag = g16(e);
        const type = g16(e + 2);
        const cnt = g32(e + 4);
        const sizes = { 1: 1, 2: 1, 3: 2, 4: 4, 5: 8, 7: 1, 9: 4, 10: 8 };
        const sz = (sizes[type] || 1) * cnt;
        const vOff = sz <= 4 ? e + 8 : base + g32(e + 8);
        if (tag === 34665) { readIFD(g32(e + 8), TIFF_TAGS, ''); continue; }
        if (tag === 34853) { readIFD(g32(e + 8), GPS_TAGS, ''); continue; }
        const name = tagMap[tag];
        if (!name) continue;
        let val;
        if (type === 2) val = dec.decode(b.subarray(vOff, vOff + Math.max(0, sz - 1))).replace(/\0+$/, '').trim();
        else if (type === 3) val = g16(vOff);
        else if (type === 4) val = g32(vOff);
        else if (type === 5) {
          const parts = [];
          for (let q = 0; q < Math.min(cnt, 3); q++) {
            const num = g32(vOff + q * 8), den = g32(vOff + q * 8 + 4);
            parts.push(den ? num / den : 0);
          }
          val = parts.length === 3 ? parts : parts[0];
        } else if (type === 1 || type === 7) {
          const raw = b.subarray(vOff, vOff + Math.min(sz, 64));
          val = dec.decode(raw).replace(/[^\x20-\x7e]/g, '').trim();
        }
        if (name.startsWith('XP') && (type === 1 || type === 7)) {
          const raw = b.subarray(vOff, vOff + sz);
          let s = '';
          for (let q = 0; q + 1 < raw.length; q += 2) {
            const c = raw[q] | (raw[q + 1] << 8);
            if (c) s += String.fromCharCode(c);
          }
          val = s.trim();
        }
        if (val !== undefined && val !== '' && !(Array.isArray(val) && !val.length)) out[prefix + name] = val;
      }
      return g32(base + off + 2 + n * 12);
    };
    let ifd = g32(base + 4);
    let guard = 0;
    while (ifd && guard++ < 4) ifd = readIFD(ifd, TIFF_TAGS, '');
    if (out.GPSLat && Array.isArray(out.GPSLat)) {
      const dms = (a) => a[0] + a[1] / 60 + (a[2] || 0) / 3600;
      const lat = dms(out.GPSLat) * (String(out.GPSLatRef).startsWith('S') ? -1 : 1);
      const lon = Array.isArray(out.GPSLon) ? dms(out.GPSLon) * (String(out.GPSLonRef).startsWith('W') ? -1 : 1) : null;
      out.GPS = lon === null ? `${lat.toFixed(6)}` : `${lat.toFixed(6)}, ${lon.toFixed(6)}`;
      delete out.GPSLat; delete out.GPSLon; delete out.GPSLatRef; delete out.GPSLonRef;
    }
    for (const k of Object.keys(out)) if (Array.isArray(out[k])) out[k] = out[k].map((x) => (typeof x === 'number' ? +x.toFixed(4) : x)).join(', ');
  } catch (e) {
    /* metadata is best-effort by nature */
  }
  return out;
}

export function sniffMime(bytes) {
  const f = imageInfo(bytes).format;
  return { png: 'image/png', jpeg: 'image/jpeg', gif: 'image/gif', bmp: 'image/bmp', tiff: 'image/tiff', webp: 'image/webp' }[f] || 'application/octet-stream';
}
