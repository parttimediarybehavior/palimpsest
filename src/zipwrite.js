// zipwrite.js — dependency-free ZIP writer (the other half of zipxml.js).
// Browsers get deflate from CompressionStream; Node injects zlib.deflateRawSync.
// With no deflater at all it stores — still a valid archive Word will open.

const NAME_ENC = new TextEncoder();

const CRC_TABLE = (() => {
  const t = new Uint32Array(256);
  for (let n = 0; n < 256; n++) {
    let c = n;
    for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
    t[n] = c >>> 0;
  }
  return t;
})();

export function crc32(buf) {
  let crc = 0xffffffff;
  for (let i = 0; i < buf.length; i++) crc = CRC_TABLE[(crc ^ buf[i]) & 0xff] ^ (crc >>> 8);
  return (crc ^ 0xffffffff) >>> 0;
}

export const DEFAULT_DEFLATE =
  typeof CompressionStream !== 'undefined'
    ? async (bytes) => {
        try {
          const cs = new CompressionStream('deflate-raw');
          const w = cs.writable.getWriter();
          w.write(bytes);
          w.close();
          const buf = await new Response(cs.readable).arrayBuffer();
          return new Uint8Array(buf);
        } catch {
          return null; // no deflate-raw on this runtime -> caller stores
        }
      }
    : null;

// DOS date/time. Word itself writes 1980-01-01 00:00 into every entry, which is
// both a deliberate privacy choice and convenient: nothing to leak here.
const DOS_DATE = 0x0021;
const DOS_TIME = 0x0000;

/**
 * Build a ZIP archive.
 * @param {Array<{name: string, data: Uint8Array, store?: boolean}>} files
 * @param {{deflate?: (b: Uint8Array) => Promise<Uint8Array|null>}} [opts]
 * @returns {Promise<Uint8Array>}
 */
export async function writeZip(files, opts = {}) {
  const deflate = opts.deflate === undefined ? DEFAULT_DEFLATE : opts.deflate;
  const chunks = [];
  const central = [];
  let offset = 0;
  const push = (arr) => {
    chunks.push(arr);
    offset += arr.length;
  };

  for (const f of files) {
    const raw = f.data instanceof Uint8Array ? f.data : new Uint8Array(f.data);
    let comp = null;
    if (!f.store && deflate && raw.length > 64) comp = await deflate(raw);
    // never let "compression" inflate a part
    const stored = !comp || comp.length >= raw.length;
    if (stored) comp = raw;
    const nameB = NAME_ENC.encode(f.name);
    const crc = crc32(raw);

    const lh = new Uint8Array(30 + nameB.length);
    const dv = new DataView(lh.buffer);
    dv.setUint32(0, 0x04034b50, true);
    dv.setUint16(4, 20, true);
    dv.setUint16(8, stored ? 0 : 8, true);
    dv.setUint16(10, DOS_TIME, true);
    dv.setUint16(12, DOS_DATE, true);
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
    cdv.setUint16(4, 20, true); // made by
    cdv.setUint16(6, 20, true); // needs 2.0
    cdv.setUint16(10, stored ? 0 : 8, true);
    cdv.setUint16(12, DOS_TIME, true);
    cdv.setUint16(14, DOS_DATE, true);
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

  const out = new Uint8Array(offset);
  let p = 0;
  for (const c of chunks) {
    out.set(c, p);
    p += c.length;
  }
  return out;
}
