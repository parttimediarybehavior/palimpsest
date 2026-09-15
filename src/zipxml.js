// zipxml.js — dependency-free ZIP reader + XML parser.
// Runs identically in browsers and Node (the inflater is injected).

export const DEFAULT_INFLATE =
  typeof DecompressionStream !== 'undefined'
    ? async (bytes) => {
        const ds = new DecompressionStream('deflate-raw');
        const w = ds.writable.getWriter();
        w.write(bytes);
        w.close();
        return new Uint8Array(await new Response(ds.readable).arrayBuffer());
      }
    : null;

const td = new TextDecoder('utf-8');

/* ---------------------------------------------------------------- ZIP ---- */

// DOS date/time (as used in local + central headers) -> ISO string or null.
function dosTime(dt, dd) {
  if (!dt && !dd) return null;
  const sec = (dt & 0x1f) * 2;
  const min = (dt >> 5) & 0x3f;
  const hour = (dt >> 11) & 0x1f;
  const day = dd & 0x1f;
  const month = (dd >> 5) & 0x0f;
  const year = ((dd >> 9) & 0x7f) + 1980;
  if (!day || !month) return null;
  const pad = (n) => String(n).padStart(2, '0');
  return `${year}-${pad(month)}-${pad(day)}T${pad(hour)}:${pad(min)}:${pad(sec)}`;
}

/**
 * Read a ZIP archive from raw bytes.
 * @returns {Promise<{entries: Map<string, Entry>, order: string[], comment: string}>}
 * Entry = { name, bytes, size, compressedSize, method, crc, mtime, offset, versionMadeBy, flags, extraIds }
 */
export async function readZip(buf, inflate = DEFAULT_INFLATE) {
  const u8 = buf instanceof Uint8Array ? buf : new Uint8Array(buf);
  if (u8.length >= 8 && u8[0] === 0xd0 && u8[1] === 0xcf && u8[2] === 0x11 && u8[3] === 0xe0) {
    throw new Error('this is an OLE container, not a .docx zip — either a legacy .doc or a password-encrypted Office file');
  }
  if (u8.length < 22) throw new Error('file is too short to be a .docx');
  const dv = new DataView(u8.buffer, u8.byteOffset, u8.byteLength);
  // locate End Of Central Directory
  let eocd = -1;
  for (let i = u8.length - 22; i >= Math.max(0, u8.length - 66000); i--) {
    if (dv.getUint32(i, true) === 0x06054b50) { eocd = i; break; }
  }
  if (eocd < 0) throw new Error('not a zip file (no end-of-central-directory record)');
  const count = dv.getUint16(eocd + 10, true);
  let ptr = dv.getUint32(eocd + 16, true);
  const commentLen = dv.getUint16(eocd + 20, true);
  const comment = commentLen ? td.decode(u8.subarray(eocd + 22, eocd + 22 + commentLen)) : '';

  const entries = new Map();
  const order = [];
  for (let i = 0; i < count; i++) {
    if (ptr + 46 > u8.length || dv.getUint32(ptr, true) !== 0x02014b50) break;
    const versionMadeBy = dv.getUint16(ptr + 4, true);
    const flags = dv.getUint16(ptr + 8, true);
    const method = dv.getUint16(ptr + 10, true);
    const mtime = dosTime(dv.getUint16(ptr + 12, true), dv.getUint16(ptr + 14, true));
    const crc = dv.getUint32(ptr + 16, true);
    const compressedSize = dv.getUint32(ptr + 20, true);
    const size = dv.getUint32(ptr + 24, true);
    const nameLen = dv.getUint16(ptr + 28, true);
    const extraLen = dv.getUint16(ptr + 30, true);
    const cmtLen = dv.getUint16(ptr + 32, true);
    const offset = dv.getUint32(ptr + 42, true);
    const name = td.decode(u8.subarray(ptr + 46, ptr + 46 + nameLen));
    const extra = u8.subarray(ptr + 46 + nameLen, ptr + 46 + nameLen + extraLen);
    const extraIds = [];
    for (let j = 0; j + 4 <= extra.length; ) {
      const id = extra[j] | (extra[j + 1] << 8);
      const len = extra[j + 2] | (extra[j + 3] << 8);
      extraIds.push(id);
      j += 4 + len;
    }
    const e = { name, size, compressedSize, method, crc, mtime, offset, versionMadeBy, flags, extraIds, _u8: u8, _bytes: null };
    entries.set(name, e);
    order.push(name);
    ptr += 46 + nameLen + extraLen + cmtLen;
  }

  const getBytes = async (e) => {
    if (e._bytes) return e._bytes;
    const dv2 = new DataView(u8.buffer, u8.byteOffset, u8.byteLength);
    if (e.offset + 30 > u8.length || dv2.getUint32(e.offset, true) !== 0x04034b50) throw new Error('corrupt zip: bad local header for ' + e.name);
    const nameLen = dv2.getUint16(e.offset + 26, true);
    const extraLen = dv2.getUint16(e.offset + 28, true);
    const start = e.offset + 30 + nameLen + extraLen;
    const raw = u8.subarray(start, start + (e.compressedSize || 0));
    if (e.method === 0) e._bytes = raw;
    else if (e.method === 8) {
      if (!inflate) throw new Error('no inflater available');
      e._bytes = await inflate(raw);
    } else throw new Error('unsupported compression method ' + e.method + ' for ' + e.name);
    return e._bytes;
  };

  return {
    entries,
    order,
    comment,
    has: (n) => entries.has(n),
    bytes: async (n) => (entries.has(n) ? await getBytes(entries.get(n)) : null),
    text: async (n) => {
      const b = entries.has(n) ? await getBytes(entries.get(n)) : null;
      return b ? td.decode(b) : null;
    },
  };
}

/* ---------------------------------------------------------------- XML ---- */

const ENT = { amp: '&', lt: '<', gt: '>', quot: '"', apos: "'", nbsp: '\u00a0' };
export function unesc(s) {
  if (s.indexOf('&') < 0) return s;
  return s.replace(/&(#x?[0-9a-fA-F]+|[a-zA-Z]+);/g, (m, g) => {
    if (g[0] === '#') return String.fromCodePoint(parseInt(g[1] === 'x' || g[1] === 'X' ? g.slice(2) : g.slice(1), g[1] === 'x' || g[1] === 'X' ? 16 : 10));
    return ENT[g] !== undefined ? ENT[g] : m;
  });
}

/**
 * Minimal, fast, allocation-light XML parser.
 * Node: { tag, attrs, kids }  — kids hold child Nodes and raw text strings.
 */
export function parseXML(src) {
  const root = { tag: '#root', attrs: {}, kids: [] };
  const stack = [root];
  let i = 0;
  const n = src.length;
  while (i < n) {
    const lt = src.indexOf('<', i);
    if (lt < 0) break;
    if (lt > i) {
      const txt = src.slice(i, lt);
      if (txt.trim() || /[ \t\n]/.test(txt)) stack[stack.length - 1].kids.push(unesc(txt));
    }
    if (src.startsWith('<!--', lt)) { i = src.indexOf('-->', lt) + 3 || n; continue; }
    if (src.startsWith('<![CDATA[', lt)) {
      const end = src.indexOf(']]>', lt);
      stack[stack.length - 1].kids.push(src.slice(lt + 9, end < 0 ? n : end));
      i = end < 0 ? n : end + 3;
      continue;
    }
    if (src.startsWith('<?', lt) || src.startsWith('<!', lt)) { i = src.indexOf('>', lt) + 1 || n; continue; }
    if (src[lt + 1] === '/') {
      const end = src.indexOf('>', lt);
      if (stack.length > 1) stack.pop();
      i = end + 1;
      continue;
    }
    // open tag: scan name then attributes, respecting quotes
    let j = lt + 1;
    while (j < n && !/[\s/>]/.test(src[j])) j++;
    const tag = src.slice(lt + 1, j);
    const node = { tag, attrs: {}, kids: [] };
    let selfClose = false;
    while (j < n) {
      while (j < n && /\s/.test(src[j])) j++;
      if (src[j] === '/') { selfClose = true; j++; continue; }
      if (src[j] === '>') { j++; break; }
      let k = j;
      while (k < n && src[k] !== '=' && src[k] !== '>' && !/\s/.test(src[k])) k++;
      const aname = src.slice(j, k);
      let aval = '';
      if (src[k] === '=') {
        const q = src[k + 1];
        if (q === '"' || q === "'") {
          const end = src.indexOf(q, k + 2);
          aval = src.slice(k + 2, end < 0 ? n : end);
          k = (end < 0 ? n : end) + 1;
        } else {
          let e2 = k + 1;
          while (e2 < n && !/[\s>]/.test(src[e2])) e2++;
          aval = src.slice(k + 1, e2);
          k = e2;
        }
      }
      if (aname) node.attrs[aname] = unesc(aval);
      j = k;
      if (j >= n) break;
    }
    stack[stack.length - 1].kids.push(node);
    if (!selfClose) stack.push(node);
    i = j;
  }
  return root;
}

/* ------------------------------------------------------- tree helpers ---- */

// iterative: real documents nest table cells hundreds deep, and fuzzed ones worse
export function* walk(node) {
  const stack = [[node, 0]];
  while (stack.length) {
    const top = stack[stack.length - 1];
    const [n, i] = top;
    if (i >= n.kids.length) {
      stack.pop();
      continue;
    }
    top[1]++;
    const k = n.kids[i];
    if (typeof k === 'object') {
      yield k;
      stack.push([k, 0]);
    }
  }
}
export function children(node, tag) {
  const out = [];
  for (const k of node.kids) if (typeof k === 'object' && k.tag === tag) out.push(k);
  return out;
}
export function child(node, tag) {
  for (const k of node.kids) if (typeof k === 'object' && k.tag === tag) return k;
  return null;
}
export function find(node, tag) {
  for (const k of walk(node)) if (k.tag === tag) return k;
  return null;
}
export function findAll(node, tag) {
  const out = [];
  for (const k of walk(node)) if (k.tag === tag) out.push(k);
  return out;
}
export function textOf(node) {
  let s = '';
  const stack = [[node, 0]];
  while (stack.length) {
    const top = stack[stack.length - 1];
    const [n, i] = top;
    if (i >= n.kids.length) {
      stack.pop();
      continue;
    }
    top[1]++;
    const k = n.kids[i];
    if (typeof k === 'string') s += k;
    else stack.push([k, 0]);
  }
  return s;
}
