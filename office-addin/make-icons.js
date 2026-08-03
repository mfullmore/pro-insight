// Generates minimal, valid solid-color PNG icons for the manifest (no
// external image libraries available in this environment). Brass (#a97a2e)
// on transparent, matching the rest of the product's ledger palette.
const fs = require('fs');
const path = require('path');
const zlib = require('zlib');

function crc32(buf) {
  let c;
  const table = crc32.table || (crc32.table = (() => {
    const t = [];
    for (let n = 0; n < 256; n++) {
      c = n;
      for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
      t[n] = c >>> 0;
    }
    return t;
  })());
  c = 0xffffffff;
  for (let i = 0; i < buf.length; i++) c = table[(c ^ buf[i]) & 0xff] ^ (c >>> 8);
  return (c ^ 0xffffffff) >>> 0;
}

function chunk(type, data) {
  const len = Buffer.alloc(4);
  len.writeUInt32BE(data.length, 0);
  const typeData = Buffer.concat([Buffer.from(type, 'ascii'), data]);
  const crc = Buffer.alloc(4);
  crc.writeUInt32BE(crc32(typeData), 0);
  return Buffer.concat([len, typeData, crc]);
}

function makePng(size, [r, g, b]) {
  const sig = Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]);
  const ihdr = Buffer.alloc(13);
  ihdr.writeUInt32BE(size, 0);
  ihdr.writeUInt32BE(size, 4);
  ihdr[8] = 8; // bit depth
  ihdr[9] = 6; // color type RGBA
  ihdr[10] = 0; ihdr[11] = 0; ihdr[12] = 0;

  // Solid fill with a simple rounded-corner-ish inset so it doesn't look
  // like a bare swatch: full color square with a slightly darker 1px border.
  const raw = Buffer.alloc((size * 4 + 1) * size);
  for (let y = 0; y < size; y++) {
    const rowStart = y * (size * 4 + 1);
    raw[rowStart] = 0; // filter type: none
    for (let x = 0; x < size; x++) {
      const isEdge = x === 0 || y === 0 || x === size - 1 || y === size - 1;
      const off = rowStart + 1 + x * 4;
      raw[off] = isEdge ? Math.max(0, r - 30) : r;
      raw[off + 1] = isEdge ? Math.max(0, g - 30) : g;
      raw[off + 2] = isEdge ? Math.max(0, b - 30) : b;
      raw[off + 3] = 255;
    }
  }
  const idat = zlib.deflateSync(raw);
  return Buffer.concat([sig, chunk('IHDR', ihdr), chunk('IDAT', idat), chunk('IEND', Buffer.alloc(0))]);
}

const BRASS = [0xa9, 0x7a, 0x2e];
const sizes = [16, 32, 64, 80, 128];
const outDir = path.join(__dirname, 'assets');
sizes.forEach((size) => {
  const png = makePng(size, BRASS);
  fs.writeFileSync(path.join(outDir, `icon-${size}.png`), png);
});
console.log(`Wrote icons: ${sizes.map((s) => `icon-${s}.png`).join(', ')}`);
