/**
 * Generates solid-color PNG icon files for the add-in manifest.
 * Uses only Node.js built-ins (zlib + fs) — no npm dependencies.
 * Output: src/taskpane/icon-{16,32,80}.png
 */

'use strict';

const zlib = require('zlib');
const fs = require('fs');
const path = require('path');

// CRC-32 lookup table (standard polynomial 0xEDB88320)
const crcTable = (function () {
  const t = new Uint32Array(256);
  for (let n = 0; n < 256; n++) {
    let c = n;
    for (let k = 0; k < 8; k++) {
      c = c & 1 ? 0xEDB88320 ^ (c >>> 1) : c >>> 1;
    }
    t[n] = c;
  }
  return t;
})();

function crc32(buf) {
  let crc = 0xFFFFFFFF;
  for (let i = 0; i < buf.length; i++) {
    crc = crcTable[(crc ^ buf[i]) & 0xFF] ^ (crc >>> 8);
  }
  return (~crc) >>> 0;
}

function pngChunk(type, data) {
  const typeBytes = Buffer.from(type, 'ascii');
  const crcInput = Buffer.concat([typeBytes, data]);
  const crcVal = crc32(crcInput);
  const out = Buffer.alloc(4 + 4 + data.length + 4);
  out.writeUInt32BE(data.length, 0);
  typeBytes.copy(out, 4);
  data.copy(out, 8);
  out.writeUInt32BE(crcVal, 8 + data.length);
  return out;
}

function makeSolidPNG(size, r, g, b) {
  // PNG signature
  const sig = Buffer.from([137, 80, 78, 71, 13, 10, 26, 10]);

  // IHDR: width, height, bit-depth=8, color-type=2 (RGB), compress=0, filter=0, interlace=0
  const ihdrData = Buffer.alloc(13);
  ihdrData.writeUInt32BE(size, 0);
  ihdrData.writeUInt32BE(size, 4);
  ihdrData[8] = 8;
  ihdrData[9] = 2;
  ihdrData.writeUInt8(0, 10);
  ihdrData.writeUInt8(0, 11);
  ihdrData.writeUInt8(0, 12);
  const ihdr = pngChunk('IHDR', ihdrData);

  // Raw image data: one filter byte (0 = None) + RGB pixels per row
  const rowBytes = 1 + size * 3;
  const raw = Buffer.alloc(size * rowBytes);
  for (let y = 0; y < size; y++) {
    raw[y * rowBytes] = 0; // filter type: None
    for (let x = 0; x < size; x++) {
      raw[y * rowBytes + 1 + x * 3] = r;
      raw[y * rowBytes + 1 + x * 3 + 1] = g;
      raw[y * rowBytes + 1 + x * 3 + 2] = b;
    }
  }
  const compressed = zlib.deflateSync(raw, { level: 9 });
  const idat = pngChunk('IDAT', compressed);

  // IEND
  const iend = pngChunk('IEND', Buffer.alloc(0));

  return Buffer.concat([sig, ihdr, idat, iend]);
}

// Campus blue #003E7E = rgb(0, 62, 126)
const R = 0, G = 62, B = 126;

const outDir = path.join(__dirname, '..', 'src', 'taskpane');

[16, 32, 80].forEach(function (size) {
  const png = makeSolidPNG(size, R, G, B);
  const outPath = path.join(outDir, 'icon-' + size + '.png');
  fs.writeFileSync(outPath, png);
  console.log('Created ' + outPath + ' (' + png.length + ' bytes)');
});
