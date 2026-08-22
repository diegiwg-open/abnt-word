const fs = require('fs');
const path = require('path');
const zlib = require('zlib');

function createPng(width, height, drawFn) {
  // RGBA buffer
  const buffer = Buffer.alloc(width * height * 4);
  
  for (let y = 0; y < height; y++) {
    for (let x = 0; x < width; x++) {
      const idx = (y * width + x) * 4;
      const [r, g, b, a] = drawFn(x, y, width, height);
      buffer[idx] = r;
      buffer[idx + 1] = g;
      buffer[idx + 2] = b;
      buffer[idx + 3] = a;
    }
  }

  // PNG creation with raw filter 0 for scanlines
  const scanlines = [];
  for (let y = 0; y < height; y++) {
    scanlines.push(Buffer.from([0])); // filter type 0 (None)
    scanlines.push(buffer.subarray(y * width * 4, (y + 1) * width * 4));
  }
  const rawData = Buffer.concat(scanlines);
  const compressed = zlib.deflateSync(rawData);

  // PNG Header
  const signature = Buffer.from([137, 80, 78, 71, 13, 10, 26, 10]);

  // IHDR chunk
  const ihdrData = Buffer.alloc(13);
  ihdrData.writeUInt32BE(width, 0);
  ihdrData.writeUInt32BE(height, 4);
  ihdrData.writeUInt8(8, 8); // bit depth
  ihdrData.writeUInt8(6, 9); // RGBA
  ihdrData.writeUInt8(0, 10); // compression
  ihdrData.writeUInt8(0, 11); // filter
  ihdrData.writeUInt8(0, 12); // interlace
  const ihdrChunk = createChunk('IHDR', ihdrData);

  // IDAT chunk
  const idatChunk = createChunk('IDAT', compressed);

  // IEND chunk
  const iendChunk = createChunk('IEND', Buffer.alloc(0));

  return Buffer.concat([signature, ihdrChunk, idatChunk, iendChunk]);
}

function createChunk(type, data) {
  const length = data.length;
  const chunk = Buffer.alloc(8 + length + 4);
  chunk.writeUInt32BE(length, 0);
  chunk.write(type, 4, 4, 'ascii');
  data.copy(chunk, 8);

  const crcData = Buffer.concat([Buffer.from(type, 'ascii'), data]);
  const crc = calculateCrc32(crcData);
  chunk.writeInt32BE(crc, 8 + length);
  return chunk;
}

// Standard CRC32 table
const crcTable = [];
for (let n = 0; n < 256; n++) {
  let c = n;
  for (let k = 0; k < 8; k++) {
    c = (c & 1) ? (0xedb88320 ^ (c >>> 1)) : (c >>> 1);
  }
  crcTable[n] = c;
}

function calculateCrc32(buf) {
  let crc = -1;
  for (let i = 0; i < buf.length; i++) {
    crc = (crc >>> 8) ^ crcTable[(crc ^ buf[i]) & 0xff];
  }
  return crc ^ (-1);
}

// Drawing function for ABNT Assistant icon: Elegant gradient blue square with rounded corners, book/document icon & gold badge
function drawIcon(x, y, w, h) {
  const nx = x / w;
  const ny = y / h;
  const cornerRadius = 0.18;

  // Check rounded corner bounding box
  const cx = Math.min(Math.max(nx, cornerRadius), 1 - cornerRadius);
  const cy = Math.min(Math.max(ny, cornerRadius), 1 - cornerRadius);
  const dist = Math.hypot(nx - cx, ny - cy);

  if (dist > cornerRadius) {
    return [0, 0, 0, 0]; // Transparent outside
  }

  // Background Gradient (Deep Royal Blue to Indigo)
  const rBg = Math.floor(26 + (ny * 15));
  const gBg = Math.floor(54 + (ny * 40));
  const bBg = Math.floor(138 + (ny * 60));

  // Border highlight
  const isBorder = (nx < 0.04 || nx > 0.96 || ny < 0.04 || ny > 0.96);
  if (isBorder && dist <= cornerRadius) {
    return [59, 130, 246, 255];
  }

  // Draw Document / Book Sheet
  const docLeft = 0.22, docRight = 0.78, docTop = 0.18, docBottom = 0.82;
  if (nx >= docLeft && nx <= docRight && ny >= docTop && ny <= docBottom) {
    // Top right fold effect
    if (nx > 0.62 && ny < 0.34 && (nx - 0.62) + (0.34 - ny) > 0.20) {
      // Background shows through the cut corner or fold
      return [rBg, gBg, bBg, 255];
    }

    // Document lines (horizontal lines representing formatting)
    const line1 = ny >= 0.40 && ny <= 0.44 && nx >= 0.30 && nx <= 0.70;
    const line2 = ny >= 0.50 && ny <= 0.54 && nx >= 0.30 && nx <= 0.65;
    const line3 = ny >= 0.60 && ny <= 0.64 && nx >= 0.30 && nx <= 0.70;
    const line4 = ny >= 0.70 && ny <= 0.74 && nx >= 0.30 && nx <= 0.55;

    if (line1 || line2 || line3 || line4) {
      return [16, 185, 129, 255]; // Emerald green ABNT formatted lines
    }

    // Header ABNT badge accent
    if (ny >= 0.24 && ny <= 0.32 && nx >= 0.30 && nx <= 0.55) {
      return [59, 130, 246, 255]; // Blue header
    }

    // Document Paper Body
    return [255, 255, 255, 255];
  }

  // Bottom Gold/Emerald Accent Dot
  const dotDist = Math.hypot(nx - 0.78, ny - 0.78);
  if (dotDist < 0.12) {
    return [245, 158, 11, 255]; // Amber Gold seal
  }

  return [rBg, gBg, bBg, 255];
}

const assetsDir = path.join(__dirname, '..', 'src', 'assets');
if (!fs.existsSync(assetsDir)) {
  fs.mkdirSync(assetsDir, { recursive: true });
}

[16, 32, 64, 80, 128].forEach(size => {
  const png = createPng(size, size, drawIcon);
  fs.writeFileSync(path.join(assetsDir, `icon-${size}.png`), png);
  console.log(`Generated icon-${size}.png`);
});
