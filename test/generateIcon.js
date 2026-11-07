const fs = require('fs');
const path = require('path');
const zlib = require('zlib');

function createCrcTable() {
  const table = new Uint32Array(256);
  for (let n = 0; n < 256; n++) {
    let c = n;
    for (let k = 0; k < 8; k++) {
      c = (c & 1) ? (0xedb88320 ^ (c >>> 1)) : (c >>> 1);
    }
    table[n] = c;
  }
  return table;
}

const crcTable = createCrcTable();

function calculateCrc(buf) {
  let crc = 0xffffffff;
  for (let i = 0; i < buf.length; i++) {
    crc = crcTable[(crc ^ buf[i]) & 0xff] ^ (crc >>> 8);
  }
  return (crc ^ 0xffffffff) >>> 0;
}

function makeChunk(type, data) {
  const len = data.length;
  const chunk = Buffer.alloc(12 + len);
  chunk.writeUInt32BE(len, 0);
  chunk.write(type, 4, 4, 'ascii');
  data.copy(chunk, 8);
  const typeAndData = chunk.subarray(4, 8 + len);
  const crc = calculateCrc(typeAndData);
  chunk.writeUInt32BE(crc, 8 + len);
  return chunk;
}

function generateIconPng() {
  const width = 128;
  const height = 128;

  // Raw image buffer with filter byte at beginning of each scanline
  const rawData = Buffer.alloc((width * 4 + 1) * height);

  for (let y = 0; y < height; y++) {
    const rowOffset = y * (width * 4 + 1);
    rawData[rowOffset] = 0; // Filter: None

    for (let x = 0; x < width; x++) {
      const pixelOffset = rowOffset + 1 + x * 4;

      // Rounded rectangle bounds (radius: 20px)
      const r = 20;
      let inBounds = true;
      if (x < r && y < r && (x - r) ** 2 + (y - r) ** 2 > r ** 2) inBounds = false;
      if (x > width - r && y < r && (x - (width - r)) ** 2 + (y - r) ** 2 > r ** 2) inBounds = false;
      if (x < r && y > height - r && (x - r) ** 2 + (y - (height - r)) ** 2 > r ** 2) inBounds = false;
      if (x > width - r && y > height - r && (x - (width - r)) ** 2 + (y - (height - r)) ** 2 > r ** 2) inBounds = false;

      if (!inBounds) {
        rawData[pixelOffset] = 0;
        rawData[pixelOffset + 1] = 0;
        rawData[pixelOffset + 2] = 0;
        rawData[pixelOffset + 3] = 0;
        continue;
      }

      // Background Gradient: Deep Slate Blue (#1a2233 to #0d131f)
      const gradRatio = (x + y) / (width + height);
      let red = Math.round(26 * (1 - gradRatio) + 13 * gradRatio);
      let green = Math.round(34 * (1 - gradRatio) + 19 * gradRatio);
      let blue = Math.round(51 * (1 - gradRatio) + 31 * gradRatio);
      let alpha = 255;

      // Draw Document Page (White Sheet with folded corner)
      // Doc boundaries: x: 30 to 98, y: 22 to 106
      const docX1 = 30, docX2 = 98, docY1 = 22, docY2 = 106;
      const foldSize = 20;

      if (x >= docX1 && x <= docX2 && y >= docY1 && y <= docY2) {
        // Top right folded corner check
        const inCorner = x > docX2 - foldSize && y < docY1 + foldSize;
        const cornerCut = (x - (docX2 - foldSize)) + (docY1 + foldSize - y) > foldSize;

        if (!(inCorner && cornerCut)) {
          // Document surface
          red = 245; green = 247; blue = 250;

          // Corner flap
          if (inCorner && !cornerCut) {
            red = 180; green = 190; blue = 205;
          }

          // Document content decorative lines (Word blue & Excel green)
          // Line 1: Word Blue (#2b579a)
          if (y >= 44 && y <= 50 && x >= 42 && x <= 72) {
            red = 43; green = 87; blue = 154;
          }
          // Line 2: Excel Green (#217346)
          if (y >= 58 && y <= 64 && x >= 42 && x <= 86) {
            red = 33; green = 115; blue = 70;
          }
          // Line 3: PDF Red (#d9381e)
          if (y >= 72 && y <= 78 && x >= 42 && x <= 80) {
            red = 217; green = 56; blue = 30;
          }
          // Grid dots / table accent
          if (y >= 86 && y <= 94 && x >= 42 && x <= 86) {
            if ((x - 42) % 11 < 7) {
              red = 140; green = 155; blue = 175;
            }
          }
        }
      }

      // Border outline of the icon
      if (x === 0 || x === width - 1 || y === 0 || y === height - 1) {
        red = 60; green = 80; blue = 120;
      }

      rawData[pixelOffset] = red;
      rawData[pixelOffset + 1] = green;
      rawData[pixelOffset + 2] = blue;
      rawData[pixelOffset + 3] = alpha;
    }
  }

  // PNG Signature
  const signature = Buffer.from([137, 80, 78, 71, 13, 10, 26, 10]);

  // IHDR Chunk
  const ihdr = Buffer.alloc(13);
  ihdr.writeUInt32BE(width, 0);
  ihdr.writeUInt32BE(height, 4);
  ihdr[8] = 8; // bit depth
  ihdr[9] = 6; // color type RGBA
  ihdr[10] = 0; // compression
  ihdr[11] = 0; // filter
  ihdr[12] = 0; // interlace
  const ihdrChunk = makeChunk('IHDR', ihdr);

  // IDAT Chunk
  const compressedData = zlib.deflateSync(rawData);
  const idatChunk = makeChunk('IDAT', compressedData);

  // IEND Chunk
  const iendChunk = makeChunk('IEND', Buffer.alloc(0));

  const pngBuffer = Buffer.concat([signature, ihdrChunk, idatChunk, iendChunk]);

  const assetsDir = path.join(__dirname, '..', 'assets');
  if (!fs.existsSync(assetsDir)) {
    fs.mkdirSync(assetsDir, { recursive: true });
  }

  const outputPath = path.join(assetsDir, 'icon.png');
  fs.writeFileSync(outputPath, pngBuffer);
  console.log('Icon PNG oluşturuldu:', outputPath, `(${pngBuffer.length} bytes)`);
}

generateIconPng();
