// Product assertions stay here; browser installation, launch and lifetime belong to HtmlTinkerX.
async function runRichPdfContracts({ moduleBase, fixtures, drawingFont, classic = false }) {
  const api = classic ? globalThis.OfficeIMO : await import(moduleBase + '/pdf/index.js');
  const { writePdf, writePdfDocument, writePdfDocumentTo, pdfTable, PdfFont } = api;
  const { Crc32 } = classic ? api.zip : await import(moduleBase + '/zip/index.js');
  const checks = [];
  const require = (value, message) => { if (!value) throw Error(message); };
  const bytes = base64 => Uint8Array.from(atob(base64), character => character.charCodeAt(0));
  const join = parts => { const output = new Uint8Array(parts.reduce((size, part) => size + part.length, 0)); let position = 0; for (const part of parts) { output.set(part, position); position += part.length; } return output; };
  function chunk(type, data = new Uint8Array()) {
    const name = new TextEncoder().encode(type), crc = new Crc32(); crc.update(name); crc.update(data);
    const result = new Uint8Array(data.length + 12), view = new DataView(result.buffer);
    view.setUint32(0, data.length); result.set(name, 4); result.set(data, 8); view.setUint32(result.length - 4, crc.value); return result;
  }
  function latin(data) { let text = ''; for (let offset = 0; offset < data.length; offset += 8192) text += String.fromCharCode(...data.subarray(offset, offset + 8192)); return text; }
  async function inspect(blob) {
    const data = new Uint8Array(await blob.arrayBuffer()), text = latin(data), images = [];
    require(text.startsWith('%PDF-1.7\n') && text.endsWith('%%EOF\n'), 'Invalid PDF envelope');
    for (const match of text.matchAll(/\d+ 0 obj\n(<<[^\n]*\/Subtype \/Image[^\n]*)\nstream\n/g)) {
      const length = Number(/\/Length (\d+)/.exec(match[1])[1]), start = match.index + match[0].length;
      images.push({ body: match[1], data: data.subarray(start, start + length) });
    }
    return { data, text, images };
  }
  async function save(name, blob, required = []) {
    const data = new Uint8Array(await blob.arrayBuffer()); let encoded = '';
    for (let offset = 0; offset < data.length; offset += 8192) encoded += String.fromCharCode(...data.subarray(offset, offset + 8192));
    await writeRichPdfFixture((classic ? 'classic-' : '') + name + '.pdf', btoa(encoded), JSON.stringify({ required }));
  }
  async function rejected(blocks, options, expected, label) {
    let caught;
    try { await writePdfDocument(blocks, options); } catch (error) { caught = error; }
    require(caught && expected(caught), 'Expected rejection: ' + label + '; got ' + caught); checks.push(label);
  }
  const graphic = data => ({ kind: 'image', image: { format: 'png', data } });
  const configuration = { pageNumbers: false, compression: false };
  let tagReads = 0;
  const row = Object.defineProperty({ value: 17 }, Symbol.toStringTag, { get() { tagReads++; throw Error('Ordinary row string-tag getter was invoked'); } });
  const rowColumns = [{ header: 'Value', key: 'value' }];
  await save('row-tag-getter', await writePdf([row], { ...configuration, columns: rowColumns }), ['Value','17']);
  await save('composed-row-tag-getter', await writePdfDocument([pdfTable([row], { columns: rowColumns })], configuration), ['Value','17']);
  require(tagReads === 0, 'Ordinary rows must not invoke Symbol.toStringTag getters'); checks.push('ordinary-row-symbol-tag-getter');
  const colors = [[240,20,20], [20,230,20], [20,20,235], [235,220,20]], alphas = [255,128,0,64];
  for (const file of ['rgb-pillow.png', 'rgba-pillow.png', 'rgba-filters.png']) {
    const data = bytes(fixtures[file]), block = graphic(data);
    const blob = await writePdfDocument([block, block], { ...configuration, limits: { maxImageBytes: data.length } });
    const pdf = await inspect(blob), rgb = pdf.images.find(image => image.body.includes('/DeviceRGB')), alpha = pdf.images.find(image => image.body.includes('/DeviceGray'));
    const expected = [], expectedAlpha = [];
    for (let y = 0; y < (file === 'rgba-filters.png' ? 5 : 8); y++) for (let x = 0; x < (file === 'rgba-filters.png' ? 4 : 12); x++) {
      if (file === 'rgba-filters.png') { expected.push((x*67+y*29)%256, (x*13+y*83)%256, (x*113+y*7)%256); expectedAlpha.push((x*51+y*37)%256); }
      else { const q = Math.floor(y/4)*2+Math.floor(x/6); expected.push(...colors[q]); expectedAlpha.push(alphas[q]); }
    }
    require(rgb && rgb.data.length === expected.length && expected.every((value,index) => value === rgb.data[index]), file + ' changed RGB samples');
    require(pdf.images.length === (file.startsWith('rgba') ? 2 : 1), 'Image cache did not reuse PNG resources');
    if (file.startsWith('rgba')) require(alpha && expectedAlpha.every((value,index) => value === alpha.data[index]), file + ' changed straight alpha');
    checks.push(file + '-samples-and-cache');
    await save(file.replace('.png',''), blob);
  }
  const rgb = bytes(fixtures['rgb-pillow.png']), signature = rgb.subarray(0,8), header = rgb.subarray(8,33);
  for (const [name, block] of [
    ['oversized-png', { ...graphic(rgb), width: 50000, height: 50000 }],
    ['implicit-panorama', { kind: 'image', image: { format: 'rgb', width: 20000, height: 1, data: new Uint8Array(60000) } }]
  ]) {
    const blob = await writePdfDocument([block], { ...configuration, pageSize: 'A5' }), pdf = await inspect(blob);
    const transform = /q ([\d.]+) 0 0 ([\d.]+) 36 [\d.]+ cm \/Image1 Do/.exec(pdf.text);
    require(transform && Number(transform[1]) > 0 && Number(transform[1]) <= 347.528 && Number(transform[2]) > 0 && Number(transform[2]) <= 523.276,
      name + ' was not fitted to the printable page');
    require((pdf.text.match(/\/Type \/Page /g) ?? []).length === 1, name + ' emitted an extra blank page');
    if (name === 'oversized-png') await save(name, blob);
  }
  checks.push('oversized-and-implicit-image-fit');
  for (const type of ['sRGB','gAMA','cHRM','iCCP','cICP','mDCV','cLLI','sBIT','tRNS','eXIf','acTL','ABCD'])
    await rejected([graphic(join([rgb.subarray(0,33),chunk(type),rgb.subarray(33)]))], configuration,
      error => error.code === 'NOT_SUPPORTED' && error.message.includes(type), 'reject-' + type);
  for (const [offset,value,label] of [[8,16,'16-bit'],[9,3,'indexed'],[12,1,'interlaced']]) {
    const changed = header.subarray(8,21).slice(); changed[offset] = value;
    await rejected([graphic(join([signature,chunk('IHDR',changed),rgb.subarray(33)]))], configuration,
      error => error.code === 'NOT_SUPPORTED', 'reject-' + label);
  }
  const corrupt = rgb.slice(); corrupt[corrupt.length-1] ^= 1;
  await rejected([graphic(corrupt)], configuration, error => String(error).includes('CRC'), 'reject-crc');
  await rejected([graphic(rgb)], { ...configuration, limits: { maxImageBytes: rgb.length-1 } }, error => error.code === 'RESOURCE_LIMIT', 'source-byte-limit');
  await rejected([graphic(rgb)], { ...configuration, limits: { maxImagePixels: 95 } }, error => error.code === 'RESOURCE_LIMIT', 'decoded-pixel-limit');
  const tiny = header.subarray(8,21).slice(); new DataView(tiny.buffer).setUint32(0,1); new DataView(tiny.buffer).setUint32(4,1);
  const inflated = new Uint8Array(4*1024*1024);
  const compressed = new Uint8Array(await new Response(new Blob([inflated]).stream().pipeThrough(new CompressionStream('deflate'))).arrayBuffer());
  await rejected([graphic(join([signature,chunk('IHDR',tiny),chunk('IDAT',compressed),chunk('IEND')]))], configuration,
    error => String(error).includes('exact IHDR scanline length'), 'inflated-length-limit');
  const native = globalThis.DecompressionStream;
  try {
    globalThis.DecompressionStream = undefined;
    await rejected([graphic(rgb)], configuration, error => error.code === 'PLATFORM_UNAVAILABLE', 'native-inflater-unavailable');
  } finally { globalThis.DecompressionStream = native; }
  const abort = new AbortController(); let returned = false;
  const destination = new WritableStream({ write() { throw Error('No page should be emitted during PNG cancellation'); } });
  function* blocks() { try { yield graphic(rgb); } finally { returned = true; } }
  try {
    globalThis.DecompressionStream = class extends native { constructor(format) { super(format); queueMicrotask(() => abort.abort(Error('PNG native inflate cancelled'))); } };
    let caught; try { await writePdfDocumentTo(blocks(), destination, { ...configuration, signal: abort.signal }); } catch (error) { caught = error; }
    require(String(caught).includes('PNG native inflate cancelled') && returned && !destination.locked, 'PNG cancellation lost cleanup');
    checks.push('native-inflate-cancellation');
  } finally { globalThis.DecompressionStream = native; }
  const tagged = { language: 'en-US' }, columns = [{ header: 'Item', key: 'item' }, { header: 'Amount', key: 'amount' }];
  const report = await writePdfDocument([
    { kind: 'paragraph', text: 'Browser PNG report', headingLevel: 1, fontSize: 18 },
    { kind: 'paragraph', text: 'Independent RGB samples and straight alpha.' },
    { ...graphic(bytes(fixtures['rgba-pillow.png'])), width: 180, height: 120, alternativeText: 'Four color transparency quadrants' },
    pdfTable([{ item: 'Services', amount: 120 }, { item: 'Support', amount: 30 }], { columns, footer: { values: ['Total'], totals: { amount: 'sum' } } })
  ], { ...configuration, tagged, pageHeader: 'OfficeIMO', pageFooter: 'Browser proof', pageSize: 'A5' });
  const structured = await inspect(report);
  for (const token of ['/StructTreeRoot','/ParentTree','/StructParents','/MarkInfo << /Marked true >>','/S /H1','/S /Figure','/S /Table','/S /TH','/S /TD','/Scope /Column','/MCID','/Artifact'])
    require(structured.text.includes(token), 'Tagged browser PDF missing ' + token);
  checks.push('tagged-structure-and-figures');
  await save('tagged-png-report',report,['Browser PNG report','Independent RGB samples and straight alpha.','Item','Amount','Services','Support','Total','150']);
  await rejected([{ ...graphic(rgb) }], { ...configuration, tagged }, error => String(error).includes('alternativeText'), 'tagged-image-description');
  await rejected([pdfTable([], { columns, headerRows: [[{ value: 'Group', columnSpan: 2 },null]] })], { ...configuration, tagged },
    error => String(error).includes('Tagged PDF supports leaf'), 'tagged-spanned-heading');
  await rejected([{ kind: 'paragraph', text: 'Limited' }], { ...configuration, tagged, limits: { maxStructureElements: 1 } },
    error => error.code === 'RESOURCE_LIMIT', 'tagged-structure-limit');
  const drawing = { version: 1, width: 220, height: 120, elements: [
    { kind: 'group', clip: { x: 20, y: 20, width: 180, height: 80 }, elements: [
      { kind: 'path', fill: 'CCE5FF', commands: [{ op: 'move', x: 0, y: 0 }, { op: 'line', x: 220, y: 0 }, { op: 'line', x: 220, y: 120 }, { op: 'line', x: 0, y: 120 }, { op: 'close' }] },
      { kind: 'path', stroke: 'E02020', strokeWidth: 20, lineCap: 'round', lineJoin: 'round', commands: [{ op: 'move', x: 40, y: 60 }, { op: 'line', x: 180, y: 60 }] },
      { kind: 'path', stroke: '2050E0', strokeWidth: 20, lineCap: 'butt', lineJoin: 'bevel', commands: [{ op: 'move', x: 110, y: 0 }, { op: 'line', x: 110, y: 120 }] }
    ] },
    { kind: 'text', text: 'Portable drawing label', x: 25, y: 115, font: 'regular', fontSize: 10, color: '202020' }
  ] };
  const drawingBlock = { kind: 'drawing', drawing, alternativeText: 'Clipped red round-cap line overlapped by blue butt-cap line' };
  const drawingOptions = { ...configuration, tagged, fonts: { regular: new PdfFont(bytes(drawingFont)) }, pageSize: 'A5' };
  const drawingBlob = await writePdfDocument([{ kind: 'paragraph', text: 'Portable drawing', headingLevel: 1, fontSize: 18 }, drawingBlock], drawingOptions);
  const drawingPdf = await inspect(drawingBlob);
  for (const token of ['20 20 180 80 re W n','1 J','0 J','/Subtype /Type0','/FontFile2','/S /Figure'])
    require(drawingPdf.text.includes(token), 'Portable drawing lost clip/cap/font/figure contract: ' + token);
  await save('tagged-drawing', drawingBlob, ['Portable drawing','Portable drawing label']); checks.push('drawing-clip-caps-embedded-text-and-tagged-figure');
  await rejected([drawingBlock], { ...configuration, tagged }, error => String(error).includes('exact embedded font'), 'drawing-exact-font-required');
  await rejected([{ ...drawingBlock, drawing: { ...drawing, opacity: 0.5 } }], drawingOptions, error => String(error).includes('drawing properties'), 'drawing-unsupported-schema');
  await rejected([drawingBlock], { ...drawingOptions, limits: { maxDrawingElements: 1 } }, error => error.code === 'RESOURCE_LIMIT', 'drawing-element-limit');
  await verifyCapturedPdfInputs({ writePdfDocument, PdfFont, inspect, save, rejected, require, bytes, drawingFont, configuration, checks });
  return { checks, nativeInflater: typeof native === 'function', documents: 12, profile: classic ? 'classic' : 'esm' };
}

// Changing public getters must not replace values already admitted by resource and schema checks.
async function verifyCapturedPdfInputs({ writePdfDocument, PdfFont, inspect, save, rejected, require, bytes, drawingFont, configuration, checks }) {
  let dataReads = 0, formatReads = 0, heightReads = 0;
  const image = { width: 1, height: 1, get format() { return ++formatReads === 1 ? 'rgb' : 'jpeg'; },
    get data() { return ++dataReads === 1 ? Uint8Array.of(1,2,3) : Uint8Array.of(4,5,6,7,8,9); } };
  const imageBlob = await writePdfDocument([{ kind: 'image', image, get height() { return ++heightReads === 1 ? 17 : 34; } }], {
    ...configuration, limits: { maxImageBytes: 3, maxImagePixels: 1 }
  });
  const imagePdf = await inspect(imageBlob), resource = imagePdf.images[0];
  require(resource && resource.data.length === 3 && resource.data.every((value,index) => value === index+1), 'Image getter replaced admitted bytes');
  require(resource.body.includes('/DeviceRGB') && !resource.body.includes('/DCTDecode'), 'Image getter replaced admitted format');
  require(imagePdf.text.includes('q 17 0 0 17 ') && !imagePdf.text.includes('q 34 0 0 34 '), 'Image getter replaced admitted display height');
  await save('getter-image', imageBlob); checks.push('image-getter-admitted-bytes-format-display');
  await rejected([{ kind: 'image', image: { format: 'rgb', width: 1, height: 1, get data() { return new Uint8Array(6); } } }],
    { ...configuration, limits: { maxImageBytes: 3 } }, error => error.code === 'RESOURCE_LIMIT', 'image-getter-source-byte-limit');

  let commandReads = 0, textReads = 0, opReads = 0, fillReads = 0, strokeReads = 0, fontReads = 0, clipReads = 0, kindReads = 0;
  const commands = [{ get op() { return ++opReads === 1 ? 'move' : 'line'; }, x: 0, y: 0 }];
  const path = { get kind() { return ++kindReads === 1 ? 'path' : 'text'; },
    get fill() { return ++fillReads === 1 ? 'FF0000' : '0000FF'; },
    get stroke() { return ++strokeReads === 1 ? '008000' : '0000FF'; }, strokeWidth: 2,
    get commands() { return ++commandReads === 1 ? commands : [...commands,{ op: 'line', x: 9, y: 9 }]; } };
  const text = { kind: 'text', get text() { return ++textReads === 1 ? 'A' : 'BBBB'; }, x: 10, y: 25,
    get font() { return ++fontReads === 1 ? 'regular' : 'bold'; }, fontSize: 14, color: '000000' };
  const group = { kind: 'group', get clip() { return ++clipReads === 1 ? { x: 3, y: 4, width: 12, height: 13 } : { x: Infinity }; }, elements: [path,text] };
  const drawingBlob = await writePdfDocument([{ kind: 'drawing', drawing: { version: 1, width: 200, height: 100, elements: [group] } }], {
    ...configuration, fonts: { regular: new PdfFont(bytes(drawingFont)) }, limits: { maxVectorCommands: 1, maxCellCharacters: 1 }
  });
  const drawingPdf = await inspect(drawingBlob);
  for (const token of ['3 4 12 13 re W n','1 0 0 rg\n0 0.502 0 RG','0 0 m\nB'])
    require(drawingPdf.text.includes(token), 'Drawing getter replaced admitted path/paint/clip: ' + token);
  require(!drawingPdf.text.includes('9 9 l'), 'Drawing getter exceeded admitted command count');
  await save('getter-drawing', drawingBlob, ['A']); checks.push('drawing-getter-command-text-budget'); checks.push('drawing-getter-op-paint-font-clip');

  const vector = { kind: 'vector', width: 10, height: 10, stroke: '000000', commands: [{ op: 'move', x: 0, y: 0 },{ op: 'line', x: 10, y: 10 }] };
  let alternativeReads = 0, decorativeReads = 0;
  const alternativeBlob = await writePdfDocument([
    { ...vector, get alternativeText() { return ++alternativeReads === 1 ? 'A' : 'BBBB'; } },
    { ...vector, get decorative() { return ++decorativeReads === 1 ? true : false; } }
  ], { ...configuration, tagged: { language: 'en-US' }, limits: { maxCellCharacters: 1, maxCells: 1, maxTextCharacters: 1 } });
  const alternativePdf = await inspect(alternativeBlob);
  require((alternativePdf.text.match(/\/S \/Figure /g) ?? []).length === 1 && alternativePdf.text.includes('/Alt <feff0041>'), 'Graphic getters replaced admitted alternative text');
  require((alternativePdf.text.match(/\/Artifact << \/Type \/Layout >> BDC/g) ?? []).length === 1, 'Graphic getter replaced decorative intent');
  await save('getter-alternative', alternativeBlob); checks.push('graphic-getter-alternative-decorative-budget');

  let paragraphReads = 0, levelReads = 0, presentationReads = 0, boldReads = 0, italicReads = 0, colorReads = 0, alignmentReads = 0;
  const paragraphStyle = {
    get bold() { return ++boldReads <= 2 ? false : true; }, get italic() { return ++italicReads <= 2 ? false : true; },
    get color() { return ++colorReads <= 2 ? '008000' : '0000FF'; }, get alignment() { return ++alignmentReads <= 2 ? 'left' : 'right'; }
  };
  const paragraphBlob = await writePdfDocument([{ kind: 'paragraph', get text() { return ++paragraphReads === 1 ? 'A' : 'BBBB'; },
    get headingLevel() { return ++levelReads === 1 ? 2 : 9; },
    get presentation() { return ++presentationReads === 1 ? paragraphStyle : { background: 'FF0000' }; }
  }], { ...configuration, tagged: { language: 'en-US' }, limits: { maxCellCharacters: 1, maxTextCharacters: 1 } });
  const paragraphPdf = await inspect(paragraphBlob);
  require(paragraphPdf.text.includes('/S /H2 ') && !paragraphPdf.text.includes('/S /H9 '), 'Paragraph getter replaced admitted heading');
  await save('getter-paragraph', paragraphBlob, ['A']); checks.push('paragraph-getter-text-heading-budget');
  require([boldReads,italicReads,colorReads,alignmentReads].every(count => count === 1), 'Paragraph presentation getters were reread after capture');
  require(paragraphPdf.text.includes('0 0.502 0 rg') && !paragraphPdf.text.includes('0 0 1 rg'), 'Paragraph presentation getter replaced admitted color');
  require(!paragraphPdf.text.includes('Helvetica-Bold') && !paragraphPdf.text.includes('Helvetica-Oblique') && paragraphPdf.text.includes('1 0 0 1 36 '), 'Paragraph presentation getter replaced admitted font/alignment');
  checks.push('paragraph-nested-presentation-getter-capture');
  for (const [key,invalid,valid] of [['bold','yes',false],['italic','yes',false],['color','invalid','008000'],['alignment','justify','left']]) {
    let reads = 0;
    await rejected([{ kind: 'paragraph', text: 'A', presentation: { get [key]() { return ++reads === 1 ? invalid : valid; } } }], configuration,
      error => error instanceof TypeError, 'paragraph-presentation-invalid-' + key);
    require(reads === 1, 'Invalid paragraph presentation was reread: ' + key);
  }
}
