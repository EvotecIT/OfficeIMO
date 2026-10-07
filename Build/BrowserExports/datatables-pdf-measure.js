// Test-only normalization of the requested report geometry; generation remains the installed native button action.
let comparisonPdf;
globalThis.preparePdfMeasurement = function (encoded) {
  const decode = text => Uint8Array.from(atob(text), c => c.charCodeAt(0));
  const programs = { regular: decode(encoded.regular), bold: decode(encoded.bold) };
  const spacing = bytes => {
    const view = new DataView(bytes.buffer); let units, ascent, descent;
    for (let i = 0; i < view.getUint16(4); i++) {
      const p = 12 + i * 16, at = view.getUint32(p + 8), tag = String.fromCharCode(...bytes.subarray(p, p + 4));
      if (tag === 'head') units = view.getUint16(at + 18);
      if (tag === 'hhea') { ascent = view.getInt16(at + 4); descent = view.getInt16(at + 6); }
    }
    const ratio = (ascent - descent) / units;
    return Math.max(1.3, ratio + .1) / ratio;
  };
  const files = { 'report-regular.ttf': encoded.regular, 'report-bold.ttf': encoded.bold };
  const faces = { Report: { normal: 'report-regular.ttf', bold: 'report-bold.ttf', italics: 'report-regular.ttf', bolditalics: 'report-bold.ttf' } };
  if (typeof pdfMake.addVirtualFileSystem === 'function') pdfMake.addVirtualFileSystem(files); else pdfMake.vfs = files;
  if (typeof pdfMake.addFonts === 'function') pdfMake.addFonts(faces); else pdfMake.fonts = faces;
  DataTable.Buttons.pdfMake(pdfMake);
  comparisonPdf = { fonts: { regular: new OfficeIMO.PdfFont(programs.regular), bold: new OfficeIMO.PdfFont(programs.bold) },
    regularSpacing: spacing(programs.regular), boldSpacing: spacing(programs.bold) };
};

globalThis.customizePdfMeasurement = function (doc, spec) {
  const table = doc.content.find(item => item.table), width = (1190.551 - 72) / spec.columns;
  doc.pageMargins = [36, 36, 36, 36];
  doc.defaultStyle = { font: 'Report', fontSize: 6, lineHeight: comparisonPdf.regularSpacing, color: '#111827' };
  doc.styles.table = { margin: [0, 0, 0, 0] };
  doc.styles.tableHeader = { bold: true, fontSize: 6, lineHeight: comparisonPdf.boldSpacing, alignment: 'left', color: '#111827' };
  // pdfmake measures content plus padding and shared strokes; OfficeIMO measures the outer cell rectangle.
  table.table.widths = Array.from({ length: spec.columns }, () => width - 8 - .4 / spec.columns);
  table.layout = {
    hLineWidth: () => .4, vLineWidth: () => .4, hLineColor: () => '#d1d9e3', vLineColor: () => '#d1d9e3',
    paddingLeft: () => 3.8, paddingRight: () => 3.8, paddingTop: () => 3.8, paddingBottom: () => 3.8,
    fillColor: index => index === 0 ? '#e7edf5' : spec.styled && index % 2 === 0 ? '#f3f4f6' : null
  };
};

globalThis.pdfMeasurementOptions = function (spec) {
  return { fonts: comparisonPdf.fonts, pageSize: 'A3', orientation: 'landscape', fontSize: 6,
    columnWidths: Array.from({ length: spec.columns }, () => (1190.551 - 72) / spec.columns), pageNumbers: false,
    ...(spec.styled ? { alternateRowColor: 'F3F4F6' } : {}) };
};
