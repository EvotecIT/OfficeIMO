function styleRegistry() {
  const formats = new Map(), fills = new Map(), indexes = new Map(), styles = [];
  function add(column = {}, header = false, fill, date = false) {
    const format = column.format ?? (column.type === "date" || date ? "yyyy-mm-dd hh:mm:ss" : "");
    if (typeof format !== "string") throw new TypeError("Column format must be a string.");
    const key = JSON.stringify([format, !!column.wrapText, column.alignment ?? "", header, fill ?? ""]);
    if (indexes.has(key)) return indexes.get(key);
    if (styles.length >= 64000) throw new RangeError("Workbook exceeds Excel's cell style limit.");
    let numFmtId = 0, fillId = 0;
    if (format) {
      if (!formats.has(format)) formats.set(format, formats.size + 164);
      numFmtId = formats.get(format);
    }
    if (fill) {
      if (!fills.has(fill)) fills.set(fill, fills.size + 2);
      fillId = fills.get(fill);
    }
    const alignment = column.wrapText || column.alignment
      ? '<alignment' + (column.wrapText ? ' wrapText="1"' : "") +
        (column.alignment ? ' horizontal="' + xml(column.alignment) + '"' : "") + '/>' : "";
    styles.push('<xf numFmtId="' + numFmtId + '" fontId="' + (header ? 1 : 0) + '" fillId="' + fillId +
      '" borderId="0" xfId="0"' + (numFmtId ? ' applyNumberFormat="1"' : "") +
      (header ? ' applyFont="1"' : "") + (fillId ? ' applyFill="1"' : "") +
      (alignment ? ' applyAlignment="1"' : "") + '>' + alignment + '</xf>');
    const index = styles.length - 1;
    indexes.set(key, index);
    return index;
  }
  add();
  return {
    add,
    xml() {
      return xmlDeclaration + '<styleSheet xmlns="' + spreadsheetNs + '">' +
        '<numFmts count="' + formats.size + '">' + [...formats].map(([format, id]) =>
          '<numFmt numFmtId="' + id + '" formatCode="' + xml(format) + '"/>').join("") + '</numFmts>' +
        '<fonts count="2"><font><sz val="11"/><name val="Calibri"/></font>' +
        '<font><b/><sz val="11"/><name val="Calibri"/></font></fonts>' +
        '<fills count="' + (fills.size + 2) + '"><fill><patternFill patternType="none"/></fill>' +
        '<fill><patternFill patternType="gray125"/></fill>' + [...fills.keys()].map(fill =>
          '<fill><patternFill patternType="solid"><fgColor rgb="' + fill +
          '"/><bgColor indexed="64"/></patternFill></fill>').join("") + '</fills>' +
        '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>' +
        '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>' +
        '<cellXfs count="' + styles.length + '">' + styles.join("") + '</cellXfs>' +
        '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>';
    }
  };
}
