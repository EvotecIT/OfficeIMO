using DocumentFormat.OpenXml.Spreadsheet;
using System.Threading;
using System.Xml;

namespace OfficeIMO.Excel {
    /// <summary>
    /// Range enumeration for <see cref="ExcelSheetReader"/>.
    /// </summary>
    public sealed partial class ExcelSheetReader {
        /// <summary>
        /// Enumerates non-empty cells within the given A1 range as typed values.
        /// </summary>
        public IEnumerable<CellValueInfo> EnumerateRange(string a1Range) {
            var (r1, c1, r2, c2) = A1.ParseRange(a1Range);

            return CanUseEnumerateRangeXmlReader()
                ? EnumerateRangeXmlFast(r1, c1, r2, c2, CancellationToken.None)
                : EnumerateRangeDom(r1, c1, r2, c2, CancellationToken.None);
        }

        private IEnumerable<CellValueInfo> EnumerateRangeDom(int r1, int c1, int r2, int c2, CancellationToken ct) {
            bool canCancel = ct.CanBeCanceled;
            foreach (var row in EnumerateWorksheetRows(ct)) {
                if (canCancel) {
                    ct.ThrowIfCancellationRequested();
                }

                var rIndex = checked((int)row.RowIndex!.Value);
                if (rIndex < r1) continue;
                if (rIndex > r2) continue;

                foreach (var cell in row.Elements<Cell>()) {
                    if (canCancel) {
                        ct.ThrowIfCancellationRequested();
                    }

                    int cIndex = A1.ParseColumnIndexFromCellReferenceFast(cell.CellReference?.Value);
                    if (cIndex < c1 || cIndex > c2) continue;
                    if (TryConvertCell(cell, out var value))
                        yield return new CellValueInfo(rIndex, cIndex, value);
                }
            }
        }

        private IEnumerable<CellValueInfo> EnumerateRangeXmlFast(int r1, int c1, int r2, int c2, CancellationToken ct) {
            using var stream = _wsPart.GetStream(System.IO.FileMode.Open, System.IO.FileAccess.Read);
            RewindWorksheetStream(stream);
            using var reader = OpenWorksheetXmlReader(stream);
            bool canCancel = ct.CanBeCanceled;
            bool fillBlanks = _opt.FillBlanksInRanges;
            bool hasCustomConverter = _opt.CellValueConverter != null;
            int nextRowIndex = 1;

            while (reader.Read()) {
                if (canCancel) {
                    ct.ThrowIfCancellationRequested();
                }

                if (reader.NodeType != XmlNodeType.Element || reader.LocalName != "row") {
                    continue;
                }

                int rowIndex = ParsePositiveIntAttribute(reader.GetAttribute("r"));
                if (rowIndex <= 0) {
                    rowIndex = nextRowIndex;
                }

                nextRowIndex = rowIndex + 1;
                if (rowIndex < r1 || rowIndex > r2) {
                    SkipXmlElement(reader, "row");
                    continue;
                }

                if (reader.IsEmptyElement) {
                    continue;
                }

                int depth = reader.Depth;
                int nextColumnIndex = 1;
                while (reader.Read()) {
                    if (canCancel) {
                        ct.ThrowIfCancellationRequested();
                    }

                    if (reader.NodeType == XmlNodeType.EndElement && reader.Depth == depth && reader.LocalName == "row") {
                        break;
                    }

                    if (reader.NodeType != XmlNodeType.Element || reader.LocalName != "c") {
                        continue;
                    }

                    int columnIndex = GetXmlCellColumnIndex(reader, ref nextColumnIndex);
                    if (columnIndex <= 0) {
                        SkipXmlElement(reader, "c");
                        continue;
                    }

                    if (columnIndex < c1 || columnIndex > c2) {
                        SkipXmlElement(reader, "c");
                        continue;
                    }

                    if (hasCustomConverter) {
                        if (TryReadXmlCellValueForEnumeration(reader, rowIndex, columnIndex, out object? customValue)) {
                            yield return new CellValueInfo(rowIndex, columnIndex, customValue);
                        }
                    } else if (fillBlanks) {
                        yield return new CellValueInfo(rowIndex, columnIndex, ReadXmlCellValue(reader));
                    } else if (!reader.IsEmptyElement) {
                        object? cellValue = ReadXmlCellValue(reader);
                        if (cellValue != null) {
                            yield return new CellValueInfo(rowIndex, columnIndex, cellValue);
                        }
                    }
                }
            }
        }

        private bool TryReadXmlCellValueForEnumeration(XmlReader cellReader, int rowIndex, int columnIndex, out object? value) {
            XmlCellKind cellKind = ParseXmlCellKind(cellReader.GetAttribute("t"));
            bool readStyleIndex = true;

            CellRaw raw = ReadXmlCellRaw(cellReader, rowIndex, columnIndex, cellKind, readStyleIndex);
            if (raw.RawText == null
                && raw.InlineText == null
                && raw.FormulaText == null
                && !_opt.FillBlanksInRanges) {
                value = null;
                return false;
            }

            value = ConvertRaw(raw).TypedValue;
            return true;
        }

        private bool CanUseEnumerateRangeXmlReader() {
            return (_opt.CellValueConverter != null || _opt.Culture == System.Globalization.CultureInfo.InvariantCulture)
                && CanStreamWorksheetPart();
        }
    }
}
