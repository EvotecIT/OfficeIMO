using System.Globalization;
using System.Text;
using OfficeIMO.Core.Internal.OpenXml;

namespace OfficeIMO.Excel {
    /// <summary>A cell value the owned save writes, with its shared-string index or text already resolved.</summary>
    internal sealed class ExcelDeferredCellOutput {
        internal ExcelDeferredCellOutput(string type, string value, bool date) {
            Type = type;
            Value = value;
            Date = date;
        }

        /// <summary>The cell type attribute: <c>s</c>, <c>str</c>, <c>n</c> or <c>b</c>.</summary>
        internal string Type { get; }
        /// <summary>The text of the value element, not yet escaped for XML.</summary>
        internal string Value { get; }
        /// <summary>True when the cell takes the automatic date format derived from its base style.</summary>
        internal bool Date { get; }
    }

    /// <summary>
    /// Writes a worksheet with edited cell values by streaming its source: rows and cells the edits do not touch are
    /// copied as they were read, edited cells are written as the model's save writes plain value cells, and the
    /// dimension covers every cell, as the model's save computes it. Markup the merge does not understand where it
    /// would have to change something returns null, and the save goes through the SDK model instead.
    /// </summary>
    internal static class ExcelDeferredWorksheetWriter {
        private static readonly UTF8Encoding Utf8 = new UTF8Encoding(encoderShouldEmitUTF8Identifier: false);

        /// <summary>
        /// Returns the merged worksheet, or null when it must be written through the model. <paramref name="rows"/>
        /// maps row numbers to the edited cells of the row by column. <paramref name="dateStyle"/> maps a cell's base
        /// style to the style its date value takes.
        /// </summary>
        internal static byte[]? TryMerge(byte[] xml, int length, SortedDictionary<int, SortedDictionary<int, ExcelDeferredCellOutput>> rows,
            Func<uint, uint?> dateStyle, out int resultLength, out int removedSharedReferences) {
            resultLength = 0;
            removedSharedReferences = 0;
            var reader = new OfficeXmlPullReader(xml, length, processMarkupCompatibility: false);
            var merge = new Merge(xml, rows, dateStyle);
            try {
                if (!merge.Run(reader, out resultLength)) return null;
                removedSharedReferences = merge.RemovedSharedReferences;
                return merge.Output;
            } catch (Exception exception) when (exception is System.Xml.XmlException or DateStyleUnavailableException) {
                return null;
            }
        }

        private sealed class Merge {
            private readonly byte[] _source;
            private readonly SortedDictionary<int, SortedDictionary<int, ExcelDeferredCellOutput>> _rows;
            private readonly Func<uint, uint?> _dateStyle;
            private readonly List<int> _pendingRows;
            private int _nextPendingRow;
            private byte[] _output;
            private int _outputLength;
            private int _copied;
            // Column ranges and styles of the first <cols>, which new cells inherit when their row has no format.
            private readonly List<(int Min, int Max, uint Style)> _columns = new();
            private bool _columnsSeen;
            private int _minRow = int.MaxValue, _maxRow, _minColumn = int.MaxValue, _maxColumn;

            internal Merge(byte[] source, SortedDictionary<int, SortedDictionary<int, ExcelDeferredCellOutput>> rows, Func<uint, uint?> dateStyle) {
                _source = source;
                _rows = rows;
                _dateStyle = dateStyle;
                _pendingRows = new List<int>(rows.Keys);
                _output = new byte[Math.Max(256, source.Length + 4096)];
            }

            internal byte[] Output => _output;

            /// <summary>Edited cells that held a shared string before the edit.</summary>
            internal int RemovedSharedReferences { get; private set; }

            internal bool Run(OfficeXmlPullReader reader, out int resultLength) {
                resultLength = 0;
                string mainPrefix = string.Empty;
                // Where the dimension's reference value is, in the output, or where a new dimension element goes.
                int dimensionValueAt = -1, dimensionValueLength = 0, dimensionInsertAt = -1;
                bool rootSeen = false, dimensionSeen = false, sheetDataSeen = false, inSheetData = false, inColumns = false;
                int previousRow = 0;

                while (reader.Read()) {
                    int depth = reader.Depth;
                    if (reader.NodeKind == OfficeXmlNodeKind.StartElement) {
                        if (depth == 1) {
                            if (reader.Transcoded || !reader.IsElement(OfficeXmlNamespaces.SpreadsheetML, "worksheet")
                                || reader.ElementUriId != OfficeXmlNamespaces.TransitionalUriId(OfficeXmlNamespaces.SpreadsheetML)) {
                                return false;
                            }
                            rootSeen = true;
                            mainPrefix = reader.ElementPrefix;
                            // A worksheet without sheet properties takes its new dimension first.
                            dimensionInsertAt = reader.MarkupEnd;
                            continue;
                        }
                        if (depth == 2) {
                            if (reader.NamespaceId != OfficeXmlNamespaces.SpreadsheetML) continue;
                            if (reader.LocalNameEquals("sheetPr")) {
                                if (reader.IsEmptyElement) dimensionInsertAt = reader.MarkupEnd;
                                else {
                                    reader.SkipElement();
                                    dimensionInsertAt = reader.MarkupEnd;
                                }
                                continue;
                            }
                            if (reader.LocalNameEquals("dimension")) {
                                // The model removes extra dimensions and adds a missing reference; those stay with it.
                                if (dimensionSeen || !TryFindAttribute(reader, "ref", out int start, out int valueLength)) return false;
                                dimensionSeen = true;
                                CopyTo(start);
                                dimensionValueAt = _outputLength;
                                dimensionValueLength = 0;
                                _copied = start + valueLength;
                                continue;
                            }
                            if (reader.LocalNameEquals("cols") && !_columnsSeen && !sheetDataSeen) {
                                _columnsSeen = true;
                                inColumns = !reader.IsEmptyElement;
                                continue;
                            }
                            if (reader.LocalNameEquals("sheetData")) {
                                if (sheetDataSeen) return false;
                                sheetDataSeen = true;
                                string prefix = reader.ElementPrefix;
                                if (reader.IsEmptyElement) {
                                    if (_pendingRows.Count > 0) {
                                        // <sheetData/> becomes an element holding the new rows.
                                        CopyTo(reader.MarkupStart);
                                        AppendAscii("<" + Qualified(prefix, "sheetData") + ">");
                                        if (!WritePendingRowsBefore(int.MaxValue, prefix)) return false;
                                        AppendAscii("</" + Qualified(prefix, "sheetData") + ">");
                                        _copied = reader.MarkupEnd;
                                    }
                                    continue;
                                }
                                inSheetData = true;
                                if (!ReadSheetData(reader, prefix, ref previousRow)) return false;
                                inSheetData = false;
                                continue;
                            }
                            continue;
                        }
                        if (depth == 3 && inColumns && reader.IsElement(OfficeXmlNamespaces.SpreadsheetML, "col")) {
                            ReadColumn(reader);
                        }
                        continue;
                    }
                    if (reader.NodeKind == OfficeXmlNodeKind.EndElement && depth == 2) inColumns = false;
                }

                if (!rootSeen || !sheetDataSeen || inSheetData || _nextPendingRow < _pendingRows.Count) return false;
                CopyTo(_source.Length > 0 ? reader.DataLength : 0);

                string reference = _maxRow == 0 || _maxColumn == 0
                    ? "A1"
                    : Range(_minRow, _minColumn, _maxRow, _maxColumn);
                byte[] referenceBytes = Encoding.ASCII.GetBytes(reference);
                if (dimensionSeen) {
                    Splice(dimensionValueAt, dimensionValueLength, referenceBytes);
                } else {
                    if (dimensionInsertAt < 0) return false;
                    // The insertion point is a source position; everything before it was copied unchanged.
                    byte[] element = Utf8.GetBytes("<" + Qualified(mainPrefix, "dimension") + " ref=\"" + reference + "\" />");
                    Splice(OutputPositionOf(dimensionInsertAt), 0, element);
                }
                resultLength = _outputLength;
                return true;
            }

            // Source positions before the first change map to the same output positions.
            private int OutputPositionOf(int sourcePosition) => sourcePosition <= _firstChangeAt ? sourcePosition : -1;

            private int _firstChangeAt = int.MaxValue;

            private void ReadColumn(OfficeXmlPullReader reader) {
                if (!reader.TryGetAttribute(OfficeXmlNamespaces.None, "min", out string minText)
                    || !reader.TryGetAttribute(OfficeXmlNamespaces.None, "max", out string maxText)
                    || !uint.TryParse(minText, NumberStyles.None, CultureInfo.InvariantCulture, out uint min)
                    || !uint.TryParse(maxText, NumberStyles.None, CultureInfo.InvariantCulture, out uint max)) {
                    _columns.Add((int.MaxValue, 0, 0U));
                    return;
                }
                uint style = 0U;
                if (reader.TryGetAttribute(OfficeXmlNamespaces.None, "style", out string styleText)
                    && !uint.TryParse(styleText, NumberStyles.None, CultureInfo.InvariantCulture, out style)) {
                    style = 0U;
                }
                _columns.Add(((int)Math.Min(min, int.MaxValue), (int)Math.Min(max, int.MaxValue), style));
            }

            private bool ReadSheetData(OfficeXmlPullReader reader, string sheetDataPrefix, ref int previousRow) {
                while (reader.Read()) {
                    if (reader.NodeKind == OfficeXmlNodeKind.EndElement && reader.Depth == 2) {
                        // New rows after the last one go before </sheetData>.
                        CopyTo(reader.MarkupStart);
                        return WritePendingRowsBefore(int.MaxValue, sheetDataPrefix);
                    }
                    if (reader.NodeKind != OfficeXmlNodeKind.StartElement || reader.Depth != 3) continue;
                    if (!reader.IsElement(OfficeXmlNamespaces.SpreadsheetML, "row")) {
                        // Anything else between rows leaves the place of a new row unclear.
                        if (_nextPendingRow < _pendingRows.Count) return false;
                        reader.SkipElement();
                        continue;
                    }
                    int rowNumber;
                    if (reader.TryGetAttribute(OfficeXmlNamespaces.None, "r", out string rowText)) {
                        if (!int.TryParse(rowText, NumberStyles.None, CultureInfo.InvariantCulture, out rowNumber)
                            || rowNumber < 1 || rowNumber > ExcelA1.MaxRows) return false;
                    } else {
                        rowNumber = previousRow + 1;
                    }
                    // Rows out of order leave the place of a new row unclear.
                    if (rowNumber <= previousRow) return false;
                    previousRow = rowNumber;

                    int rowStart = reader.MarkupStart;
                    string rowPrefix = reader.ElementPrefix;
                    if (_nextPendingRow < _pendingRows.Count && _pendingRows[_nextPendingRow] < rowNumber) {
                        CopyTo(rowStart);
                        if (!WritePendingRowsBefore(rowNumber, rowPrefix)) return false;
                    }

                    SortedDictionary<int, ExcelDeferredCellOutput>? edits = null;
                    if (_nextPendingRow < _pendingRows.Count && _pendingRows[_nextPendingRow] == rowNumber) {
                        edits = _rows[rowNumber];
                        _nextPendingRow++;
                    }
                    if (!ReadRow(reader, rowNumber, rowPrefix, edits)) return false;
                }
                return false;
            }

            private bool ReadRow(OfficeXmlPullReader reader, int rowNumber, string rowPrefix, SortedDictionary<int, ExcelDeferredCellOutput>? edits) {
                uint? rowStyle = null;
                if (edits != null && reader.TryGetAttribute(OfficeXmlNamespaces.None, "customFormat", out string customFormat)
                    && customFormat is "1" or "true"
                    && reader.TryGetAttribute(OfficeXmlNamespaces.None, "s", out string rowStyleText)) {
                    if (!uint.TryParse(rowStyleText, NumberStyles.None, CultureInfo.InvariantCulture, out uint parsed)) return false;
                    rowStyle = parsed;
                }
                using IEnumerator<KeyValuePair<int, ExcelDeferredCellOutput>>? pending = edits?.GetEnumerator();
                bool hasPending = pending?.MoveNext() == true;

                if (reader.IsEmptyElement) {
                    int tagStart = reader.MarkupStart, tagEnd = reader.MarkupEnd;
                    reader.Read();
                    if (!hasPending) return true;
                    // <row .../> becomes an element holding its new cells.
                    int slash = tagEnd - 2;
                    while (slash > tagStart && _source[slash] != (byte)'/') slash--;
                    if (_source[slash] != (byte)'/') return false;
                    CopyTo(slash);
                    AppendAscii(">");
                    _copied = tagEnd;
                    while (hasPending) {
                        WriteCell(rowNumber, pending!.Current.Key, pending.Current.Value, rowPrefix, existingStyle: null,
                            inheritedStyle: ResolveInheritedStyle(rowStyle, pending.Current.Key));
                        hasPending = pending.MoveNext();
                    }
                    AppendAscii("</" + Qualified(rowPrefix, "row") + ">");
                    return true;
                }

                int ordinal = 0, nextColumn = 1, previousColumn = 0;
                while (reader.Read()) {
                    if (reader.NodeKind == OfficeXmlNodeKind.EndElement && reader.Depth == 3) {
                        if (hasPending) {
                            CopyTo(reader.MarkupStart);
                            while (hasPending) {
                                WriteCell(rowNumber, pending!.Current.Key, pending.Current.Value, rowPrefix, existingStyle: null,
                                    inheritedStyle: ResolveInheritedStyle(rowStyle, pending.Current.Key));
                                hasPending = pending.MoveNext();
                            }
                        }
                        return true;
                    }
                    if (reader.NodeKind != OfficeXmlNodeKind.StartElement || reader.Depth != 4) continue;
                    if (!reader.IsElement(OfficeXmlNamespaces.SpreadsheetML, "c")) {
                        if (edits != null) return false;
                        reader.SkipElement();
                        continue;
                    }
                    ordinal++;
                    string? reference = reader.TryGetAttribute(OfficeXmlNamespaces.None, "r", out string referenceText) ? referenceText : null;
                    int resolvedRow = rowNumber, resolvedColumn = 0;
                    if (ExcelA1.TryParseCellReferenceFast(reference, out int parsedRow, out int parsedColumn)) {
                        if (parsedRow > 0) resolvedRow = parsedRow;
                        if (parsedColumn > 0) resolvedColumn = parsedColumn;
                    }
                    if (resolvedColumn <= 0) resolvedColumn = string.IsNullOrEmpty(reference) ? nextColumn : ordinal;
                    nextColumn = resolvedColumn + 1;
                    if (resolvedRow > 0 && resolvedRow <= ExcelA1.MaxRows && resolvedColumn > 0 && resolvedColumn <= ExcelA1.MaxColumns) {
                        Include(resolvedRow, resolvedColumn);
                    }

                    if (edits == null) {
                        reader.SkipElement();
                        continue;
                    }
                    // An edited row is merged by column, which needs canonical references in order.
                    if (resolvedRow != rowNumber || resolvedColumn <= previousColumn
                        || reference == null || !string.Equals(reference, ExcelA1.CellReference(rowNumber, resolvedColumn), StringComparison.Ordinal)) {
                        return false;
                    }
                    previousColumn = resolvedColumn;

                    int cellStart = reader.MarkupStart;
                    if (hasPending && pending!.Current.Key < resolvedColumn) {
                        CopyTo(cellStart);
                        while (hasPending && pending!.Current.Key < resolvedColumn) {
                            WriteCell(rowNumber, pending.Current.Key, pending.Current.Value, rowPrefix, existingStyle: null,
                                inheritedStyle: ResolveInheritedStyle(rowStyle, pending.Current.Key));
                            hasPending = pending.MoveNext();
                        }
                    }
                    if (!hasPending || pending!.Current.Key != resolvedColumn) {
                        reader.SkipElement();
                        continue;
                    }

                    // The edited cell keeps its style; anything beyond a plain value goes through the model.
                    if (!TryReadPlainCellStyle(reader, out string? existingStyle)) return false;
                    if (reader.TryGetAttribute(OfficeXmlNamespaces.None, "t", out string existingType) && existingType == "s") RemovedSharedReferences++;
                    string cellPrefix = reader.ElementPrefix;
                    if (!reader.IsEmptyElement) {
                        while (reader.Read()) {
                            if (reader.NodeKind == OfficeXmlNodeKind.EndElement && reader.Depth == 4) break;
                            if (reader.NodeKind == OfficeXmlNodeKind.StartElement && reader.Depth == 5) {
                                if (!reader.IsElement(OfficeXmlNamespaces.SpreadsheetML, "v")) return false;
                                reader.SkipElement();
                            }
                        }
                    } else {
                        reader.Read();
                    }
                    int cellEnd = reader.MarkupEnd;
                    CopyTo(cellStart);
                    WriteCell(rowNumber, resolvedColumn, pending.Current.Value, cellPrefix, existingStyle, inheritedStyle: 0U);
                    _copied = cellEnd;
                    hasPending = pending.MoveNext();
                }
                return false;
            }

            /// <summary>Reads the style of a cell holding only a reference, a style, a type and a value.</summary>
            private static bool TryReadPlainCellStyle(OfficeXmlPullReader reader, out string? style) {
                style = null;
                for (int index = 0; index < reader.AttributeCount; index++) {
                    int namespaceId = reader.GetAttributeNamespace(index, out _);
                    string name = reader.GetAttributeLocalName(index);
                    if (namespaceId == OfficeXmlNamespaces.Xmlns) return false;
                    if (namespaceId != OfficeXmlNamespaces.None) return false;
                    switch (name) {
                        case "r":
                        case "t":
                            break;
                        case "s":
                            style = reader.GetAttributeValue(index);
                            if (!uint.TryParse(style, NumberStyles.None, CultureInfo.InvariantCulture, out _)) return false;
                            break;
                        default:
                            return false;
                    }
                }
                return true;
            }

            private uint ResolveInheritedStyle(uint? rowStyle, int column) {
                if (rowStyle.HasValue) return rowStyle.Value;
                foreach ((int min, int max, uint style) in _columns) {
                    if (min <= column && column <= max) return style;
                }
                return 0U;
            }

            private bool WritePendingRowsBefore(int limit, string prefix) {
                while (_nextPendingRow < _pendingRows.Count && _pendingRows[_nextPendingRow] < limit) {
                    int rowNumber = _pendingRows[_nextPendingRow++];
                    string row = Qualified(prefix, "row");
                    AppendAscii("<" + row + " r=\"" + rowNumber.ToString(CultureInfo.InvariantCulture) + "\">");
                    foreach (KeyValuePair<int, ExcelDeferredCellOutput> cell in _rows[rowNumber]) {
                        WriteCell(rowNumber, cell.Key, cell.Value, prefix, existingStyle: null, inheritedStyle: ResolveInheritedStyle(null, cell.Key));
                    }
                    AppendAscii("</" + row + ">");
                }
                return true;
            }

            private void WriteCell(int row, int column, ExcelDeferredCellOutput value, string prefix, string? existingStyle, uint inheritedStyle) {
                Include(row, column);
                string cell = Qualified(prefix, "c");
                var text = new StringBuilder(64);
                text.Append('<').Append(cell).Append(" r=\"").Append(ExcelA1.CellReference(row, column)).Append('"');
                if (value.Date) {
                    uint baseStyle = existingStyle != null ? uint.Parse(existingStyle, CultureInfo.InvariantCulture) : inheritedStyle;
                    uint style = _dateStyle(baseStyle) ?? throw new DateStyleUnavailableException();
                    text.Append(" s=\"").Append(style.ToString(CultureInfo.InvariantCulture)).Append('"');
                } else if (existingStyle != null) {
                    text.Append(" s=\"").Append(existingStyle).Append('"');
                } else if (inheritedStyle != 0U) {
                    text.Append(" s=\"").Append(inheritedStyle.ToString(CultureInfo.InvariantCulture)).Append('"');
                }
                text.Append(" t=\"").Append(value.Type).Append("\"><").Append(Qualified(prefix, "v")).Append('>');
                AppendEscaped(text, value.Value);
                text.Append("</").Append(Qualified(prefix, "v")).Append("></").Append(cell).Append('>');
                AppendText(text.ToString());
            }

            private void Include(int row, int column) {
                if (row < _minRow) _minRow = row;
                if (row > _maxRow) _maxRow = row;
                if (column < _minColumn) _minColumn = column;
                if (column > _maxColumn) _maxColumn = column;
            }

            private static string Range(int minRow, int minColumn, int maxRow, int maxColumn) {
                string start = ExcelA1.CellReference(minRow, minColumn);
                string end = ExcelA1.CellReference(maxRow, maxColumn);
                return start == end ? start : start + ":" + end;
            }

            private static bool TryFindAttribute(OfficeXmlPullReader reader, string localName, out int start, out int length) {
                for (int index = 0; index < reader.AttributeCount; index++) {
                    if (reader.GetAttributeNamespace(index, out _) == OfficeXmlNamespaces.None
                        && string.Equals(reader.GetAttributeLocalName(index), localName, StringComparison.Ordinal)) {
                        reader.GetAttributeRawValue(index, out start, out length);
                        return true;
                    }
                }
                start = length = 0;
                return false;
            }

            private static string Qualified(string prefix, string localName) => prefix.Length == 0 ? localName : prefix + ":" + localName;

            private static void AppendEscaped(StringBuilder text, string value) {
                foreach (char character in value) {
                    switch (character) {
                        case '&': text.Append("&amp;"); break;
                        case '<': text.Append("&lt;"); break;
                        case '>': text.Append("&gt;"); break;
                        default: text.Append(character); break;
                    }
                }
            }

            private void CopyTo(int sourcePosition) {
                if (sourcePosition <= _copied) return;
                Ensure(sourcePosition - _copied);
                Buffer.BlockCopy(_source, _copied, _output, _outputLength, sourcePosition - _copied);
                _outputLength += sourcePosition - _copied;
                _copied = sourcePosition;
            }

            private void AppendAscii(string text) => AppendText(text);

            private void AppendText(string text) {
                if (_firstChangeAt == int.MaxValue) _firstChangeAt = _copied;
                int count = Utf8.GetByteCount(text);
                Ensure(count);
                _outputLength += Utf8.GetBytes(text, 0, text.Length, _output, _outputLength);
            }

            private void Splice(int at, int removeLength, byte[] insert) {
                if (at < 0) throw new InvalidOperationException("The dimension position is unknown.");
                Ensure(insert.Length - removeLength);
                Buffer.BlockCopy(_output, at + removeLength, _output, at + insert.Length, _outputLength - at - removeLength);
                Buffer.BlockCopy(insert, 0, _output, at, insert.Length);
                _outputLength += insert.Length - removeLength;
            }

            private void Ensure(int additional) {
                if (_outputLength + additional <= _output.Length) return;
                Array.Resize(ref _output, Math.Max(_output.Length * 2, _outputLength + additional + 1024));
            }
        }

        /// <summary>The styles part cannot take the automatic date format without the model.</summary>
        internal sealed class DateStyleUnavailableException : Exception {
        }
    }
}
