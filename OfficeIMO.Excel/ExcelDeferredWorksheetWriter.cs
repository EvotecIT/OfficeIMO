using System.Globalization;
using System.Text;
using OfficeIMO.Core.Internal.OpenXml;

namespace OfficeIMO.Excel {
    /// <summary>
    /// What the owned save writes for one cell: its last value, with the shared-string index or text already resolved,
    /// and the style changes in the order they were made. A cell with style changes only keeps its value.
    /// </summary>
    internal sealed class ExcelDeferredCellOutput {
        /// <summary>The cell type attribute: <c>s</c>, <c>str</c>, <c>n</c> or <c>b</c>; null when the value is unchanged.</summary>
        internal string? Type { get; private set; }
        /// <summary>The text of the value element, not yet escaped for XML.</summary>
        internal string? Value { get; private set; }
        internal List<ExcelStyleOperation>? Operations { get; private set; }

        internal bool HasValue => Type != null;

        internal void SetValue(string type, string value, bool date) {
            Type = type;
            Value = value;
            if (date) AddStyle(ExcelStyleOperation.Date);
        }

        internal void AddStyle(ExcelStyleOperation operation) => (Operations ??= new List<ExcelStyleOperation>()).Add(operation);
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
        /// maps row numbers to the edited cells of the row by column. <paramref name="applyStyle"/> gives the style a
        /// cell takes from a style change, or null when the model must make it.
        /// </summary>
        internal static byte[]? TryMerge(byte[] xml, int length, SortedDictionary<int, SortedDictionary<int, ExcelDeferredCellOutput>> rows,
            Func<uint, ExcelStyleOperation, uint?> applyStyle, out int resultLength, out int removedSharedReferences, out string? reason) {
            resultLength = 0;
            removedSharedReferences = 0;
            reason = null;
            var reader = new OfficeXmlPullReader(xml, length, processMarkupCompatibility: false);
            var merge = new Merge(xml, rows, applyStyle);
            try {
                if (!merge.Run(reader, out resultLength)) {
                    reason = merge.Reason ?? "the worksheet markup needs the model";
                    return null;
                }
                removedSharedReferences = merge.RemovedSharedReferences;
                return merge.Output;
            } catch (System.Xml.XmlException exception) {
                reason = "the worksheet is not well-formed XML: " + exception.Message;
                return null;
            } catch (DateStyleUnavailableException) {
                reason = "a style change needs the model to complete the stylesheet";
                return null;
            }
        }

        private sealed class Merge {
            private readonly byte[] _source;
            private readonly SortedDictionary<int, SortedDictionary<int, ExcelDeferredCellOutput>> _rows;
            private readonly Func<uint, ExcelStyleOperation, uint?> _applyStyle;
            private readonly List<int> _pendingRows;
            private int _nextPendingRow;
            private byte[] _output;
            private int _outputLength;
            private int _copied;
            // Column ranges and styles of the first <cols>, which new cells inherit when their row has no format.
            private readonly List<(int Min, int Max, uint Style)> _columns = new();
            private bool _columnsSeen;
            private int _minRow = int.MaxValue, _maxRow, _minColumn = int.MaxValue, _maxColumn;

            internal Merge(byte[] source, SortedDictionary<int, SortedDictionary<int, ExcelDeferredCellOutput>> rows, Func<uint, ExcelStyleOperation, uint?> applyStyle) {
                _source = source;
                _rows = rows;
                _applyStyle = applyStyle;
                _pendingRows = new List<int>(rows.Keys);
                _output = new byte[Math.Max(256, source.Length + 4096)];
            }

            internal byte[] Output => _output;

            /// <summary>Why the merge left the worksheet to the model.</summary>
            internal string? Reason { get; private set; }

            private bool Fail(string reason) {
                Reason ??= reason;
                return false;
            }

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
                                return Fail("the worksheet root is not a Transitional worksheet");
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
                                if (dimensionSeen || !TryFindAttribute(reader, "ref", out int start, out int valueLength)) return Fail("the worksheet has more than one dimension or one without a reference");
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
                                if (sheetDataSeen) return Fail("the worksheet has more than one sheetData");
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

                if (!rootSeen || !sheetDataSeen || inSheetData || _nextPendingRow < _pendingRows.Count) return Fail("the worksheet has no sheetData");
                CopyTo(_source.Length > 0 ? reader.DataLength : 0);

                string reference = _maxRow == 0 || _maxColumn == 0
                    ? "A1"
                    : Range(_minRow, _minColumn, _maxRow, _maxColumn);
                byte[] referenceBytes = Encoding.ASCII.GetBytes(reference);
                if (dimensionSeen) {
                    Splice(dimensionValueAt, dimensionValueLength, referenceBytes);
                } else {
                    if (dimensionInsertAt < 0) return Fail("the worksheet has no place for a dimension");
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
                        if (_nextPendingRow < _pendingRows.Count) return Fail("sheetData holds markup other than rows");
                        reader.SkipElement();
                        continue;
                    }
                    int rowNumber;
                    if (reader.TryGetAttribute(OfficeXmlNamespaces.None, "r", out string rowText)) {
                        if (!int.TryParse(rowText, NumberStyles.None, CultureInfo.InvariantCulture, out rowNumber)
                            || rowNumber < 1 || rowNumber > ExcelA1.MaxRows) return Fail("a row number is not valid");
                    } else {
                        rowNumber = previousRow + 1;
                    }
                    // Rows out of order leave the place of a new row unclear.
                    if (rowNumber <= previousRow) return Fail("rows are out of order");
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
                    if (!uint.TryParse(rowStyleText, NumberStyles.None, CultureInfo.InvariantCulture, out uint parsed)) return Fail("a row style is not a number");
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
                    if (_source[slash] != (byte)'/') return Fail("an empty row element is malformed");
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
                        if (edits != null) return Fail("an edited row holds markup other than cells");
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
                        return Fail("an edited row has cells without canonical references in order");
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

                    ExcelDeferredCellOutput edit = pending.Current.Value;
                    if (!edit.HasValue) {
                        // A style change keeps the cell as it is apart from its style attribute.
                        if (!TryRestyleCell(reader, edit, rowNumber, resolvedColumn)) return false;
                        hasPending = pending.MoveNext();
                        continue;
                    }
                    // The edited cell keeps its style and formula, as the model does; its cached value and any inline
                    // string are replaced. Rich-value metadata and extensions need the model.
                    if (!TryReadEditableCell(reader, out string? existingStyle, out string extraAttributes)) return Fail("an edited cell holds rich-value metadata or unknown attributes");
                    if (reader.TryGetAttribute(OfficeXmlNamespaces.None, "t", out string existingType) && existingType == "s") RemovedSharedReferences++;
                    string cellPrefix = reader.ElementPrefix;
                    string? formula = null;
                    if (!reader.IsEmptyElement) {
                        while (reader.Read()) {
                            if (reader.NodeKind == OfficeXmlNodeKind.EndElement && reader.Depth == 4) break;
                            if (reader.NodeKind != OfficeXmlNodeKind.StartElement || reader.Depth != 5) continue;
                            if (reader.IsElement(OfficeXmlNamespaces.SpreadsheetML, "f") && formula == null) {
                                int formulaStart = reader.MarkupStart;
                                reader.SkipElement();
                                formula = Utf8.GetString(_source, formulaStart, reader.MarkupEnd - formulaStart);
                            } else if (reader.IsElement(OfficeXmlNamespaces.SpreadsheetML, "v") || reader.IsElement(OfficeXmlNamespaces.SpreadsheetML, "is")) {
                                reader.SkipElement();
                            } else {
                                return Fail("an edited cell holds markup other than a formula and a value");
                            }
                        }
                    } else {
                        reader.Read();
                    }
                    int cellEnd = reader.MarkupEnd;
                    CopyTo(cellStart);
                    WriteCell(rowNumber, resolvedColumn, edit, cellPrefix, existingStyle, inheritedStyle: 0U, extraAttributes, formula);
                    _copied = cellEnd;
                    hasPending = pending.MoveNext();
                }
                return false;
            }

            /// <summary>Rewrites the style attribute of the cell the reader is on, leaving the rest of its markup as it was.</summary>
            private bool TryRestyleCell(OfficeXmlPullReader reader, ExcelDeferredCellOutput edit, int row, int column) {
                string? existingStyle = null;
                int styleStart = -1, styleLength = 0;
                for (int index = 0; index < reader.AttributeCount; index++) {
                    if (reader.GetAttributeNamespace(index, out _) == OfficeXmlNamespaces.None && reader.GetAttributeLocalName(index) == "s") {
                        existingStyle = reader.GetAttributeValue(index);
                        if (!uint.TryParse(existingStyle, NumberStyles.None, CultureInfo.InvariantCulture, out _)) return Fail("a cell style is not a number");
                        reader.GetAttributeRawValue(index, out styleStart, out styleLength);
                    }
                }
                uint style = ResolveStyle(edit, existingStyle != null ? uint.Parse(existingStyle, CultureInfo.InvariantCulture) : 0U)!.Value;
                string text = style.ToString(CultureInfo.InvariantCulture);
                if (styleStart >= 0) {
                    CopyTo(styleStart);
                    AppendText(text);
                    _copied = styleStart + styleLength;
                } else {
                    // The reference comes first; the style follows it, as the model writes it.
                    if (!TryFindAttribute(reader, "r", out int referenceStart, out int referenceLength)) return Fail("a restyled cell has no reference");
                    int afterReference = referenceStart + referenceLength + 1;
                    CopyTo(afterReference);
                    AppendText(" s=\"" + text + "\"");
                }
                Include(row, column);
                reader.SkipElement();
                return true;
            }

            /// <summary>Applies the cell's style changes to <paramref name="baseStyle"/>; null when there are none.</summary>
            private uint? ResolveStyle(ExcelDeferredCellOutput edit, uint baseStyle) {
                if (edit.Operations == null) return null;
                uint style = baseStyle;
                foreach (ExcelStyleOperation operation in edit.Operations) {
                    style = _applyStyle(style, operation) ?? throw new DateStyleUnavailableException();
                }
                return style;
            }

            /// <summary>
            /// Reads the style of an edited cell and the attributes it keeps (<c>cm</c>, <c>ph</c>) as written; false for
            /// rich-value metadata, namespaced or unknown attributes.
            /// </summary>
            private static bool TryReadEditableCell(OfficeXmlPullReader reader, out string? style, out string extraAttributes) {
                style = null;
                extraAttributes = string.Empty;
                for (int index = 0; index < reader.AttributeCount; index++) {
                    if (reader.GetAttributeNamespace(index, out _) != OfficeXmlNamespaces.None) return false;
                    switch (reader.GetAttributeLocalName(index)) {
                        case "r":
                        case "t":
                            break;
                        case "s":
                            style = reader.GetAttributeValue(index);
                            if (!uint.TryParse(style, NumberStyles.None, CultureInfo.InvariantCulture, out _)) return false;
                            break;
                        case "cm":
                        case "ph":
                            reader.GetAttributeRawValue(index, out int start, out int length);
                            extraAttributes += " " + reader.GetAttributeLocalName(index) + "=\"" + Utf8.GetString(reader.Data, start, length) + "\"";
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

            private void WriteCell(int row, int column, ExcelDeferredCellOutput value, string prefix, string? existingStyle, uint inheritedStyle,
                string extraAttributes = "", string? formula = null) {
                Include(row, column);
                string cell = Qualified(prefix, "c");
                var text = new StringBuilder(64);
                text.Append('<').Append(cell).Append(" r=\"").Append(ExcelA1.CellReference(row, column)).Append('"');
                uint baseStyle = existingStyle != null ? uint.Parse(existingStyle, CultureInfo.InvariantCulture) : inheritedStyle;
                uint? style = ResolveStyle(value, baseStyle);
                if (style.HasValue) {
                    text.Append(" s=\"").Append(style.Value.ToString(CultureInfo.InvariantCulture)).Append('"');
                } else if (existingStyle != null) {
                    text.Append(" s=\"").Append(existingStyle).Append('"');
                } else if (inheritedStyle != 0U) {
                    text.Append(" s=\"").Append(inheritedStyle.ToString(CultureInfo.InvariantCulture)).Append('"');
                }
                if (!value.HasValue) {
                    // A new cell that only takes a style has no value.
                    text.Append(" />");
                    AppendText(text.ToString());
                    return;
                }
                text.Append(" t=\"").Append(value.Type).Append('"').Append(extraAttributes).Append('>');
                if (formula != null) text.Append(formula);
                text.Append('<').Append(Qualified(prefix, "v")).Append('>');
                AppendEscaped(text, value.Value!);
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
