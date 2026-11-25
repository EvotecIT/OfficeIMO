using System.Globalization;
using System.Text;
using OfficeIMO.Core.Internal.OpenXml;

namespace OfficeIMO.Excel {
    /// <summary>
    /// A worksheet's recorded changes in the order the save merges them: by row, then column, then the order they
    /// were made. Several changes to one cell are one group; its last value wins and its style changes apply in turn.
    /// </summary>
    internal sealed class ExcelDeferredSheetCells {
        private readonly ExcelDeferredCellStore _store;
        private readonly int[]? _order;

        internal ExcelDeferredSheetCells(ExcelDeferredCellStore store, Func<string, int> sharedStringIndex, Func<DateTime, double> serial) {
            _store = store;
            _order = store.CellOrder();
            SharedStringIndex = sharedStringIndex;
            Serial = serial;
        }

        internal int Count => _store.Count;

        internal ref ExcelDeferredCellEdit At(int position) => ref _store[_order != null ? _order[position] : position];

        /// <summary>The shared-string index of a text, or -1 when the text is written into the cell.</summary>
        internal Func<string, int> SharedStringIndex { get; }

        internal Func<DateTime, double> Serial { get; }
    }

    /// <summary>
    /// Writes a worksheet with recorded changes by streaming its source: rows and cells the changes do not touch are
    /// copied as they were read, changed cells are written as the model's save writes plain value cells, and the
    /// dimension covers every cell, as the model's save computes it. Markup the merge does not understand where it
    /// would have to change something returns null, and the save goes through the SDK model instead.
    /// </summary>
    internal static class ExcelDeferredWorksheetWriter {
        private static readonly UTF8Encoding Utf8 = new UTF8Encoding(encoderShouldEmitUTF8Identifier: false);

        /// <summary>
        /// Returns the merged worksheet in a buffer rented from the shared array pool, or null when it must be written
        /// through the model. <paramref name="applyStyle"/>
        /// gives the style a cell takes from a style change, or null when the model must make it.
        /// </summary>
        internal static byte[]? TryMerge(byte[] xml, int length, ExcelDeferredSheetCells cells, Func<uint, ExcelStyleOperation, uint?> applyStyle,
            out int resultLength, out int addedSharedReferences, out int removedSharedReferences, out string? reason) {
            resultLength = 0;
            addedSharedReferences = removedSharedReferences = 0;
            reason = null;
            var reader = new OfficeXmlPullReader(xml, length, processMarkupCompatibility: false);
            var merge = new Merge(xml, length, cells, applyStyle);
            try {
                if (!merge.Run(reader, out resultLength)) {
                    reason = merge.Reason ?? "the worksheet markup needs the model";
                    merge.ReturnOutput();
                    return null;
                }
                addedSharedReferences = merge.AddedSharedReferences;
                removedSharedReferences = merge.RemovedSharedReferences;
                return merge.Output;
            } catch (System.Xml.XmlException exception) {
                reason = "the worksheet is not well-formed XML: " + exception.Message;
                merge.ReturnOutput();
                return null;
            } catch (StyleUnavailableException) {
                reason = "a style change needs the model to complete the stylesheet";
                merge.ReturnOutput();
                return null;
            }
        }

        private sealed class Merge {
            private readonly byte[] _source;
            private readonly int _sourceLength;
            private readonly ExcelDeferredSheetCells _cells;
            private readonly Func<uint, ExcelStyleOperation, uint?> _applyStyle;
            private int _position;
            private byte[] _output;
            private int _outputLength;
            private int _copied;
            // Column ranges and styles of the first <cols>, which new cells inherit when their row has no format.
            private readonly List<(int Min, int Max, uint Style)> _columns = new();
            private bool _columnsSeen;
            private int _minRow = int.MaxValue, _maxRow, _minColumn = int.MaxValue, _maxColumn;

            internal Merge(byte[] source, int sourceLength, ExcelDeferredSheetCells cells, Func<uint, ExcelStyleOperation, uint?> applyStyle) {
                _source = source;
                _sourceLength = sourceLength;
                _cells = cells;
                _applyStyle = applyStyle;
                _output = System.Buffers.ArrayPool<byte>.Shared.Rent(Math.Max(256, sourceLength + cells.Count * 48 + 1024));
            }

            internal byte[] Output => _output;

            /// <summary>Why the merge left the worksheet to the model.</summary>
            internal string? Reason { get; private set; }

            /// <summary>Written cells that hold a shared string.</summary>
            internal int AddedSharedReferences { get; private set; }

            /// <summary>Changed cells that held a shared string before the change.</summary>
            internal int RemovedSharedReferences { get; private set; }

            private bool Fail(string reason) {
                Reason ??= reason;
                return false;
            }

            // Encoded tag fragments for the prefix last written with, so a cell is written with byte copies.
            private string? _fragmentPrefix;
            private byte[] _cellOpen = Array.Empty<byte>(), _valueOpen = Array.Empty<byte>(), _valueClose = Array.Empty<byte>(),
                _rowOpen = Array.Empty<byte>(), _rowClose = Array.Empty<byte>();

            private void UseFragments(string prefix) {
                if (ReferenceEquals(prefix, _fragmentPrefix) || prefix == _fragmentPrefix) return;
                _fragmentPrefix = prefix;
                string p = prefix.Length == 0 ? string.Empty : prefix + ":";
                _cellOpen = Utf8.GetBytes("<" + p + "c r=\"");
                _valueOpen = Utf8.GetBytes("<" + p + "v>");
                _valueClose = Utf8.GetBytes("</" + p + "v></" + p + "c>");
                _rowOpen = Utf8.GetBytes("<" + p + "row r=\"");
                _rowClose = Utf8.GetBytes("</" + p + "row>");
            }

            private static readonly byte[] StyleOpen = Encoding.ASCII.GetBytes(" s=\"");
            private static readonly byte[] TypeShared = Encoding.ASCII.GetBytes("\" t=\"s\"");
            private static readonly byte[] TypeInline = Encoding.ASCII.GetBytes("\" t=\"str\"");
            private static readonly byte[] TypeNumber = Encoding.ASCII.GetBytes("\" t=\"n\"");
            private static readonly byte[] TypeBoolean = Encoding.ASCII.GetBytes("\" t=\"b\"");
            private static readonly byte[] EmptyClose = Encoding.ASCII.GetBytes("\" />");

            private void AppendFragment(byte[] fragment) => AppendBytes(fragment, 0, fragment.Length);

            private bool HasPending => _position < _cells.Count;

            private int PendingRow => _cells.At(_position).Row;

            private int PendingColumn => _cells.At(_position).Column;

            /// <summary>Takes the changes of the next cell: positions <paramref name="start"/> up to <paramref name="end"/>.</summary>
            private void NextGroup(out int start, out int end) {
                start = _position;
                ref ExcelDeferredCellEdit first = ref _cells.At(_position);
                int row = first.Row, column = first.Column;
                do {
                    _position++;
                } while (_position < _cells.Count && _cells.At(_position).Row == row && _cells.At(_position).Column == column);
                end = _position;
            }

            internal bool Run(OfficeXmlPullReader reader, out int resultLength) {
                resultLength = 0;
                string mainPrefix = string.Empty;
                // Where the dimension's reference value is in the output, or where a new dimension element goes.
                int dimensionValueAt = -1, dimensionInsertAt = -1;
                bool rootSeen = false, dimensionSeen = false, sheetDataSeen = false, inColumns = false;
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
                                if (!reader.IsEmptyElement) reader.SkipElement();
                                else reader.Read();
                                dimensionInsertAt = reader.MarkupEnd;
                                continue;
                            }
                            if (reader.LocalNameEquals("dimension")) {
                                // The model removes extra dimensions and adds a missing reference; those stay with it.
                                if (dimensionSeen || !TryFindAttribute(reader, "ref", out int start, out int valueLength)) {
                                    return Fail("the worksheet has more than one dimension or one without a reference");
                                }
                                dimensionSeen = true;
                                CopyTo(start);
                                dimensionValueAt = _outputLength;
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
                                    if (HasPending) {
                                        // <sheetData/> becomes an element holding the new rows.
                                        CopyTo(reader.MarkupStart);
                                        AppendTag("<", prefix, "sheetData", ">");
                                        WritePendingRowsBefore(int.MaxValue, prefix);
                                        AppendTag("</", prefix, "sheetData", ">");
                                        _copied = reader.MarkupEnd;
                                    }
                                    continue;
                                }
                                if (!ReadSheetData(reader, prefix, ref previousRow)) return false;
                                continue;
                            }
                            continue;
                        }
                        if (depth == 3 && inColumns && reader.IsElement(OfficeXmlNamespaces.SpreadsheetML, "col")) ReadColumn(reader);
                        continue;
                    }
                    if (reader.NodeKind == OfficeXmlNodeKind.EndElement && depth == 2) inColumns = false;
                }

                if (!rootSeen || !sheetDataSeen) return Fail("the worksheet has no sheetData");
                if (HasPending) return Fail("the worksheet ends before its rows");
                CopyTo(_sourceLength);

                string reference = _maxRow == 0 || _maxColumn == 0 ? "A1" : Range(_minRow, _minColumn, _maxRow, _maxColumn);
                if (dimensionSeen) {
                    Splice(dimensionValueAt, 0, Encoding.ASCII.GetBytes(reference));
                } else {
                    // Nothing before the insertion point changed, so it sits where it sat in the source.
                    Splice(dimensionInsertAt, 0, Utf8.GetBytes("<" + Qualified(mainPrefix, "dimension") + " ref=\"" + reference + "\" />"));
                }
                resultLength = _outputLength;
                return true;
            }

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
                        WritePendingRowsBefore(int.MaxValue, sheetDataPrefix);
                        return true;
                    }
                    if (reader.NodeKind != OfficeXmlNodeKind.StartElement || reader.Depth != 3) continue;
                    if (!reader.IsElement(OfficeXmlNamespaces.SpreadsheetML, "row")) {
                        // Anything else between rows leaves the place of a new row unclear.
                        if (HasPending) return Fail("sheetData holds markup other than rows");
                        reader.SkipElement();
                        continue;
                    }
                    int rowNumber;
                    if (TryFindAttribute(reader, "r", out int rowStart, out int rowLength)) {
                        if (!TryParseNumber(rowStart, rowLength, out rowNumber) || rowNumber < 1 || rowNumber > ExcelA1.MaxRows) {
                            return Fail("a row number is not valid");
                        }
                    } else {
                        rowNumber = previousRow + 1;
                    }
                    // Rows out of order leave the place of a new row unclear.
                    if (rowNumber <= previousRow) return Fail("rows are out of order");
                    previousRow = rowNumber;

                    string rowPrefix = reader.ElementPrefix;
                    if (HasPending && PendingRow < rowNumber) {
                        CopyTo(reader.MarkupStart);
                        WritePendingRowsBefore(rowNumber, rowPrefix);
                    }
                    bool edited = HasPending && PendingRow == rowNumber;
                    if (!ReadRow(reader, rowNumber, rowPrefix, edited)) return false;
                }
                return false;
            }

            private bool ReadRow(OfficeXmlPullReader reader, int rowNumber, string rowPrefix, bool edited) {
                uint? rowStyle = null;
                if (edited && reader.TryGetAttribute(OfficeXmlNamespaces.None, "customFormat", out string customFormat)
                    && customFormat is "1" or "true"
                    && reader.TryGetAttribute(OfficeXmlNamespaces.None, "s", out string rowStyleText)) {
                    if (!uint.TryParse(rowStyleText, NumberStyles.None, CultureInfo.InvariantCulture, out uint parsed)) return Fail("a row style is not a number");
                    rowStyle = parsed;
                }

                if (reader.IsEmptyElement) {
                    int tagStart = reader.MarkupStart, tagEnd = reader.MarkupEnd;
                    reader.Read();
                    if (!edited) return true;
                    // <row .../> becomes an element holding its new cells.
                    int slash = tagEnd - 2;
                    while (slash > tagStart && _source[slash] != (byte)'/') slash--;
                    if (_source[slash] != (byte)'/') return Fail("an empty row element is malformed");
                    CopyTo(slash);
                    Append((byte)'>');
                    _copied = tagEnd;
                    WritePendingCells(rowNumber, int.MaxValue, rowPrefix, rowStyle);
                    AppendTag("</", rowPrefix, "row", ">");
                    return true;
                }

                int ordinal = 0, nextColumn = 1, previousColumn = 0;
                while (reader.Read()) {
                    if (reader.NodeKind == OfficeXmlNodeKind.EndElement && reader.Depth == 3) {
                        if (edited && HasPending && PendingRow == rowNumber) {
                            CopyTo(reader.MarkupStart);
                            WritePendingCells(rowNumber, int.MaxValue, rowPrefix, rowStyle);
                        }
                        return true;
                    }
                    if (reader.NodeKind != OfficeXmlNodeKind.StartElement || reader.Depth != 4) continue;
                    if (!reader.IsElement(OfficeXmlNamespaces.SpreadsheetML, "c")) {
                        if (edited) return Fail("an edited row holds markup other than cells");
                        reader.SkipElement();
                        continue;
                    }
                    ordinal++;
                    bool hasReference = TryFindAttribute(reader, "r", out int referenceStart, out int referenceLength);
                    int resolvedRow = rowNumber, resolvedColumn = 0;
                    bool canonical = false;
                    if (hasReference) {
                        if (TryParseCanonicalReference(referenceStart, referenceLength, out int parsedRow, out int parsedColumn)) {
                            resolvedRow = parsedRow;
                            resolvedColumn = parsedColumn;
                            canonical = true;
                        } else {
                            string reference = Utf8.GetString(_source, referenceStart, referenceLength);
                            if (ExcelA1.TryParseCellReferenceFast(reference, out parsedRow, out parsedColumn)) {
                                if (parsedRow > 0) resolvedRow = parsedRow;
                                if (parsedColumn > 0) resolvedColumn = parsedColumn;
                            }
                        }
                    }
                    if (resolvedColumn <= 0) resolvedColumn = !hasReference || referenceLength == 0 ? nextColumn : ordinal;
                    nextColumn = resolvedColumn + 1;
                    if (resolvedRow > 0 && resolvedRow <= ExcelA1.MaxRows && resolvedColumn > 0 && resolvedColumn <= ExcelA1.MaxColumns) {
                        Include(resolvedRow, resolvedColumn);
                    }

                    if (!edited) {
                        reader.SkipElement();
                        continue;
                    }
                    // An edited row is merged by column, which needs canonical references in order.
                    if (!canonical || resolvedRow != rowNumber || resolvedColumn <= previousColumn) {
                        return Fail("an edited row has cells without canonical references in order");
                    }
                    previousColumn = resolvedColumn;

                    int cellStart = reader.MarkupStart;
                    if (HasPending && PendingRow == rowNumber && PendingColumn < resolvedColumn) {
                        CopyTo(cellStart);
                        WritePendingCells(rowNumber, resolvedColumn, rowPrefix, rowStyle);
                    }
                    if (!HasPending || PendingRow != rowNumber || PendingColumn != resolvedColumn) {
                        reader.SkipElement();
                        continue;
                    }

                    NextGroup(out int start, out int end);
                    if (LastValue(start, end) < 0) {
                        // A style change keeps the cell as it is apart from its style attribute.
                        if (!TryRestyleCell(reader, start, end)) return false;
                        continue;
                    }
                    // The edited cell keeps its style and formula, as the model does; its cached value and any inline
                    // string are replaced. Rich-value metadata and extensions need the model.
                    if (!TryReadEditableCell(reader, out uint? existingStyle, out string extraAttributes)) {
                        return Fail("an edited cell holds rich-value metadata or unknown attributes");
                    }
                    if (reader.TryGetAttribute(OfficeXmlNamespaces.None, "t", out string existingType) && existingType == "s") RemovedSharedReferences++;
                    string cellPrefix = reader.ElementPrefix;
                    int formulaStart = -1, formulaEnd = -1;
                    if (!reader.IsEmptyElement) {
                        while (reader.Read()) {
                            if (reader.NodeKind == OfficeXmlNodeKind.EndElement && reader.Depth == 4) break;
                            if (reader.NodeKind != OfficeXmlNodeKind.StartElement || reader.Depth != 5) continue;
                            if (reader.IsElement(OfficeXmlNamespaces.SpreadsheetML, "f") && formulaStart < 0) {
                                formulaStart = reader.MarkupStart;
                                reader.SkipElement();
                                formulaEnd = reader.MarkupEnd;
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
                    WriteCell(rowNumber, resolvedColumn, start, end, cellPrefix, existingStyle, 0U, extraAttributes, formulaStart, formulaEnd);
                    _copied = cellEnd;
                }
                return false;
            }

            /// <summary>Rewrites the style attribute of the cell the reader is on, leaving the rest of its markup as it was.</summary>
            private bool TryRestyleCell(OfficeXmlPullReader reader, int start, int end) {
                uint existingStyle = 0U;
                int styleStart = -1, styleLength = 0;
                if (TryFindAttribute(reader, "s", out int found, out int foundLength)) {
                    if (!TryParseNumber(found, foundLength, out int parsed)) return Fail("a cell style is not a number");
                    existingStyle = (uint)parsed;
                    styleStart = found;
                    styleLength = foundLength;
                }
                uint style = ResolveStyle(start, end, existingStyle)!.Value;
                if (styleStart >= 0) {
                    CopyTo(styleStart);
                    AppendNumber(style);
                    _copied = styleStart + styleLength;
                } else {
                    // The reference comes first; the style follows it, as the model writes it.
                    if (!TryFindAttribute(reader, "r", out int referenceStart, out int referenceLength)) return Fail("a restyled cell has no reference");
                    CopyTo(referenceStart + referenceLength + 1);
                    AppendAscii(" s=\"");
                    AppendNumber(style);
                    Append((byte)'"');
                }
                reader.SkipElement();
                return true;
            }

            /// <summary>The position of the cell's last value change, or -1 when it only changes style.</summary>
            private int LastValue(int start, int end) {
                for (int position = end - 1; position >= start; position--) {
                    if (_cells.At(position).Kind != ExcelDeferredCellKind.Style) return position;
                }
                return -1;
            }

            /// <summary>Applies the cell's style changes to <paramref name="baseStyle"/>; null when there are none.</summary>
            private uint? ResolveStyle(int start, int end, uint baseStyle) {
                uint style = baseStyle;
                bool changed = false;
                for (int position = start; position < end; position++) {
                    ref ExcelDeferredCellEdit edit = ref _cells.At(position);
                    ExcelStyleOperation? operation = edit.Kind switch {
                        ExcelDeferredCellKind.Date => ExcelStyleOperation.Date,
                        ExcelDeferredCellKind.Style => edit.Style,
                        _ => null
                    };
                    if (operation == null) continue;
                    style = _applyStyle(style, operation) ?? throw new StyleUnavailableException();
                    changed = true;
                }
                return changed ? style : null;
            }

            /// <summary>
            /// Reads the style of an edited cell and the attributes it keeps (<c>cm</c>, <c>ph</c>) as written; false for
            /// rich-value metadata, namespaced or unknown attributes.
            /// </summary>
            private bool TryReadEditableCell(OfficeXmlPullReader reader, out uint? style, out string extraAttributes) {
                style = null;
                extraAttributes = string.Empty;
                for (int index = 0; index < reader.AttributeCount; index++) {
                    if (reader.GetAttributeNamespace(index, out _) != OfficeXmlNamespaces.None) return false;
                    reader.GetAttributeRawValue(index, out int start, out int length);
                    switch (reader.GetAttributeLocalName(index)) {
                        case "r":
                        case "t":
                            break;
                        case "s":
                            if (!TryParseNumber(start, length, out int parsed)) return false;
                            style = (uint)parsed;
                            break;
                        case "cm":
                        case "ph":
                            extraAttributes += " " + reader.GetAttributeLocalName(index) + "=\"" + Utf8.GetString(_source, start, length) + "\"";
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

            private void WritePendingRowsBefore(int limit, string prefix) {
                UseFragments(prefix);
                while (HasPending && PendingRow < limit) {
                    int rowNumber = PendingRow;
                    AppendFragment(_rowOpen);
                    AppendLong(rowNumber);
                    AppendAscii("\">");
                    WritePendingCells(rowNumber, int.MaxValue, prefix, rowStyle: null);
                    UseFragments(prefix);
                    AppendFragment(_rowClose);
                }
            }

            /// <summary>Writes the new cells of <paramref name="row"/> before column <paramref name="limit"/>.</summary>
            private void WritePendingCells(int row, int limit, string prefix, uint? rowStyle) {
                while (HasPending && PendingRow == row && PendingColumn < limit) {
                    int column = PendingColumn;
                    NextGroup(out int start, out int end);
                    WriteCell(row, column, start, end, prefix, null, ResolveInheritedStyle(rowStyle, column), string.Empty, -1, -1);
                }
            }

            private void WriteCell(int row, int column, int start, int end, string prefix, uint? existingStyle, uint inheritedStyle,
                string extraAttributes, int formulaStart, int formulaEnd) {
                Include(row, column);
                UseFragments(prefix);
                AppendFragment(_cellOpen);
                AppendReference(row, column);
                uint? style = ResolveStyle(start, end, existingStyle ?? inheritedStyle);
                uint? written = style ?? existingStyle ?? (inheritedStyle != 0U ? inheritedStyle : null);
                if (written.HasValue) {
                    Append((byte)'"');
                    AppendFragment(StyleOpen);
                    AppendLong(written.Value);
                }
                int value = LastValue(start, end);
                if (value < 0) {
                    // A new cell that only takes a style has no value.
                    AppendFragment(EmptyClose);
                    return;
                }
                ref ExcelDeferredCellEdit edit = ref _cells.At(value);
                int sharedIndex = -1;
                switch (edit.Kind) {
                    case ExcelDeferredCellKind.Text:
                        sharedIndex = _cells.SharedStringIndex(edit.Text);
                        if (sharedIndex >= 0) {
                            AppendFragment(TypeShared);
                            AddedSharedReferences++;
                        } else {
                            AppendFragment(TypeInline);
                        }
                        break;
                    case ExcelDeferredCellKind.Boolean:
                        AppendFragment(TypeBoolean);
                        break;
                    default:
                        AppendFragment(TypeNumber);
                        break;
                }
                if (extraAttributes.Length > 0) AppendAscii(extraAttributes);
                Append((byte)'>');
                if (formulaStart >= 0) AppendBytes(_source, formulaStart, formulaEnd - formulaStart);
                AppendFragment(_valueOpen);
                switch (edit.Kind) {
                    case ExcelDeferredCellKind.Text:
                        if (sharedIndex >= 0) AppendLong(sharedIndex);
                        else AppendEscaped(Utilities.ExcelSanitizer.EncodeXString(edit.Text));
                        break;
                    case ExcelDeferredCellKind.Number: AppendDouble(edit.Number); break;
                    case ExcelDeferredCellKind.Integer: AppendLong((long)edit.Number); break;
                    case ExcelDeferredCellKind.Decimal: AppendAscii(edit.Text); break;
                    case ExcelDeferredCellKind.Boolean: Append(edit.Number != 0D ? (byte)'1' : (byte)'0'); break;
                    default: AppendDouble(_cells.Serial(edit.Date)); break;
                }
                AppendFragment(_valueClose);
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

            /// <summary>Reads a non-negative decimal integer written without leading zeros, sign or spaces.</summary>
            private bool TryParseNumber(int start, int length, out int value) {
                value = 0;
                if (length == 0 || length > 10 || length > 1 && _source[start] == (byte)'0') return false;
                long result = 0;
                for (int index = start; index < start + length; index++) {
                    int digit = _source[index] - '0';
                    if ((uint)digit > 9U) return false;
                    result = result * 10 + digit;
                }
                if (result > int.MaxValue) return false;
                value = (int)result;
                return true;
            }

            /// <summary>Reads a reference written as upper-case column letters followed by the row number, as the model writes it.</summary>
            private bool TryParseCanonicalReference(int start, int length, out int row, out int column) {
                row = column = 0;
                int index = start, end = start + length;
                while (index < end && _source[index] >= (byte)'A' && _source[index] <= (byte)'Z') {
                    column = column * 26 + (_source[index] - 'A' + 1);
                    if (column > ExcelA1.MaxColumns) return false;
                    index++;
                }
                if (index == start || index - start > 3) return false;
                return TryParseNumber(index, end - index, out row) && row >= 1 && row <= ExcelA1.MaxRows;
            }

            private static string Qualified(string prefix, string localName) => prefix.Length == 0 ? localName : prefix + ":" + localName;

            private void CopyTo(int sourcePosition) {
                if (sourcePosition <= _copied) return;
                AppendBytes(_source, _copied, sourcePosition - _copied);
                _copied = sourcePosition;
            }

            private void AppendBytes(byte[] bytes, int start, int count) {
                Ensure(count);
                Buffer.BlockCopy(bytes, start, _output, _outputLength, count);
                _outputLength += count;
            }

            private void Append(byte value) {
                Ensure(1);
                _output[_outputLength++] = value;
            }

            private void AppendAscii(string text) {
                Ensure(Utf8.GetMaxByteCount(text.Length));
                _outputLength += Utf8.GetBytes(text, 0, text.Length, _output, _outputLength);
            }

            private void AppendQualified(string prefix, string localName) {
                if (prefix.Length > 0) {
                    AppendAscii(prefix);
                    Append((byte)':');
                }
                AppendAscii(localName);
            }

            private void AppendTag(string open, string prefix, string localName, string close) {
                AppendAscii(open);
                AppendQualified(prefix, localName);
                AppendAscii(close);
            }

            private void AppendEscaped(string value) {
                if (value.IndexOf('&') < 0 && value.IndexOf('<') < 0 && value.IndexOf('>') < 0) {
                    AppendAscii(value);
                    return;
                }
                var escaped = new StringBuilder(value.Length + 16);
                foreach (char character in value) {
                    switch (character) {
                        case '&': escaped.Append("&amp;"); break;
                        case '<': escaped.Append("&lt;"); break;
                        case '>': escaped.Append("&gt;"); break;
                        default: escaped.Append(character); break;
                    }
                }
                AppendAscii(escaped.ToString());
            }

            private void AppendNumber(uint value) => AppendLong(value);

            private void AppendLong(long value) {
                Ensure(21);
                if (value < 0) {
                    _output[_outputLength++] = (byte)'-';
                    value = -value;
                }
                int digits = 1;
                for (long rest = value / 10; rest > 0; rest /= 10) digits++;
                int position = _outputLength + digits;
                do {
                    _output[--position] = (byte)('0' + value % 10);
                    value /= 10;
                } while (value > 0);
                _outputLength += digits;
            }

            /// <summary>Writes a double as <see cref="InvariantNumberText.Get(double)"/> writes it.</summary>
            private void AppendDouble(double value) {
                if (value >= int.MinValue && value <= int.MaxValue) {
                    int integer = (int)value;
                    if (value == integer) {
                        AppendLong(integer);
                        return;
                    }
                }
#if NETCOREAPP3_0_OR_GREATER
                Span<char> characters = stackalloc char[32];
                if (value.TryFormat(characters, out int written, "R", CultureInfo.InvariantCulture)) {
                    Ensure(written);
                    for (int index = 0; index < written; index++) _output[_outputLength++] = (byte)characters[index];
                    return;
                }
#endif
                AppendAscii(InvariantNumberText.Get(value));
            }

            private void AppendReference(int row, int column) {
                Ensure(16);
                Span<byte> letters = stackalloc byte[3];
                int count = 0;
                while (column > 0) {
                    int remainder = (column - 1) % 26;
                    letters[count++] = (byte)('A' + remainder);
                    column = (column - 1) / 26;
                }
                while (count > 0) _output[_outputLength++] = letters[--count];
                AppendLong(row);
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
                byte[] larger = System.Buffers.ArrayPool<byte>.Shared.Rent(Math.Max(_output.Length * 2, _outputLength + additional + 1024));
                Buffer.BlockCopy(_output, 0, larger, 0, _outputLength);
                System.Buffers.ArrayPool<byte>.Shared.Return(_output);
                _output = larger;
            }

            /// <summary>Gives the output buffer back when the merge's result is not used.</summary>
            internal void ReturnOutput() => System.Buffers.ArrayPool<byte>.Shared.Return(_output);
        }

        /// <summary>The stylesheet cannot take a style change without the model.</summary>
        internal sealed class StyleUnavailableException : Exception {
        }
    }
}
