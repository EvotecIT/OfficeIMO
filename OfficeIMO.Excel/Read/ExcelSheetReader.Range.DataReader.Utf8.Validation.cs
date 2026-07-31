#nullable enable

using System.Threading;
using System.Xml;

namespace OfficeIMO.Excel {
    internal sealed partial class ExcelSheetReader {
        private sealed partial class ExcelUtf8RangeRowSource {
            private void ValidateBufferedWorksheetXml(CancellationToken ct) {
                using var stream = new MemoryStream(
                    _buffer!,
                    0,
                    _length,
                    writable: false,
                    publiclyVisible: false);
                using XmlReader reader = OpenWorksheetXmlReader(stream);
                int nodesUntilCancellationCheck = 1024;
                int sheetDataDepth = -1;
                int rowDepth = -1;
                int cellDepth = -1;
                int inlineStringDepth = -1;
                bool indexedSheetDataCompleted = false;
                ct.ThrowIfCancellationRequested();
                while (reader.Read()) {
                    if (--nodesUntilCancellationCheck == 0) {
                        ct.ThrowIfCancellationRequested();
                        nodesUntilCancellationCheck = 1024;
                    }

                    if (reader.NodeType == XmlNodeType.Element) {
                        bool spreadsheetElement = IsSpreadsheetNamespace(reader.NamespaceURI);
                        if (reader.LocalName == "sheetData") {
                            if (indexedSheetDataCompleted) {
                                continue;
                            }
                            if (!spreadsheetElement) {
                                ThrowUnexpectedIndexedElementNamespace(reader.LocalName);
                            }
                            if (reader.IsEmptyElement) {
                                indexedSheetDataCompleted = true;
                            } else {
                                sheetDataDepth = reader.Depth;
                            }
                            continue;
                        }

                        if (sheetDataDepth >= 0
                            && reader.Depth == sheetDataDepth + 1
                            && reader.LocalName == "row") {
                            if (!spreadsheetElement) {
                                ThrowUnexpectedIndexedElementNamespace(reader.LocalName);
                            }
                            rowDepth = reader.IsEmptyElement ? -1 : reader.Depth;
                            continue;
                        }

                        if (rowDepth >= 0
                            && reader.Depth == rowDepth + 1
                            && reader.LocalName == "c") {
                            if (!spreadsheetElement) {
                                ThrowUnexpectedIndexedElementNamespace(reader.LocalName);
                            }
                            cellDepth = reader.IsEmptyElement ? -1 : reader.Depth;
                            inlineStringDepth = -1;
                            continue;
                        }

                        if (cellDepth >= 0 && reader.Depth == cellDepth + 1) {
                            if (reader.LocalName == "v" || reader.LocalName == "f" || reader.LocalName == "is") {
                                if (!spreadsheetElement) {
                                    ThrowUnexpectedIndexedElementNamespace(reader.LocalName);
                                }
                                if (reader.LocalName == "is" && !reader.IsEmptyElement) {
                                    inlineStringDepth = reader.Depth;
                                }
                            }
                            continue;
                        }

                        if (inlineStringDepth >= 0
                            && reader.Depth == inlineStringDepth + 1
                            && reader.LocalName == "t"
                            && !spreadsheetElement) {
                            ThrowUnexpectedIndexedElementNamespace(reader.LocalName);
                        }
                        continue;
                    }

                    if (reader.NodeType != XmlNodeType.EndElement) {
                        continue;
                    }
                    if (reader.Depth == inlineStringDepth && reader.LocalName == "is") {
                        inlineStringDepth = -1;
                    }
                    if (reader.Depth == cellDepth && reader.LocalName == "c") {
                        cellDepth = -1;
                        inlineStringDepth = -1;
                    }
                    if (reader.Depth == rowDepth && reader.LocalName == "row") {
                        rowDepth = -1;
                    }
                    if (reader.Depth == sheetDataDepth && reader.LocalName == "sheetData") {
                        sheetDataDepth = -1;
                        indexedSheetDataCompleted = true;
                    }
                }
                ct.ThrowIfCancellationRequested();
            }

            private static bool IsSpreadsheetNamespace(string namespaceUri) =>
                string.Equals(namespaceUri, SpreadsheetNamespace, StringComparison.Ordinal)
                || string.Equals(namespaceUri, StrictSpreadsheetNamespace, StringComparison.Ordinal);

            private void ThrowUnexpectedIndexedElementNamespace(string localName) {
                throw new InvalidDataException(
                    $"Worksheet '{_owner._sheetName}' element '{localName}' must use a SpreadsheetML namespace.");
            }

            private bool TryIndexCell(
                ref int position,
                int cellIndex,
                Utf8CellKind kind,
                out bool sharedFormulaFollower,
                out bool hasCachedValue,
                out int valueStart,
                out int valueLength) {
                sharedFormulaFollower = false;
                hasCachedValue = false;
                valueStart = -1;
                valueLength = -1;
                while (TryReadNextTag(ref position, _length, out Utf8Tag tag)) {
                    if (tag.IsEnd && LocalNameEquals(tag, "c")) {
                        return true;
                    }

                    if (!tag.IsEnd && kind == Utf8CellKind.InlineString && LocalNameEquals(tag, "is")) {
                        return TryIndexSimpleInlineStringCell(ref position, tag, cellIndex);
                    }

                    if (tag.IsEnd || (!LocalNameEquals(tag, "v") && !LocalNameEquals(tag, "f"))) {
                        return false;
                    }

                    if (tag.IsEmpty) {
                        if (cellIndex >= 0) {
                            if (LocalNameEquals(tag, "v")) {
                                _valueStarts![cellIndex] = tag.End;
                                _valueLengths![cellIndex] = 0;
                            } else {
                                _formulaStarts![cellIndex] = tag.End;
                                _formulaLengths![cellIndex] = 0;
                            }
                        }
                        if (LocalNameEquals(tag, "v")) {
                            hasCachedValue = true;
                            valueStart = tag.End;
                            valueLength = 0;
                        } else if (IsSharedFormulaTag(tag)) {
                            sharedFormulaFollower = true;
                        }
                        continue;
                    }

                    if (!TryReadNextTag(ref position, _length, out Utf8Tag endTag)
                        || !endTag.IsEnd
                        || !LocalNamesEqual(tag, endTag)
                        || ContainsByte(tag.End + 1, endTag.Start, (byte)'<')) {
                        return false;
                    }

                    int contentStart = tag.End + 1;
                    int contentLength = Math.Max(0, endTag.Start - contentStart);
                    if (cellIndex >= 0) {
                        if (LocalNameEquals(tag, "v")) {
                            _valueStarts![cellIndex] = contentStart;
                            _valueLengths![cellIndex] = contentLength;
                        } else {
                            _formulaStarts![cellIndex] = contentStart;
                            _formulaLengths![cellIndex] = contentLength;
                        }
                    }
                    if (LocalNameEquals(tag, "v")) {
                        hasCachedValue = true;
                        valueStart = contentStart;
                        valueLength = contentLength;
                    } else if (IsSharedFormulaTag(tag)
                               && !ContainsNonWhitespace(contentStart, endTag.Start)) {
                        sharedFormulaFollower = true;
                    }
                }

                return false;
            }

            private bool IsSharedFormulaTag(Utf8Tag tag) {
                if (!LocalNameEquals(tag, "f")
                    || !TryGetAttribute(tag, "t", out bool hasType, out int typeStart, out int typeLength)
                    || !hasType) {
                    return false;
                }

                if (AsciiEqualsIgnoreCase(typeStart, typeLength, "shared")) {
                    return true;
                }

                return ContainsByte(typeStart, typeStart + typeLength, (byte)'&')
                    && string.Equals(
                        DecodeXmlText(typeStart, typeLength),
                        "shared",
                        StringComparison.OrdinalIgnoreCase);
            }

            private void ValidateIndexedCell(
                int rowIndex,
                int columnIndex,
                Utf8CellKind kind,
                int styleIndex,
                bool sharedFormulaFollower,
                bool hasCachedValue,
                int valueStart,
                int valueLength) {
                if (styleIndex >= 0 && (uint)styleIndex >= (uint)_owner.Styles.CellFormatCount) {
                    string reference = A1.CellReference(rowIndex, columnIndex);
                    throw new InvalidDataException(
                        $"Worksheet '{_owner._sheetName}' cell {reference} references a missing cell style.");
                }

                if (kind == Utf8CellKind.SharedString) {
                    ReadOnlySpan<byte> value = valueLength >= 0
                        ? _buffer!.AsSpan(valueStart, valueLength)
                        : ReadOnlySpan<byte>.Empty;
                    var sharedStrings = _owner._sharedStringItems ??= _owner._sst.GetItems();
                    bool parsed = TryParseInt32(value, out int index);
                    if (!parsed && valueLength >= 0) {
                        parsed = TryParseSharedStringIndex(
                            DecodeString(valueStart, valueLength),
                            out index);
                    }
                    if (!parsed || (uint)index >= (uint)sharedStrings.Count) {
                        string reference = A1.CellReference(rowIndex, columnIndex);
                        throw new InvalidDataException(
                            $"Worksheet '{_owner._sheetName}' cell {reference} references a missing shared string.");
                    }
                }

                if (sharedFormulaFollower
                    && (!_options.UseCachedFormulaResult || !hasCachedValue)) {
                    string reference = A1.CellReference(rowIndex, columnIndex);
                    throw new NotSupportedException(
                        $"Data-reader projection cannot safely expand the shared-formula follower " +
                        $"'{_owner._sheetName}'!{reference}. Read the workbook through ExcelDocument when resolved " +
                        "shared-formula text is required.");
                }
            }
        }
    }
}
