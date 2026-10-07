#nullable enable

using System.Threading;

namespace OfficeIMO.Excel {
    internal sealed partial class ExcelSheetReader {
        private sealed partial class ExcelUtf8RangeRowSource {
            /// <summary>
            /// Proposes bounds for an undimensioned worksheet from its first populated
            /// row and final explicit row tag. These are only an indexing candidate:
            /// the complete index and XML validation must still prove every cell fits.
            /// </summary>
            private bool TryInferUndeclaredRange(
                CancellationToken ct,
                out int firstRow,
                out int firstColumn,
                out int lastRow,
                out int lastColumn) {
                firstRow = firstColumn = lastRow = lastColumn = 0;
                if (!HasSupportedUtf8Encoding()) return false;

                // Reject a known large rectangular extent before renting its metadata.
                // Unusual final tags retain the existing XML discovery path.
                int lastRowStart = _buffer!.AsSpan(0, _length).LastIndexOf("<row"u8);
                if (lastRowStart < 0
                    || !TryParseTag(lastRowStart, _length, out Utf8Tag finalRow)
                    || finalRow.NameStart != finalRow.LocalNameStart
                    || finalRow.IsEnd
                    || !LocalNameEquals(finalRow, "row")
                    || !TryGetAttribute(finalRow, "r", out bool hasFinalReference, out int finalStart, out int finalLength)
                    || !hasFinalReference
                    || (lastRow = ParsePositiveInt(_buffer!, finalStart, finalLength)) <= 0) {
                    return false;
                }

                int position = 0;
                bool inSheetData = false;
                while (TryReadNextTag(ref position, _length, out Utf8Tag tag)) {
                    ct.ThrowIfCancellationRequested();
                    if (!inSheetData) {
                        if (!tag.IsEnd && LocalNameEquals(tag, "worksheet")) {
                            if (tag.NameStart != tag.LocalNameStart) return false;
                            SetWorksheetPrefix(tag);
                        }
                        if (!tag.IsEnd && IsIndexedTag(tag) && LocalNameEquals(tag, "sheetData")) {
                            if (tag.IsEmpty) return false;
                            inSheetData = true;
                        }
                        continue;
                    }
                    if (tag.IsEnd || !IsIndexedTag(tag) || !LocalNameEquals(tag, "row")) return false;
                    if (tag.IsEmpty) continue;
                    if (!TryGetAttribute(tag, "r", out bool hasReference, out int start, out int length)
                        || !hasReference
                        || (firstRow = ParsePositiveInt(_buffer!, start, length)) <= 0
                        || lastRow < firstRow) return false;

                    int depth = 0;
                    int nextColumn = 1;
                    int previousColumn = 0;
                    while (TryReadNextTag(ref position, _length, out Utf8Tag child)) {
                        ct.ThrowIfCancellationRequested();
                        if (child.IsEnd) {
                            if (depth == 0) {
                                if (!IsIndexedTag(child) || !LocalNameEquals(child, "row")) return false;
                                break;
                            }
                            depth--;
                            continue;
                        }
                        if (depth == 0) {
                            if (!IsIndexedTag(child) || !LocalNameEquals(child, "c")
                                || !TryGetCellAttributes(child, ref nextColumn, out int column, out _, out _)
                                || column <= previousColumn) return false;
                            if (firstColumn == 0) firstColumn = column;
                            lastColumn = previousColumn = column;
                        }
                        if (!child.IsEmpty) depth++;
                    }
                    if (firstColumn > 0) return !_parseFailed;
                }
                return false;
            }
        }
    }
}
