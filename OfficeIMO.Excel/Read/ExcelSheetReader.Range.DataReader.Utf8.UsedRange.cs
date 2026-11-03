#nullable enable

using System.Threading;
#if NET8_0_OR_GREATER
using System.Numerics;
using System.Runtime.InteropServices;
using System.Runtime.Intrinsics;
#endif

namespace OfficeIMO.Excel {
    internal sealed partial class ExcelSheetReader {
        private sealed partial class ExcelUtf8RangeRowSource {
            /// <summary>
            /// Proposes bounds for an undimensioned worksheet from its first populated
            /// row and final explicit row tag, or the number of unprefixed row tags
            /// when the final row omits its coordinate. These are only an indexing candidate:
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
                    || !TryGetAttribute(finalRow, "r", out bool hasFinalReference, out int finalStart, out int finalLength)) {
                    return false;
                }
                lastRow = hasFinalReference
                    ? ParsePositiveInt(_buffer!, finalStart, finalLength)
                    : CountUndeclaredRowCandidate(ct);
                if (lastRow <= 0) return false;

                int position = 0;
                bool inSheetData = false;
                int nextImplicitRow = 1;
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
                    if (!TryGetAttribute(tag, "r", out bool hasReference, out int start, out int length)
                        || (firstRow = hasReference ? ParsePositiveInt(_buffer!, start, length) : nextImplicitRow) <= 0
                        || lastRow < firstRow) return false;
                    nextImplicitRow = firstRow + 1;
                    if (tag.IsEmpty) continue;

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

            /// <summary>
            /// Counts possible implicit rows without decoding their cell payloads. Tokens in
            /// comments or nested content can overestimate this proposal; the complete index
            /// and XML qualification still prove coordinates, bounds and well-formedness.
            /// Explicit coordinates outside the proposal retain normal range discovery.
            /// </summary>
            private int CountUndeclaredRowCandidate(CancellationToken ct) {
                ReadOnlySpan<byte> document = _buffer!.AsSpan(0, _length);
                int position = 0;
                int count = 0;
#if NET8_0_OR_GREATER
                if (Vector128.IsHardwareAccelerated) {
                    // Keep multiple candidates from one vector instead of restarting
                    // the span search for each row. The scalar tail also covers tokens
                    // crossing the final vector and tokens without a following byte.
                    ref byte start = ref MemoryMarshal.GetReference(document);
                    Vector128<byte> head = Vector128.Create((byte)'<');
                    Vector128<byte> tail = Vector128.Create((byte)'w');
                    int lastVectorStart = document.Length - Vector128<byte>.Count - 4;
                    for (; position <= lastVectorStart; position += Vector128<byte>.Count) {
                        if ((position & 0xFFFF) == 0) ct.ThrowIfCancellationRequested();
                        uint matches = Vector128.BitwiseAnd(
                            Vector128.Equals(Vector128.LoadUnsafe(ref start, (nuint)position), head),
                            Vector128.Equals(Vector128.LoadUnsafe(ref start, (nuint)(position + 3)), tail))
                            .ExtractMostSignificantBits();
                        while (matches != 0) {
                            int candidate = position + BitOperations.TrailingZeroCount(matches);
                            matches &= matches - 1;
                            if (document[candidate + 1] == (byte)'r'
                                && document[candidate + 2] == (byte)'o'
                                && document[candidate + 4] is (byte)'>' or (byte)'/' or (byte)' ' or (byte)'\t' or (byte)'\r' or (byte)'\n') {
                                if (++count > ExcelA1.MaxRows) return 0;
                            }
                        }
                    }
                }
#endif
                while (position < document.Length) {
                    ct.ThrowIfCancellationRequested();
                    int relative = document.Slice(position).IndexOf("<row"u8);
                    if (relative < 0) break;
                    position += relative + 4;
                    if (position < document.Length
                        && document[position] is (byte)'>' or (byte)'/' or (byte)' ' or (byte)'\t' or (byte)'\r' or (byte)'\n') {
                        if (++count > ExcelA1.MaxRows) return 0;
                    }
                }
                return count;
            }
        }
    }
}
