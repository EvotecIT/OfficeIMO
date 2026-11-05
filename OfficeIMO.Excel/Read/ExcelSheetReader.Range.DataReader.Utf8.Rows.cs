using System.Threading;

namespace OfficeIMO.Excel {
    internal sealed partial class ExcelSheetReader {
        private sealed partial class ExcelUtf8RangeRowSource {
            // Look ahead in the existing buffer without consuming the row's cells.
            // Unusual references retain the general XML reader's decoding path.
            private bool TryInferImplicitRowIndex(int position, int fallback, CancellationToken ct, out int rowIndex) {
                rowIndex = fallback;
                const int MaximumReferenceLength = 32;
                int depth = 0;
                while (TryReadNextTag(ref position, _length, out Utf8Tag tag)) {
                    ct.ThrowIfCancellationRequested();
                    if (tag.IsEnd) {
                        if (depth == 0) return IsIndexedTag(tag) && LocalNameEquals(tag, "row");
                        depth--;
                        continue;
                    }
                    if (depth == 0 && IsIndexedTag(tag) && LocalNameEquals(tag, "c")) {
                        if (!TryGetAttribute(tag, "r", out bool found, out int start, out int length)) return false;
                        if (found && length > 0) {
                            if (length > MaximumReferenceLength) return false;
                            ReadOnlySpan<byte> reference = _buffer!.AsSpan(start, length);
                            for (int index = 0; index < reference.Length; index++) {
                                byte value = reference[index];
                                if (value > 127 || value == (byte)'&') return false;
                            }
                            if (ExcelA1.TryParseCellReferenceFast(reference, out int referencedRow, out _)) {
                                rowIndex = referencedRow;
                                return true;
                            }
                        }
                    }
                    if (!tag.IsEmpty) depth++;
                }
                return false;
            }
        }
    }
}
