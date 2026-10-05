#nullable enable

using System.Buffers;
using System.Threading;

namespace OfficeIMO.Excel {
    internal sealed partial class ExcelSheetReader {
        private sealed partial class ExcelUtf8RangeRowSource {
            /// <summary>
            /// Declines an oversized used-range indexing proposal using bounded
            /// beginning/end fragments. Fragments never qualify XML or publish
            /// values; the normal streaming projection still validates the part.
            /// </summary>
            private static bool ShouldSkipLargeUsedRangeBuffer(
                ExcelSheetReader owner,
                CancellationToken ct) {
                // A prefetch owns its ZIP stream independently. Do not open a
                // second stream on that reader while its task may still be active.
                if (owner._opt.EnableWorksheetPrefetch
                    || owner._partBufferReader == null
                    || !owner._partBufferReader.TryGetLength(owner._worksheetPartName, out long length)
                    || length <= MaximumBufferSize / 2
                    || length > MaximumBufferSize) {
                    return false;
                }

                using Stream stream = owner._partBufferReader.OpenPart(
                    owner._worksheetPartName, MaximumBufferSize, ct);
                if (!TryReadUtf8WorksheetPrefix(stream, ct, out byte[] prefix, out int prefixLength)) {
                    return true;
                }

                byte[]? fragments = null;
                try {
                    const int window = InitialBufferSize;
                    fragments = ArrayPool<byte>.Shared.Rent(2 * window);
                    Buffer.BlockCopy(prefix, 0, fragments, 0, prefixLength);
                    ReadBudgetFragment(stream, fragments, prefixLength, window - prefixLength, ct);

                    var first = new ExcelUtf8RangeRowSource(owner, fragments, window);
                    if (first.TryGetDeclaredRange(ct, out int firstRow, out int firstColumn,
                        out int lastRow, out int lastColumn)) {
                        return ExceedsUsedRangeIndexBudget(owner, firstRow, firstColumn, lastRow, lastColumn);
                    }

                    int remaining = checked((int)length) - window;
                    while (remaining >= window) {
                        ReadBudgetFragment(stream, fragments, window, window, ct);
                        remaining -= window;
                    }
                    if (remaining != 0) {
                        // Preserve the previous block's suffix when the final
                        // read is shorter than a complete window.
                        Buffer.BlockCopy(fragments, window + remaining,
                            fragments, window, window - remaining);
                        ReadBudgetFragment(stream, fragments, 2 * window - remaining, remaining, ct);
                    }
                    ct.ThrowIfCancellationRequested();
                    if (stream.ReadByte() != -1) {
                        throw new InvalidDataException("Worksheet part exceeds its declared length.");
                    }

                    var probe = new ExcelUtf8RangeRowSource(owner, fragments, 2 * window);
                    if (!probe.TryInferUndeclaredRange(ct, out firstRow, out firstColumn,
                        out lastRow, out lastColumn)) {
                        return false;
                    }
                    return ExceedsUsedRangeIndexBudget(owner, firstRow, firstColumn, lastRow, lastColumn);
                } finally {
                    ArrayPool<byte>.Shared.Return(prefix, clearArray: true);
                    if (fragments != null) ArrayPool<byte>.Shared.Return(fragments, clearArray: true);
                }
            }

            // Hints only choose an optimization. They never replace full XML,
            // grid, style, shared-string, formula or value qualification.
            private static bool ExceedsUsedRangeIndexBudget(ExcelSheetReader owner,
                int firstRow, int firstColumn, int lastRow, int lastColumn) {
                if (firstRow <= 0 || lastRow < firstRow || lastRow > A1.MaxRows
                    || firstColumn <= 0 || lastColumn < firstColumn || lastColumn > A1.MaxColumns) {
                    return false;
                }
                long columns = (long)lastColumn - firstColumn + 1;
                long cells = ((long)lastRow - firstRow + 1) * columns;
                return columns > owner._opt.MaxDataReaderColumns
                    || cells > owner._opt.MaxDataReaderBufferedCells
                    || cells > MaximumIndexedCells;
            }

            private static void ReadBudgetFragment(Stream stream, byte[] buffer,
                int offset, int count, CancellationToken ct) {
                int end = offset + count;
                while (offset < end) {
                    ct.ThrowIfCancellationRequested();
                    int read = stream.Read(buffer, offset, end - offset);
                    if (read == 0) throw new EndOfStreamException("Worksheet part ended before its declared length.");
                    offset += read;
                }
                ct.ThrowIfCancellationRequested();
            }
        }
    }
}
