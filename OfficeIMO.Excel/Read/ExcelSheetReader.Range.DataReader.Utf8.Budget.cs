#nullable enable

using System.Buffers;
using System.Threading;

namespace OfficeIMO.Excel {
    internal sealed partial class ExcelSheetReader {
        private sealed partial class ExcelUtf8RangeRowSource {
            /// <summary>
            /// Declines a large declared used-range proposal from a bounded prefix.
            /// The prefix never qualifies XML or publishes values; the normal
            /// streaming projection still validates the complete worksheet.
            /// </summary>
            private static bool ShouldSkipLargeDeclaredRangeBuffer(
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

                byte[]? fragment = null;
                try {
                    fragment = ArrayPool<byte>.Shared.Rent(InitialBufferSize);
                    Buffer.BlockCopy(prefix, 0, fragment, 0, prefixLength);
                    int offset = prefixLength;
                    while (offset < InitialBufferSize) {
                        ct.ThrowIfCancellationRequested();
                        int read = stream.Read(fragment, offset, InitialBufferSize - offset);
                        if (read == 0) throw new EndOfStreamException("Worksheet part ended before its declared length.");
                        offset += read;
                    }
                    ct.ThrowIfCancellationRequested();
                    var probe = new ExcelUtf8RangeRowSource(owner, fragment, InitialBufferSize);
                    if (!probe.TryGetDeclaredRange(ct, out int firstRow, out int firstColumn,
                        out int lastRow, out int lastColumn)
                        || firstRow <= 0 || lastRow < firstRow || lastRow > A1.MaxRows
                        || firstColumn <= 0 || lastColumn < firstColumn || lastColumn > A1.MaxColumns) {
                        return false;
                    }
                    long columns = (long)lastColumn - firstColumn + 1;
                    long cells = ((long)lastRow - firstRow + 1) * columns;
                    return columns > owner._opt.MaxDataReaderColumns
                        || cells > owner._opt.MaxDataReaderBufferedCells
                        || cells > MaximumIndexedCells;
                } finally {
                    ArrayPool<byte>.Shared.Return(prefix, clearArray: true);
                    if (fragment != null) ArrayPool<byte>.Shared.Return(fragment, clearArray: true);
                }
            }
        }
    }
}
