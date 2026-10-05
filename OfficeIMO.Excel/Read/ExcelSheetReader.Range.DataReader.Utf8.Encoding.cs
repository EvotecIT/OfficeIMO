#nullable enable

using System.Buffers;
using System.Threading;

namespace OfficeIMO.Excel {
    internal sealed partial class ExcelSheetReader {
        // The indexer already rejects NUL bytes in the first 256 bytes. Share
        // that eligibility check with worksheet prefetch so both routes decline
        // UTF-16/UTF-32 before renting a full buffer. Complete validation follows.
        internal static bool CanBufferUtf8WorksheetPrefix(byte[] prefix, int length) =>
            prefix.AsSpan(0, Math.Min(length, 256)).IndexOf((byte)0) < 0;

        private sealed partial class ExcelUtf8RangeRowSource {
            private static bool TryReadUtf8WorksheetPrefix(Stream stream, CancellationToken ct,
                out byte[] prefix, out int length) {
                prefix = ArrayPool<byte>.Shared.Rent(256);
                length = 0;
                bool transferred = false;
                try {
                    while (length < 256) {
                        ct.ThrowIfCancellationRequested();
                        int read = stream.Read(prefix, length, 256 - length);
                        if (read == 0) break;
                        length += read;
                    }
                    ct.ThrowIfCancellationRequested();
                    if (!CanBufferUtf8WorksheetPrefix(prefix, length)) return false;
                    transferred = true;
                    return true;
                } finally {
                    if (!transferred) ArrayPool<byte>.Shared.Return(prefix, clearArray: true);
                }
            }
        }
    }
}
