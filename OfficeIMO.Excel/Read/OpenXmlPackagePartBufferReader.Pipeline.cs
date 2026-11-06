#nullable enable

using System.Buffers;
using System.Threading;
using System.Threading.Tasks;

namespace OfficeIMO.Excel {
    internal sealed partial class OpenXmlPackagePartBufferReader {
        /// <summary>Allows overlap only on a worksheet archive independent of metadata cursors.</summary>
        internal bool AllowsPipelining { get; }

        /// <summary>
        /// Inflates on the calling thread while one worker indexes published row
        /// prefixes. Both phases finish before ownership leaves this method. Package
        /// decompression/length errors retain precedence over indexing errors.
        /// </summary>
        internal bool TryReadPipelined(
            string partName, int maximumBytes, CancellationToken ct,
            Func<byte[], int, WorksheetInflationProgress, bool> consume,
            out byte[]? buffer, out int length, out bool consumed) {
            buffer = null;
            length = 0;
            consumed = false;
            if (!AllowsPipelining) return false;
            if (_disposed) throw new ObjectDisposedException(nameof(OpenXmlPackagePartBufferReader));
            ct.ThrowIfCancellationRequested();
            string normalized = NormalizePartName(partName);
            if (!_entries.TryGetValue(normalized, out var entry)
                || entry.Length < 1_048_576 || entry.Length > maximumBytes || entry.Length > int.MaxValue
                || entry.CompressedLength >= entry.Length) return false;
            int total = checked((int)entry.Length);
            using Stream input = entry.Open();
            byte[] prefix = ArrayPool<byte>.Shared.Rent(256);
            byte[]? output = null;
            Task<bool>? consumer = null;
            var progress = new WorksheetInflationProgress();
            try {
                ReadPartPrefix(input, prefix, 256, normalized, ct);
                if (!ExcelSheetReader.CanBufferUtf8WorksheetPrefix(prefix, 256)) return false;
#if NET8_0_OR_GREATER
                if (TryReadManagedPipelined(entry, total, ct, consume, out buffer, out consumed)) {
                    length = total;
                    return true;
                }
#endif
                output = OpenXmlPartBufferPool.Rent(total);
                Buffer.BlockCopy(prefix, 0, output, 0, 256);
                byte[] bytes = output;
                consumer = Task.Run(() => consume(bytes, total, progress), CancellationToken.None);
                ReadPartContents(input, output, total, normalized, ct, 256,
                    written => progress.PublishRows(bytes, written));
                progress.Complete(total);
                consumed = consumer.GetAwaiter().GetResult();
                ct.ThrowIfCancellationRequested();
                buffer = output;
                length = total;
                output = null;
                return true;
            } finally {
                // A failing producer must wake and join its worker before the
                // shared byte array can be cleared or returned to its pool.
                progress.Abort();
                if (consumer != null) {
                    try { consumer.GetAwaiter().GetResult(); } catch { }
                }
                if (output != null) OpenXmlPartBufferPool.Return(output);
                ArrayPool<byte>.Shared.Return(prefix, clearArray: true);
            }
        }
    }

}
