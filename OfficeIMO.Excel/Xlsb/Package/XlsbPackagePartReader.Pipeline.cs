#if NET8_0_OR_GREATER
using System.Buffers;
using System.Threading;
using System.Threading.Tasks;

namespace OfficeIMO.Excel.Xlsb.Package {
    internal sealed partial class XlsbPackagePartReader {
        /// <summary>Overlaps bounded worksheet inflation with complete validation; ownership transfers only after both finish.</summary>
        internal bool TryReadPipelined<T>(string partName, CancellationToken cancellationToken,
            Func<XlsbPooledPartStream, WorksheetInflationProgress, T> consume,
            out XlsbPooledPartStream? part, out T result) {
            part = null;
            result = default!;
            cancellationToken.ThrowIfCancellationRequested();
            string normalized = NormalizePartName(partName);
            if (!_entries.TryGetValue(normalized, out var entry)
                || entry.Length < 1_048_576 || entry.Length > Math.Min(_options.MaxPartBytes, 64 * 1024 * 1024)
                || entry.CompressedLength >= entry.Length) return false;

            int length = checked((int)entry.Length);
            byte[] buffer = ArrayPool<byte>.Shared.Rent(length);
            var owned = new XlsbPooledPartStream(buffer, length);
            var progress = new WorksheetInflationProgress(initiallyWritten: 0);
            Task<T>? consumer = null;
            bool transferred = false;
            try {
                consumer = Task.Run(() => consume(owned, progress), CancellationToken.None);
                ReadSeekablePartContents(entry, normalized, buffer, length, cancellationToken, progress.PublishBytes);
                progress.Complete(length);
                result = consumer.GetAwaiter().GetResult();
                cancellationToken.ThrowIfCancellationRequested();
                part = owned;
                transferred = true;
                return true;
            } finally {
                // Producer errors keep precedence. A consumer cannot retain the
                // array after this join, including malformed input and cancellation.
                progress.Abort();
                if (consumer != null) {
                    try { consumer.GetAwaiter().GetResult(); } catch { }
                }
                if (!transferred) owned.Dispose();
            }
        }
    }
}
#endif
