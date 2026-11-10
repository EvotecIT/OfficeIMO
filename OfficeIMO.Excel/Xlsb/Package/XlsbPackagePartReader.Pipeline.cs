#if NET8_0_OR_GREATER
using System.Buffers;

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
            byte[] buffer = ExcelArrayPool.RentBytes(length);
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

        /// <summary>
        /// Starts inflating a worksheet on a background task and returns at once, for single-pass reading.
        /// The caller owns the part and must join <paramref name="producer"/> before disposing it.
        /// </summary>
        internal bool TryStartStreamedRead(string partName, CancellationToken cancellationToken,
            out XlsbPooledPartStream? part, out WorksheetInflationProgress? progress, out Task? producer,
            out Func<Exception?>? producerFault) {
            part = null;
            progress = null;
            producer = null;
            producerFault = null;
            cancellationToken.ThrowIfCancellationRequested();
            string normalized = NormalizePartName(partName);
            if (!_entries.TryGetValue(normalized, out var entry)
                || entry.Length < 1_048_576 || entry.Length > Math.Min(_options.MaxPartBytes, 64 * 1024 * 1024)
                || entry.CompressedLength >= entry.Length) return false;

            int length = checked((int)entry.Length);
            byte[] buffer = ExcelArrayPool.RentBytes(length);
            var owned = new XlsbPooledPartStream(buffer, length);
            var published = new WorksheetInflationProgress(initiallyWritten: 0);
            Exception? fault = null;
            producer = Task.Run(() => {
                try {
                    ReadSeekablePartContents(entry, normalized, buffer, length, cancellationToken, published.PublishBytes);
                    published.Complete(length);
                } catch (Exception exception) {
                    fault = exception;
                } finally {
                    // Validation waiting on bytes that will never arrive stops with an end-of-stream error.
                    published.Abort();
                }
            }, CancellationToken.None);
            part = owned;
            progress = published;
            producerFault = () => Volatile.Read(ref fault);
            return true;
        }
    }
}
#endif
