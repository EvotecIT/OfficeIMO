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
                output = OpenXmlPartBufferPool.Rent(total);
                Buffer.BlockCopy(prefix, 0, output, 0, 256);
                byte[] bytes = output;
                consumer = Task.Factory.StartNew(() => consume(bytes, total, progress),
                    CancellationToken.None,
                    TaskCreationOptions.LongRunning | TaskCreationOptions.DenyChildAttach,
                    TaskScheduler.Default);
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

    /// <summary>
    /// Publishes only fully written prefixes. Row-closing tokens are availability
    /// hints, never validation: unusual or incomplete XML retains the serial parser.
    /// </summary>
    internal sealed class WorksheetInflationProgress {
        private readonly object _gate = new object();
        private int _available;
        private int _written = 256;
        private int _scanned;
        private bool _complete;

        /// <summary>Publishes a written prefix while retaining enough overlap to find split closing tokens.</summary>
        internal void PublishRows(byte[] bytes, int written) {
            // Only the producer mutates scanning state. The release publication
            // below makes the initialized prefix visible before a consumer uses it.
            int searchStart = Math.Max(_available, Math.Max(0, _scanned - 5));
            _written = written;
            _scanned = written;
            int closing = bytes.AsSpan(searchStart, written - searchStart).LastIndexOf("</row>"u8);
            if (closing < 0) return;
            lock (_gate) {
                Volatile.Write(ref _available, searchStart + closing + 6);
                Monitor.PulseAll(_gate);
            }
        }

        /// <summary>Waits for a longer initialized prefix, or a terminal producer state.</summary>
        internal int WaitForMore(int previous, CancellationToken ct) {
            ct.ThrowIfCancellationRequested();
            int available = Volatile.Read(ref _available);
            if (available > previous) return available;
            lock (_gate) {
                while (_available <= previous && !_complete) {
                    ct.ThrowIfCancellationRequested();
                    Monitor.Wait(_gate, 20);
                }
                ct.ThrowIfCancellationRequested();
                return _available;
            }
        }

        /// <summary>Waits until global XML proof can inspect the complete produced part.</summary>
        internal int WaitForCompletion(CancellationToken ct) {
            ct.ThrowIfCancellationRequested();
            if (Volatile.Read(ref _complete)) return Volatile.Read(ref _available);
            lock (_gate) {
                while (!_complete) {
                    ct.ThrowIfCancellationRequested();
                    Monitor.Wait(_gate, 20);
                }
                ct.ThrowIfCancellationRequested();
                return _available;
            }
        }

        /// <summary>Publishes the full part only after decompression and exact-length checks succeed.</summary>
        internal void Complete(int length) {
            lock (_gate) {
                Volatile.Write(ref _available, length);
                Volatile.Write(ref _complete, true);
                Monitor.PulseAll(_gate);
            }
        }

        /// <summary>Wakes a failed producer's worker using only bytes actually written, never pooled tail capacity.</summary>
        internal void Abort() {
            lock (_gate) {
                if (_complete) return;
                Volatile.Write(ref _available, _written);
                Volatile.Write(ref _complete, true);
                Monitor.PulseAll(_gate);
            }
        }
    }
}
