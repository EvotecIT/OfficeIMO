#nullable enable

using System.Buffers;

namespace OfficeIMO.Excel {
    internal sealed partial class OpenXmlPackagePartBufferReader {
        /// <summary>
        /// Starts inflating a worksheet on a background task and returns at once. The caller reads
        /// published row prefixes while inflation continues; disposal joins the producer before the
        /// pooled buffer is returned.
        /// </summary>
        internal bool TryStartStreamedRead(string partName, int maximumBytes, CancellationToken ct,
            out StreamedWorksheetPart? part) {
            part = null;
            if (!AllowsPipelining) return false;
            if (_disposed) throw new ObjectDisposedException(nameof(OpenXmlPackagePartBufferReader));
            ct.ThrowIfCancellationRequested();
            string normalized = NormalizePartName(partName);
            if (!_entries.TryGetValue(normalized, out var entry)
                || entry.Length < 1_048_576 || entry.Length > maximumBytes || entry.Length > int.MaxValue
                || entry.CompressedLength >= entry.Length) return false;
            int total = checked((int)entry.Length);
            Stream input = OpenEntry(entry, ct);
            byte[] prefix = ArrayPool<byte>.Shared.Rent(256);
            byte[]? output = null;
            try {
                ReadPartPrefix(input, prefix, 256, normalized, ct);
                if (!ExcelSheetReader.CanBufferUtf8WorksheetPrefix(prefix, 256)) return false;
                output = OpenXmlPartBufferPool.Rent(total);
                Buffer.BlockCopy(prefix, 0, output, 0, 256);
                part = new StreamedWorksheetPart(output, total, input, normalized, ct);
                output = null;
                input = null!;
                return true;
            } finally {
                input?.Dispose();
                if (output != null) OpenXmlPartBufferPool.Return(output);
                ArrayPool<byte>.Shared.Return(prefix, clearArray: true);
            }
        }

        /// <summary>A worksheet part being inflated in the background into a pooled buffer.</summary>
        internal sealed class StreamedWorksheetPart : IDisposable {
            private readonly Task _producer;
            private byte[]? _buffer;
            private Exception? _fault;

            internal StreamedWorksheetPart(byte[] buffer, int length, Stream input, string normalizedPartName,
                CancellationToken ct) {
                _buffer = buffer;
                Length = length;
                Progress = new WorksheetInflationProgress();
                _producer = Task.Run(() => {
                    try {
                        ReadPartContents(input, buffer, length, normalizedPartName, ct, 256,
                            written => Progress.PublishRows(buffer, written));
                        Progress.Complete(length);
                    } catch (Exception exception) {
                        _fault = exception;
                    } finally {
                        // A failed producer publishes only bytes it actually wrote.
                        Progress.Abort();
                        input.Dispose();
                    }
                }, CancellationToken.None);
            }

            internal byte[] Buffer => _buffer ?? throw new ObjectDisposedException(nameof(StreamedWorksheetPart));

            internal int Length { get; }

            internal WorksheetInflationProgress Progress { get; }

            /// <summary>The producer's failure, visible once the progress reports completion.</summary>
            internal Exception? Fault => Volatile.Read(ref _fault);

            public void Dispose() {
                try { _producer.GetAwaiter().GetResult(); } catch (Exception) { }
                byte[]? buffer = Interlocked.Exchange(ref _buffer, null);
                if (buffer != null) OpenXmlPartBufferPool.Return(buffer);
            }
        }
    }
}
