#nullable enable

namespace OfficeIMO.Excel {
    /// <summary>
    /// Publishes initialized worksheet bytes and terminal producer state. XML readers
    /// use row-closing hints; binary readers retain their complete framing validation.
    /// </summary>
    internal sealed class WorksheetInflationProgress {
        private readonly object _gate = new object();
        private int _available;
        private int _written;
        private int _scanned;
        private bool _complete;

        internal WorksheetInflationProgress(int initiallyWritten = 256) => _written = initiallyWritten;

        /// <summary>Publishes initialized binary bytes; framing and payload validation stay with the consumer.</summary>
        internal void PublishBytes(int written) {
            _written = written;
            lock (_gate) {
                Volatile.Write(ref _available, written);
                Monitor.PulseAll(_gate);
            }
        }

        /// <summary>Waits for a complete header or payload without exposing an unwritten pooled tail.</summary>
        internal int WaitForAvailable(int required, CancellationToken ct) {
            int available = Volatile.Read(ref _available);
            while (available < required) {
                int next = WaitForMore(available, ct);
                if (next <= available) throw new EndOfStreamException("The worksheet ended before the requested bytes were produced.");
                available = next;
            }
            return available;
        }

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
