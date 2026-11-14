using System.Buffers;
using System.Runtime.ExceptionServices;

namespace OfficeIMO.Excel {
    /// <summary>
    /// Overlaps a private worksheet entry's compression with caller-side serialization.
    /// Four owned buffers bound the handoff. Whichever thread holds the writer role writes queued chunks
    /// in order: a thread-pool drain while the caller keeps serializing, or the caller itself when it would
    /// otherwise wait for a free buffer or for the entry to drain. A busy host can then delay the background
    /// drain without stalling the caller. Completion publishes all bytes; disposal after a failed producer
    /// aborts and waits for the active write before clearing pooled storage.
    /// A package reuses one pipeline for its entries: <see cref="Drain"/> ends an entry and
    /// <see cref="Retarget"/> starts the next, so the buffers outlive a sheet.
    /// </summary>
    internal sealed class WorksheetCompressionPipeline : Stream {
        // Large chunks keep hand-offs rare and give the background drain a long run per wake-up.
        internal const int ChunkSize = 256 * 1024;
        private const int BufferCount = 4;
        private static readonly WaitCallback BackgroundDrain = state => ((WorksheetCompressionPipeline)state!).DrainInBackground();
        private readonly object _gate = new();
        private readonly byte[][] _owned = new byte[BufferCount][];
        private readonly Stack<byte[]> _free = new(BufferCount);
        private readonly Queue<Chunk> _ready = new(BufferCount);
        private Stream _destination;
        private byte[]? _current;
        private int _count;
        // Guarded by _gate: one thread at a time writes to the destination; a queued or running pool drain.
        private bool _writing;
        private bool _backgroundActive;
        private bool _aborted;
        private ExceptionDispatchInfo? _failure;
        private bool _complete;
        private bool _disposed;
        private readonly struct Chunk(byte[] buffer, int count) {
            internal readonly byte[] Buffer = buffer;
            internal readonly int Count = count;
        }
        internal WorksheetCompressionPipeline(Stream destination) {
            _destination = destination;
            try {
                for (int i = 0; i < _owned.Length; i++) {
                    _owned[i] = ArrayPool<byte>.Shared.Rent(ChunkSize);
                    _free.Push(_owned[i]);
                }
            } catch {
                foreach (byte[] buffer in _owned) if (buffer != null) ArrayPool<byte>.Shared.Return(buffer, clearArray: true);
                throw;
            }
        }

        private void DrainInBackground() {
            lock (_gate) {
                while (_ready.Count > 0 && !_writing && !_aborted && _failure == null) {
                    WriteNext(rethrow: false);
                }
                _backgroundActive = false;
            }
        }

        /// <summary>Writes the oldest queued chunk outside the lock; the caller holds <see cref="_gate"/> and sees no active writer.</summary>
        private void WriteNext(bool rethrow) {
            Chunk chunk = _ready.Dequeue();
            Stream destination = _destination;
            _writing = true;
            Exception? error = null;
            Monitor.Exit(_gate);
            try {
                destination.Write(chunk.Buffer, 0, chunk.Count);
            } catch (Exception failure) {
                error = failure;
            } finally {
                Monitor.Enter(_gate);
                _writing = false;
                _free.Push(chunk.Buffer);
                if (error != null && _failure == null) _failure = ExceptionDispatchInfo.Capture(error);
                Monitor.PulseAll(_gate);
            }
            if (rethrow) ThrowFailure();
        }

        private void ScheduleBackgroundDrain() {
            if (_backgroundActive || _writing || _ready.Count == 0) return;
            _backgroundActive = true;
            try {
                ThreadPool.UnsafeQueueUserWorkItem(BackgroundDrain, this);
            } catch {
                // The caller drains everything itself when the pool rejects work.
                _backgroundActive = false;
            }
        }

        private void ThrowFailure() { _failure?.Throw(); }

        private void Publish() {
            if (_count == 0) return;
            lock (_gate) {
                ThrowFailure();
                _ready.Enqueue(new Chunk(_current!, _count));
                _current = null;
                _count = 0;
                ScheduleBackgroundDrain();
            }
        }

        /// <summary>Takes a free buffer, writing the oldest chunk on this thread when every buffer is queued.</summary>
        private void AcquireBuffer() {
            lock (_gate) {
                while (true) {
                    ThrowFailure();
                    if (_free.Count > 0) {
                        _current = _free.Pop();
                        return;
                    }
                    if (!_writing) {
                        WriteNext(rethrow: true);
                        ScheduleBackgroundDrain();
                    } else {
                        Monitor.Wait(_gate);
                    }
                }
            }
        }

        public override void Write(byte[] buffer, int offset, int count) {
            if (_disposed || _complete) throw new ObjectDisposedException(nameof(WorksheetCompressionPipeline));
            if (buffer == null) throw new ArgumentNullException(nameof(buffer));
            if (offset < 0 || count < 0 || offset > buffer.Length - count) throw new ArgumentOutOfRangeException();
            ThrowFailure();
            while (count > 0) {
                if (_current == null) AcquireBuffer();
                int copied = Math.Min(count, ChunkSize - _count);
                Buffer.BlockCopy(buffer, offset, _current!, _count, copied);
                _count += copied; offset += copied; count -= copied;
                if (_count == ChunkSize) Publish();
            }
        }

        /// <summary>Writes every accepted byte to the current destination, helping the background drain finish.</summary>
        private void WriteAll() {
            Publish();
            lock (_gate) {
                while (true) {
                    ThrowFailure();
                    if (_writing) Monitor.Wait(_gate);
                    else if (_ready.Count > 0) WriteNext(rethrow: true);
                    else return;
                }
            }
        }

        /// <summary>
        /// Waits until every accepted byte reached the current destination, keeping the buffers for the next entry.
        /// Surfaces the original write failure.
        /// </summary>
        internal void Drain() {
            if (_disposed || _complete) throw new ObjectDisposedException(nameof(WorksheetCompressionPipeline));
            WriteAll();
        }

        /// <summary>Sends later bytes to <paramref name="destination"/>; valid only after <see cref="Drain"/>.</summary>
        internal void Retarget(Stream destination) {
            if (_disposed || _complete) throw new ObjectDisposedException(nameof(WorksheetCompressionPipeline));
            lock (_gate) {
                ThrowFailure();
                _destination = destination;
            }
        }

        /// <summary>Writes all accepted bytes and surfaces the original write failure.</summary>
        internal void Complete() {
            if (_disposed) throw new ObjectDisposedException(nameof(WorksheetCompressionPipeline));
            if (_complete) return;
            WriteAll();
            _complete = true;
        }

        protected override void Dispose(bool disposing) {
            if (_disposed) return;
            _disposed = true;
            if (disposing) {
                // Success requires Complete. Abort waits for the active write before any buffer can be returned;
                // a pool drain that runs later finds nothing queued.
                lock (_gate) {
                    _aborted = true;
                    _ready.Clear();
                    while (_writing) Monitor.Wait(_gate);
                }
                foreach (byte[] buffer in _owned) ArrayPool<byte>.Shared.Return(buffer, clearArray: true);
            }
            base.Dispose(disposing);
        }
        public override void Flush() { ThrowFailure(); }
        public override bool CanRead => false;
        public override bool CanSeek => false;
        public override bool CanWrite => !_disposed && !_complete;
        public override long Length => throw new NotSupportedException();
        public override long Position { get => throw new NotSupportedException(); set => throw new NotSupportedException(); }
        public override int Read(byte[] buffer, int offset, int count) => throw new NotSupportedException();
        public override long Seek(long offset, SeekOrigin origin) => throw new NotSupportedException();
        public override void SetLength(long value) => throw new NotSupportedException();
    }
}
