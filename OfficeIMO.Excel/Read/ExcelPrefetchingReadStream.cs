#nullable enable

using System.Buffers;
using System.Runtime.ExceptionServices;

namespace OfficeIMO.Excel {
    /// <summary>
    /// Reads a forward-only source on a background task into a fixed set of chunks, so decompression overlaps the
    /// caller's parsing while at most <see cref="ChunkCount"/> chunks are held. Source errors surface, in order,
    /// after the bytes read before them.
    /// </summary>
    internal sealed class ExcelPrefetchingReadStream : Stream {
        private const int ChunkSize = 1024 * 1024;
        private const int ChunkCount = 4;
        private readonly Stream _source;
        private readonly object _gate = new object();
        private readonly Queue<(byte[] Buffer, int Length)> _filled = new Queue<(byte[] Buffer, int Length)>(ChunkCount);
        private readonly Stack<byte[]> _free = new Stack<byte[]>(ChunkCount);
        private readonly Task _producer;
        private readonly CancellationToken _ct;
        private byte[]? _current;
        private int _currentLength;
        private int _currentPosition;
        private bool _completed;
        private bool _stopping;
        private ExceptionDispatchInfo? _fault;
        private long _position;

        private ExcelPrefetchingReadStream(Stream source, CancellationToken ct) {
            _source = source;
            _ct = ct;
            for (int index = 0; index < ChunkCount; index++) _free.Push(ArrayPool<byte>.Shared.Rent(ChunkSize));
            _producer = Task.Run(Produce, CancellationToken.None);
        }

        /// <summary>Prefetches when the process-wide read budget has a free worker; otherwise returns the source.</summary>
        internal static Stream Wrap(Stream source, CancellationToken ct) {
            if (ExcelParallelReadBudget.Reserve(1) == 0) return source;
            try {
                return new ExcelPrefetchingReadStream(source, ct);
            } catch {
                ExcelParallelReadBudget.Release(1);
                throw;
            }
        }

        private void Produce() {
            try {
                while (true) {
                    byte[] chunk;
                    lock (_gate) {
                        while (_free.Count == 0 && !_stopping) Monitor.Wait(_gate);
                        if (_stopping) return;
                        chunk = _free.Pop();
                    }
                    int length = 0;
                    try {
                        while (length < ChunkSize) {
                            int read = _source.Read(chunk, length, ChunkSize - length);
                            if (read == 0) break;
                            length += read;
                        }
                    } catch (Exception exception) {
                        lock (_gate) {
                            _free.Push(chunk);
                            _fault = ExceptionDispatchInfo.Capture(exception);
                            _completed = true;
                            Monitor.PulseAll(_gate);
                        }
                        return;
                    }
                    lock (_gate) {
                        if (length > 0) _filled.Enqueue((chunk, length));
                        else _free.Push(chunk);
                        if (length < ChunkSize) _completed = true;
                        Monitor.PulseAll(_gate);
                        if (_completed) return;
                    }
                }
            } finally {
                ExcelParallelReadBudget.Release(1);
            }
        }

        public override int Read(byte[] buffer, int offset, int count) {
            if (count == 0) return 0;
            if (_current == null || _currentPosition == _currentLength) {
                lock (_gate) {
                    if (_current != null) {
                        _free.Push(_current);
                        _current = null;
                        Monitor.PulseAll(_gate);
                    }
                    while (_filled.Count == 0 && !_completed) {
                        _ct.ThrowIfCancellationRequested();
                        Monitor.Wait(_gate, 20);
                    }
                    if (_filled.Count == 0) {
                        _fault?.Throw();
                        return 0;
                    }
                    (_current, _currentLength) = _filled.Dequeue();
                    _currentPosition = 0;
                }
            }
            int copied = Math.Min(count, _currentLength - _currentPosition);
            Buffer.BlockCopy(_current!, _currentPosition, buffer, offset, copied);
            _currentPosition += copied;
            _position += copied;
            return copied;
        }

        public override bool CanRead => true;
        public override bool CanSeek => false;
        public override bool CanWrite => false;
        public override long Length => _source.Length;
        public override long Position {
            get => _position;
            set => throw new NotSupportedException();
        }

        public override void Flush() { }
        public override long Seek(long offset, SeekOrigin origin) => throw new NotSupportedException();
        public override void SetLength(long value) => throw new NotSupportedException();
        public override void Write(byte[] buffer, int offset, int count) => throw new NotSupportedException();

        protected override void Dispose(bool disposing) {
            if (disposing) {
                lock (_gate) {
                    _stopping = true;
                    Monitor.PulseAll(_gate);
                }
                try { _producer.GetAwaiter().GetResult(); } catch { }
                // Chunks hold worksheet text, which may be customer data.
                if (_current != null) ArrayPool<byte>.Shared.Return(_current, clearArray: true);
                _current = null;
                while (_filled.Count > 0) ArrayPool<byte>.Shared.Return(_filled.Dequeue().Buffer, clearArray: true);
                while (_free.Count > 0) ArrayPool<byte>.Shared.Return(_free.Pop(), clearArray: true);
                _source.Dispose();
            }
            base.Dispose(disposing);
        }
    }
}
