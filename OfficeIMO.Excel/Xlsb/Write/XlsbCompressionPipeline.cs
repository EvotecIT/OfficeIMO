using System.Buffers;
using System.Collections.Concurrent;
using System.Runtime.ExceptionServices;
using System.Threading;
using System.Threading.Tasks;

namespace OfficeIMO.Excel.Xlsb.Write {
    /// <summary>
    /// Overlaps one private worksheet entry's compression with caller-side serialization.
    /// Two owned buffers bound the handoff. Completion publishes all bytes; disposal
    /// after a failed producer aborts and joins before clearing pooled storage.
    /// </summary>
    internal sealed class XlsbCompressionPipeline : Stream {
        private const int ChunkSize = 64 * 1024;
        private readonly Stream _destination;
        private readonly BlockingCollection<byte[]> _available = new(2);
        private readonly BlockingCollection<Chunk> _ready = new(1);
        private readonly CancellationTokenSource _stop = new();
        private readonly byte[][] _owned = new byte[2][];
        private readonly Task _worker;
        private byte[]? _current;
        private int _count;
        private ExceptionDispatchInfo? _failure;
        private bool _complete;
        private bool _disposed;
        private readonly struct Chunk(byte[] buffer, int count) {
            internal readonly byte[] Buffer = buffer;
            internal readonly int Count = count;
        }
        internal XlsbCompressionPipeline(Stream destination) {
            _destination = destination;
            try {
                for (int i = 0; i < _owned.Length; i++) {
                    _owned[i] = ArrayPool<byte>.Shared.Rent(ChunkSize);
                    _available.Add(_owned[i]);
                }
                _current = _available.Take();
                _worker = Task.Run(Consume);
            } catch {
                foreach (byte[] buffer in _owned) if (buffer != null) ArrayPool<byte>.Shared.Return(buffer, clearArray: true);
                _available.Dispose();
                _ready.Dispose();
                _stop.Dispose();
                throw;
            }
        }
        private void Consume() {
            try {
                foreach (Chunk item in _ready.GetConsumingEnumerable(_stop.Token)) {
                    _destination.Write(item.Buffer, 0, item.Count);
                    _available.Add(item.Buffer, _stop.Token);
                }
            } catch (OperationCanceledException) when (_stop.IsCancellationRequested) {
            } catch (Exception error) {
                _failure = ExceptionDispatchInfo.Capture(error);
                _stop.Cancel();
            }
        }
        private void ThrowFailure() { _failure?.Throw(); }
        private void Publish() {
            if (_count == 0) return;
            try {
                _ready.Add(new Chunk(_current!, _count), _stop.Token);
                _current = null;
                _count = 0;
                _current = _available.Take(_stop.Token);
            } catch (OperationCanceledException) { ThrowFailure(); throw; }
        }
        public override void Write(byte[] buffer, int offset, int count) {
            if (_disposed || _complete) throw new ObjectDisposedException(nameof(XlsbCompressionPipeline));
            if (buffer == null) throw new ArgumentNullException(nameof(buffer));
            if (offset < 0 || count < 0 || offset > buffer.Length - count) throw new ArgumentOutOfRangeException();
            ThrowFailure();
            while (count > 0) {
                int copied = Math.Min(count, ChunkSize - _count);
                Buffer.BlockCopy(buffer, offset, _current!, _count, copied);
                _count += copied; offset += copied; count -= copied;
                if (_count == ChunkSize) Publish();
            }
        }
        /// <summary>Drains all accepted bytes and surfaces the original worker failure.</summary>
        internal void Complete() {
            if (_disposed) throw new ObjectDisposedException(nameof(XlsbCompressionPipeline));
            if (_complete) return;
            Publish();
            _ready.CompleteAdding();
            _worker.GetAwaiter().GetResult();
            ThrowFailure();
            _complete = true;
        }
        protected override void Dispose(bool disposing) {
            if (_disposed) return;
            _disposed = true;
            if (disposing) {
                // Success requires Complete. Abort joins before any buffer can be returned.
                _stop.Cancel();
                _ready.CompleteAdding();
                _worker.GetAwaiter().GetResult();
                foreach (byte[] buffer in _owned) ArrayPool<byte>.Shared.Return(buffer, clearArray: true);
                _available.Dispose();
                _ready.Dispose();
                _stop.Dispose();
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
