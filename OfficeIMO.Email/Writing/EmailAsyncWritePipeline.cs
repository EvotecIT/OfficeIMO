namespace OfficeIMO.Email;

/// <summary>Bridges synchronous deterministic format producers to genuinely asynchronous bounded destination I/O.</summary>
internal static class EmailAsyncWritePipeline {
    private const int ChunkSize = 64 * 1024;
    private const int BufferedChunks = 4;

    internal static async Task<long> RunAsync(Stream destination, long maximumLength,
        Action<Stream> producer, CancellationToken cancellationToken) {
        if (destination == null) throw new ArgumentNullException(nameof(destination));
        if (producer == null) throw new ArgumentNullException(nameof(producer));

        using (var linkedCancellation = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken))
        using (var queue = new BoundedAsyncChunkQueue(BufferedChunks)) {
            Task<long> producerTask = Task.Run(() => {
                try {
                    using (var producerStream = new ChunkProducerStream(queue, ChunkSize, linkedCancellation.Token))
                    using (var bounded = new EmailBoundedWriteStream(producerStream, maximumLength)) {
                        producer(bounded);
                        producerStream.Complete();
                        return bounded.BytesWritten;
                    }
                } catch (Exception exception) {
                    queue.Complete(exception);
                    throw;
                }
            }, linkedCancellation.Token);

            try {
                while (true) {
                    byte[]? chunk = await queue.DequeueAsync(cancellationToken).ConfigureAwait(false);
                    if (chunk == null) break;
                    await destination.WriteAsync(chunk, 0, chunk.Length, cancellationToken).ConfigureAwait(false);
                }
                return await producerTask.ConfigureAwait(false);
            } catch {
                linkedCancellation.Cancel();
                try {
                    await producerTask.ConfigureAwait(false);
                } catch {
                    // Preserve the consumer failure or cancellation that stopped the pipeline.
                }
                throw;
            }
        }
    }

    private sealed class ChunkProducerStream : Stream {
        private readonly BoundedAsyncChunkQueue _queue;
        private readonly int _chunkSize;
        private readonly CancellationToken _cancellationToken;
        private byte[] _buffer;
        private int _count;
        private bool _completed;

        internal ChunkProducerStream(BoundedAsyncChunkQueue queue, int chunkSize,
            CancellationToken cancellationToken) {
            _queue = queue;
            _chunkSize = chunkSize;
            _cancellationToken = cancellationToken;
            _buffer = new byte[chunkSize];
        }

        public override bool CanRead => false;
        public override bool CanSeek => false;
        public override bool CanWrite => true;
        public override long Length => throw new NotSupportedException();
        public override long Position { get => throw new NotSupportedException(); set => throw new NotSupportedException(); }
        public override void Flush() => FlushChunk();
        public override int Read(byte[] buffer, int offset, int count) => throw new NotSupportedException();
        public override long Seek(long offset, SeekOrigin origin) => throw new NotSupportedException();
        public override void SetLength(long value) => throw new NotSupportedException();

        public override void Write(byte[] buffer, int offset, int count) {
            if (buffer == null) throw new ArgumentNullException(nameof(buffer));
            if (offset < 0 || count < 0 || offset > buffer.Length - count) throw new ArgumentOutOfRangeException(nameof(offset));
            while (count > 0) {
                _cancellationToken.ThrowIfCancellationRequested();
                int copy = Math.Min(count, _buffer.Length - _count);
                Buffer.BlockCopy(buffer, offset, _buffer, _count, copy);
                offset += copy;
                count -= copy;
                _count += copy;
                if (_count == _buffer.Length) FlushChunk();
            }
        }

        internal void Complete() {
            if (_completed) return;
            FlushChunk();
            _completed = true;
            _queue.Complete();
        }

        protected override void Dispose(bool disposing) {
            if (disposing && !_completed) Complete();
            base.Dispose(disposing);
        }

        private void FlushChunk() {
            if (_count == 0) return;
            byte[] chunk;
            if (_count == _chunkSize) {
                chunk = _buffer;
            } else {
                chunk = new byte[_count];
                Buffer.BlockCopy(_buffer, 0, chunk, 0, _count);
            }
            _queue.Enqueue(chunk, _cancellationToken);
            _buffer = new byte[_chunkSize];
            _count = 0;
        }
    }

    private sealed class BoundedAsyncChunkQueue : IDisposable {
        private readonly Queue<byte[]> _queue = new Queue<byte[]>();
        private readonly SemaphoreSlim _items = new SemaphoreSlim(0);
        private readonly SemaphoreSlim _slots;
        private bool _completed;
        private Exception? _error;

        internal BoundedAsyncChunkQueue(int capacity) {
            _slots = new SemaphoreSlim(capacity, capacity);
        }

        internal void Enqueue(byte[] chunk, CancellationToken cancellationToken) {
            _slots.Wait(cancellationToken);
            lock (_queue) {
                if (_completed) {
                    _slots.Release();
                    throw new InvalidOperationException("The asynchronous write queue is complete.");
                }
                _queue.Enqueue(chunk);
            }
            _items.Release();
        }

        internal async Task<byte[]?> DequeueAsync(CancellationToken cancellationToken) {
            await _items.WaitAsync(cancellationToken).ConfigureAwait(false);
            lock (_queue) {
                if (_queue.Count > 0) {
                    byte[] chunk = _queue.Dequeue();
                    _slots.Release();
                    return chunk;
                }
                if (_error != null) throw new InvalidDataException("Email artifact production failed.", _error);
                if (_completed) return null;
            }
            throw new InvalidOperationException("The asynchronous write queue was signaled without data.");
        }

        internal void Complete(Exception? error = null) {
            lock (_queue) {
                if (_completed) return;
                _completed = true;
                _error = error;
            }
            _items.Release();
        }

        public void Dispose() {
            _items.Dispose();
            _slots.Dispose();
        }
    }
}
