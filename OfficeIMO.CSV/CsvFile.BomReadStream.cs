#nullable enable
using System.Threading;
using System.Threading.Tasks;

namespace OfficeIMO.CSV;

internal static partial class CsvFile
{
    // StreamReader examines its first read for the BOM. A decompressor may return a single
    // byte even when its destination is large. Coalesce only the bounded encoding prefix.
    private sealed class CsvBomReadStream : Stream
    {
        private readonly Stream _inner;
        private readonly byte[] _prefix = new byte[4];
        private int _count, _offset;
        private bool _initialized;
        internal CsvBomReadStream(Stream inner) => _inner = inner;
        public override bool CanRead => _inner.CanRead;
        public override bool CanSeek => false;
        public override bool CanWrite => false;
        public override long Length => throw new NotSupportedException();
        public override long Position { get => throw new NotSupportedException(); set => throw new NotSupportedException(); }
        public override int Read(byte[] buffer, int offset, int count)
        {
            if (count == 0) return 0;
            if (!_initialized)
            {
                if (count >= _prefix.Length && _count == 0)
                {
                    int first = _inner.Read(buffer, offset, count);
                    if (first == 0 || first >= _prefix.Length) { _initialized = true; return first; }
                    Array.Copy(buffer, offset, _prefix, 0, first);
                    _count = first;
                }
                while (_count < _prefix.Length)
                {
                    int read = _inner.Read(_prefix, _count, _prefix.Length - _count);
                    if (read == 0) break;
                    _count += read;
                }
                _initialized = true;
            }
            if (_offset == _count) return _inner.Read(buffer, offset, count);
            int copied = Math.Min(count, _count - _offset);
            Array.Copy(_prefix, _offset, buffer, offset, copied);
            _offset += copied;
            return copied;
        }
        public override async Task<int> ReadAsync(byte[] buffer, int offset, int count, CancellationToken token)
        {
            if (count == 0) return 0;
            if (!_initialized)
            {
                if (count >= _prefix.Length && _count == 0)
                {
                    int first = await _inner.ReadAsync(buffer, offset, count, token).ConfigureAwait(false);
                    if (first == 0 || first >= _prefix.Length) { _initialized = true; return first; }
                    Array.Copy(buffer, offset, _prefix, 0, first);
                    _count = first;
                }
                while (_count < _prefix.Length)
                {
                    int read = await _inner.ReadAsync(_prefix, _count, _prefix.Length - _count, token).ConfigureAwait(false);
                    if (read == 0) break;
                    _count += read;
                }
                _initialized = true;
            }
            if (_offset == _count) return await _inner.ReadAsync(buffer, offset, count, token).ConfigureAwait(false);
            int copied = Math.Min(count, _count - _offset);
            Array.Copy(_prefix, _offset, buffer, offset, copied);
            _offset += copied;
            return copied;
        }
#if NET8_0_OR_GREATER
        public override async ValueTask<int> ReadAsync(Memory<byte> buffer, CancellationToken token = default)
        {
            if (buffer.Length == 0) return 0;
            if (!_initialized)
            {
                if (buffer.Length >= _prefix.Length && _count == 0)
                {
                    int first = await _inner.ReadAsync(buffer, token).ConfigureAwait(false);
                    if (first == 0 || first >= _prefix.Length) { _initialized = true; return first; }
                    buffer.Slice(0, first).CopyTo(_prefix);
                    _count = first;
                }
                while (_count < _prefix.Length)
                {
                    int read = await _inner.ReadAsync(_prefix.AsMemory(_count), token).ConfigureAwait(false);
                    if (read == 0) break;
                    _count += read;
                }
                _initialized = true;
            }
            if (_offset == _count) return await _inner.ReadAsync(buffer, token).ConfigureAwait(false);
            int copied = Math.Min(buffer.Length, _count - _offset);
            _prefix.AsMemory(_offset, copied).CopyTo(buffer);
            _offset += copied;
            return copied;
        }
#endif
        public override void Flush() => _inner.Flush();
        public override long Seek(long offset, SeekOrigin origin) => throw new NotSupportedException();
        public override void SetLength(long value) => throw new NotSupportedException();
        public override void Write(byte[] buffer, int offset, int count) => throw new NotSupportedException();
        protected override void Dispose(bool disposing)
        {
            if (disposing) _inner.Dispose();
            base.Dispose(disposing);
        }
    }
}
