using System.Buffers;
using DocumentFormat.OpenXml;

namespace OfficeIMO.Excel {
    /// <summary>A reusable write target for one package part at a time, so comparing parts on save leaves no garbage.</summary>
    internal sealed class ExcelPartBuffer : Stream {
        private byte[] _buffer = ExcelArrayPool.RentBytes(64 * 1024);
        private int _length;

        internal byte[] Buffer => _buffer;
        internal int ContentLength => _length;

        internal void Clear() => _length = 0;

        /// <summary>Drops everything after the first <paramref name="length"/> bytes.</summary>
        internal void Truncate(int length) {
            if (length < 0 || length > _length) throw new ArgumentOutOfRangeException(nameof(length));
            _length = length;
        }

        internal void WriteAscii(string text) {
            if (text.Length > _buffer.Length - _length) Grow(checked(_length + text.Length));
            for (int index = 0; index < text.Length; index++) _buffer[_length + index] = (byte)text[index];
            _length += text.Length;
        }

        internal void ReadFrom(Stream source) {
            // A part stream that knows its length gets one array of that size instead of a doubling series,
            // which for a large worksheet would be a run of large-object allocations.
            if (source.CanSeek) {
                long remaining = source.Length - source.Position;
                if (remaining > 0 && remaining < int.MaxValue - 1 - _length) EnsureCapacity((int)remaining + _length + 1);
            }
            while (true) {
                if (_length == _buffer.Length) Grow(_length + 1);
                int read = source.Read(_buffer, _length, _buffer.Length - _length);
                if (read == 0) return;
                _length += read;
            }
        }

        public override void Write(byte[] buffer, int offset, int count) {
            if (count > _buffer.Length - _length) Grow(checked(_length + count));
            System.Buffer.BlockCopy(buffer, offset, _buffer, _length, count);
            _length += count;
        }

        /// <summary>Makes room for <paramref name="capacity"/> bytes in one step.</summary>
        internal void EnsureCapacity(int capacity) {
            if (capacity <= _buffer.Length) return;
            byte[] larger = ExcelArrayPool.RentBytes(capacity);
            System.Buffer.BlockCopy(_buffer, 0, larger, 0, _length);
            if (_buffer.Length != 0) ExcelArrayPool.Return(_buffer);
            _buffer = larger;
        }

        private void Grow(int required) {
            byte[] larger = ExcelArrayPool.RentBytes(Math.Max(required, _buffer.Length * 2));
            System.Buffer.BlockCopy(_buffer, 0, larger, 0, _length);
            ExcelArrayPool.Return(_buffer);
            _buffer = larger;
        }

        public override bool CanRead => false;
        public override bool CanSeek => false;
        public override bool CanWrite => true;
        public override long Length => _length;
        public override long Position {
            get => _length;
            set => throw new NotSupportedException();
        }

        public override void Flush() { }
        public override int Read(byte[] buffer, int offset, int count) => throw new NotSupportedException();
        public override long Seek(long offset, SeekOrigin origin) => throw new NotSupportedException();
        public override void SetLength(long value) => throw new NotSupportedException();

        /// <summary>Returns the storage. <see cref="OpenXmlPartRootElement.Save(Stream)"/> closes its stream, so disposal cannot be the release.</summary>
        internal void Release() {
            if (_buffer.Length == 0) return;
            ExcelArrayPool.Return(_buffer);
            _buffer = Array.Empty<byte>();
        }
    }
}
