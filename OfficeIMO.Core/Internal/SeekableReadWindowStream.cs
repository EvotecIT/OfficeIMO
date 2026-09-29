using System;
using System.IO;

namespace OfficeIMO.Core.Internal {
    /// <summary>Exposes a bounded, zero-based view of a readable seekable stream without owning it.</summary>
    internal sealed class SeekableReadWindowStream : Stream {
        private readonly Stream _source;
        private readonly long _offset;
        private readonly long _length;
        private long _position;

        internal SeekableReadWindowStream(Stream source, long offset, long length) {
            if (source == null) throw new ArgumentNullException(nameof(source));
            if (!source.CanRead || !source.CanSeek) throw new ArgumentException(
                "The source must be readable and seekable.", nameof(source));
            if (offset < 0 || length < 0 || offset > source.Length || length > source.Length - offset) {
                throw new ArgumentOutOfRangeException(nameof(length));
            }
            _source = source;
            _offset = offset;
            _length = length;
        }

        public override bool CanRead => true;
        public override bool CanSeek => true;
        public override bool CanWrite => false;
        public override long Length => _length;
        public override long Position {
            get => _position;
            set => Seek(value, SeekOrigin.Begin);
        }

        public override int Read(byte[] buffer, int offset, int count) {
            if (buffer == null) throw new ArgumentNullException(nameof(buffer));
            if (offset < 0 || count < 0 || offset > buffer.Length - count) {
                throw new ArgumentOutOfRangeException(nameof(offset));
            }
            long remaining = _length - _position;
            if (remaining <= 0) return 0;
            int requested = (int)Math.Min(count, remaining);
            _source.Position = checked(_offset + _position);
            int read = _source.Read(buffer, offset, requested);
            _position += read;
            return read;
        }

        public override long Seek(long offset, SeekOrigin origin) {
            long target = origin switch {
                SeekOrigin.Begin => offset,
                SeekOrigin.Current => checked(_position + offset),
                SeekOrigin.End => checked(_length + offset),
                _ => throw new ArgumentOutOfRangeException(nameof(origin))
            };
            if (target < 0 || target > _length) throw new IOException(
                "Attempted to seek outside the stream window.");
            _position = target;
            return target;
        }

        public override void Flush() { }
        public override void SetLength(long value) => throw new NotSupportedException();
        public override void Write(byte[] buffer, int offset, int count) => throw new NotSupportedException();
    }
}
