namespace OfficeIMO.Core.Internal;

/// <summary>
/// Hands an XML reader a package part that is smaller than its default buffer as an exact-size, seekable copy. An
/// XML reader over a compressed part stream cannot see its length and allocates its full default byte and character
/// buffers for every part; over the copy it sizes them to the part. Relationship and content-type parts are usually a
/// few hundred bytes, and opening a package reads several of them.
/// </summary>
internal static class OfficeSmallXmlPartInput {
    /// <summary>The XML reader's default buffer size; parts at least this long gain nothing from a copy.</summary>
    internal const int MaximumCopiedBytes = 4096;

    /// <summary>
    /// Returns a stream with exactly the bytes of <paramref name="source"/>. <paramref name="declaredLength"/> only
    /// selects the copy: a part that turns out longer than declared is read in full, as it would be without the copy.
    /// </summary>
    internal static Stream Prepare(Stream source, long declaredLength) {
        ArgumentNullException.ThrowIfNull(source);
        if (declaredLength <= 0 || declaredLength >= MaximumCopiedBytes) return source;

        // One byte past the declared length tells a complete part from a longer one.
        var bytes = new byte[declaredLength + 1];
        int total = 0;
        while (total < bytes.Length) {
            int read = source.Read(bytes, total, bytes.Length - total);
            if (read == 0) break;
            total += read;
        }
        return total <= declaredLength
            ? new MemoryStream(bytes, 0, total, writable: false)
            : new PrefixedReadStream(bytes, total, source);
    }

    /// <summary>Reads the bytes already taken from a part, then the rest of it.</summary>
    private sealed class PrefixedReadStream : Stream {
        private readonly byte[] _prefix;
        private readonly int _prefixLength;
        private readonly Stream _rest;
        private int _prefixOffset;

        internal PrefixedReadStream(byte[] prefix, int prefixLength, Stream rest) {
            _prefix = prefix;
            _prefixLength = prefixLength;
            _rest = rest;
        }

        public override bool CanRead => true;
        public override bool CanSeek => false;
        public override bool CanWrite => false;
        public override long Length => throw new NotSupportedException();
        public override long Position { get => throw new NotSupportedException(); set => throw new NotSupportedException(); }

        public override int Read(byte[] buffer, int offset, int count) {
            ArgumentNullException.ThrowIfNull(buffer);
            if (offset < 0 || count < 0 || offset > buffer.Length - count) throw new ArgumentOutOfRangeException(nameof(offset));
            if (_prefixOffset < _prefixLength) {
                int take = Math.Min(count, _prefixLength - _prefixOffset);
                Buffer.BlockCopy(_prefix, _prefixOffset, buffer, offset, take);
                _prefixOffset += take;
                return take;
            }
            return _rest.Read(buffer, offset, count);
        }

        public override void Flush() { }
        public override long Seek(long offset, SeekOrigin origin) => throw new NotSupportedException();
        public override void SetLength(long value) => throw new NotSupportedException();
        public override void Write(byte[] buffer, int offset, int count) => throw new NotSupportedException();
    }
}
