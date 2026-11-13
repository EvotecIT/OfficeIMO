#if NET8_0_OR_GREATER
namespace OfficeIMO.SharedSource.IO {
    internal sealed partial class PooledUtf8TextWriter {
        /// <summary>Copies validated UTF-8 bytes after completing preceding character writes.</summary>
        internal void WriteUtf8(ReadOnlySpan<byte> value) {
            char[] characters = GetCharacters();
            byte[] bytes = _bytes!;
            // A complete UTF-8 value cannot continue a pending UTF-16 surrogate.
            if ((_characterCount != 0 || _encoderMayHavePendingSurrogate)
                && !TryEncodeBufferedAscii(characters, bytes)) {
                EncodeBufferedCharacters(flushEncoder: true);
            }
            while (!value.IsEmpty) {
                if (_byteCount == bytes.Length) FlushBytes();
                int count = Math.Min(value.Length, bytes.Length - _byteCount);
                value.Slice(0, count).CopyTo(bytes.AsSpan(_byteCount));
                _byteCount += count;
                value = value.Slice(count);
            }
        }

        /// <summary>
        /// Completes preceding character writes and returns the byte buffer with at least <paramref name="sizeHint"/>
        /// free bytes after <paramref name="offset"/> when the buffer is that large. Bytes written there are output
        /// only once <see cref="CommitUtf8"/> counts them, and no other write may come in between.
        /// </summary>
        internal byte[] ReserveUtf8(int sizeHint, out int offset) {
            char[] characters = GetCharacters();
            byte[] bytes = _bytes!;
            if ((_characterCount != 0 || _encoderMayHavePendingSurrogate)
                && !TryEncodeBufferedAscii(characters, bytes)) {
                EncodeBufferedCharacters(flushEncoder: true);
            }
            if (bytes.Length - _byteCount < sizeHint) FlushBytes();
            offset = _byteCount;
            return bytes;
        }

        /// <summary>Adds <paramref name="count"/> bytes written after the offset <see cref="ReserveUtf8"/> gave to the output.</summary>
        internal void CommitUtf8(int count) => _byteCount += count;
    }
}
#endif
