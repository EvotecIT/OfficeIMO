#if NET8_0_OR_GREATER
namespace OfficeIMO.SharedSource.IO {
    internal sealed partial class PooledUtf8TextWriter {
        /// <summary>Copies validated UTF-8 bytes after completing preceding character writes.</summary>
        internal void WriteUtf8(ReadOnlySpan<byte> value) {
            char[] characters = GetCharacters();
            byte[] bytes = _bytes!;
            // A complete UTF-8 value cannot continue a pending UTF-16 surrogate.
            if ((_characterCount != 0 || _encoderMayHavePendingSurrogate)
                && (!_canWriteAsciiBeforeUtf8 || !TryEncodeBufferedAscii(characters, bytes))) {
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
    }
}
#endif
