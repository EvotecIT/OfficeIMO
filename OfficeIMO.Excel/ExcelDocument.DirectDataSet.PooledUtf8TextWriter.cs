using System.Buffers;

namespace OfficeIMO.Excel {
    /// <summary>
    /// Buffers UTF-8 text in shared arrays so repeated streaming exports do not
    /// allocate a large character and byte buffer for every worksheet.
    /// </summary>
    internal sealed class PooledUtf8TextWriter : TextWriter {
        private readonly Stream _stream;
        private readonly Encoding _encoding;
        private readonly Encoder _encoder;
        private readonly bool _leaveOpen;
        private char[]? _characters;
        private byte[]? _bytes;
        private int _characterCount;

        internal PooledUtf8TextWriter(Stream stream, Encoding encoding, int bufferSize, bool leaveOpen = false) {
            _stream = stream ?? throw new ArgumentNullException(nameof(stream));
            _encoding = encoding ?? throw new ArgumentNullException(nameof(encoding));
            if (bufferSize <= 0) {
                throw new ArgumentOutOfRangeException(nameof(bufferSize));
            }

            _encoder = encoding.GetEncoder();
            _leaveOpen = leaveOpen;
            _characters = ArrayPool<char>.Shared.Rent(bufferSize);
            _bytes = ArrayPool<byte>.Shared.Rent(encoding.GetMaxByteCount(_characters.Length));
        }

        public override Encoding Encoding => _encoding;

        public override void Write(char value) {
            char[] characters = GetCharacters();
            if (_characterCount == characters.Length) {
                FlushBuffer(flushEncoder: false);
            }

            characters[_characterCount++] = value;
        }

        public override void Write(string? value) {
            if (value == null || value.Length == 0) {
                return;
            }

            int sourceIndex = 0;
            while (sourceIndex < value.Length) {
                char[] characters = GetCharacters();
                if (_characterCount == characters.Length) {
                    FlushBuffer(flushEncoder: false);
                }

                int copyCount = Math.Min(characters.Length - _characterCount, value.Length - sourceIndex);
                value.CopyTo(sourceIndex, characters, _characterCount, copyCount);
                _characterCount += copyCount;
                sourceIndex += copyCount;
            }
        }

        public override void Write(char[] buffer, int index, int count) {
            if (buffer == null) {
                throw new ArgumentNullException(nameof(buffer));
            }
            if (index < 0 || count < 0 || index > buffer.Length - count) {
                throw new ArgumentOutOfRangeException(index < 0 ? nameof(index) : nameof(count));
            }

            WriteCharacters(buffer, index, count);
        }

#if NET6_0_OR_GREATER
        public override void Write(ReadOnlySpan<char> buffer) {
            while (!buffer.IsEmpty) {
                char[] characters = GetCharacters();
                if (_characterCount == characters.Length) {
                    FlushBuffer(flushEncoder: false);
                }

                int copyCount = Math.Min(characters.Length - _characterCount, buffer.Length);
                buffer.Slice(0, copyCount).CopyTo(characters.AsSpan(_characterCount));
                _characterCount += copyCount;
                buffer = buffer.Slice(copyCount);
            }
        }
#endif

        public override void Flush() {
            FlushBuffer(flushEncoder: false);
            _stream.Flush();
        }

        protected override void Dispose(bool disposing) {
            if (!disposing || _characters == null) {
                base.Dispose(disposing);
                return;
            }

            char[] characters = _characters;
            byte[] bytes = _bytes!;
            try {
                FlushBuffer(flushEncoder: true);
                if (!_leaveOpen) {
                    _stream.Dispose();
                }
            } finally {
                _characters = null;
                _bytes = null;
                _characterCount = 0;
                ArrayPool<char>.Shared.Return(characters, clearArray: true);
                ArrayPool<byte>.Shared.Return(bytes, clearArray: true);
                base.Dispose(disposing);
            }
        }

        private char[] GetCharacters()
            => _characters ?? throw new ObjectDisposedException(nameof(PooledUtf8TextWriter));

        private void WriteCharacters(char[] source, int sourceIndex, int count) {
            while (count > 0) {
                char[] characters = GetCharacters();
                if (_characterCount == characters.Length) {
                    FlushBuffer(flushEncoder: false);
                }

                int copyCount = Math.Min(characters.Length - _characterCount, count);
                Array.Copy(source, sourceIndex, characters, _characterCount, copyCount);
                _characterCount += copyCount;
                sourceIndex += copyCount;
                count -= copyCount;
            }
        }

        private void FlushBuffer(bool flushEncoder) {
            char[] characters = GetCharacters();
            byte[] bytes = _bytes!;
            int written = _encoder.GetBytes(
                characters,
                charIndex: 0,
                charCount: _characterCount,
                bytes,
                byteIndex: 0,
                flush: flushEncoder);
            _characterCount = 0;
            if (written != 0) {
                _stream.Write(bytes, 0, written);
            }
        }
    }
}
