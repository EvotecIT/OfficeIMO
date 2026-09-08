using System.Buffers;

namespace OfficeIMO.Excel {
    /// <summary>
    /// Buffers UTF-8 text in shared arrays so repeated streaming exports do not
    /// allocate a large character and byte buffer for every worksheet.
    /// </summary>
    internal sealed class PooledUtf8TextWriter : TextWriter {
        private const int DirectEncodingThreshold = 4096;
        private readonly Stream _stream;
        private readonly Encoding _encoding;
        private readonly Encoder _encoder;
        private readonly bool _leaveOpen;
        private readonly int _minimumByteSpace;
        private char[]? _characters;
        private byte[]? _bytes;
        private int _characterCount;
        private int _byteCount;

        internal PooledUtf8TextWriter(Stream stream, Encoding encoding, int bufferSize, bool leaveOpen = false) {
            _stream = stream ?? throw new ArgumentNullException(nameof(stream));
            _encoding = encoding ?? throw new ArgumentNullException(nameof(encoding));
            if (bufferSize <= 0) {
                throw new ArgumentOutOfRangeException(nameof(bufferSize));
            }

            _encoder = encoding.GetEncoder();
            _leaveOpen = leaveOpen;
            _minimumByteSpace = encoding.GetMaxByteCount(1);
            _characters = ArrayPool<char>.Shared.Rent(Math.Min(bufferSize, 16384));
            _bytes = ArrayPool<byte>.Shared.Rent(Math.Max(bufferSize, _minimumByteSpace));
        }

        public override Encoding Encoding => _encoding;

        public override void Write(char value) {
            char[] characters = GetCharacters();
            if (_characterCount == characters.Length) {
                EncodeBufferedCharacters(flushEncoder: false);
            }

            characters[_characterCount++] = value;
        }

        public override void Write(string? value) {
            if (value == null || value.Length == 0) {
                return;
            }
#if NET6_0_OR_GREATER
            if (value.Length >= DirectEncodingThreshold) {
                WriteEncoded(value.AsSpan());
                return;
            }
#endif

            char[] characters = GetCharacters();
            if (value.Length <= characters.Length - _characterCount) {
                value.CopyTo(0, characters, _characterCount, value.Length);
                _characterCount += value.Length;
                return;
            }
            int sourceIndex = 0;
            while (sourceIndex < value.Length) {
                if (_characterCount == characters.Length) {
                    EncodeBufferedCharacters(flushEncoder: false);
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

#if NET6_0_OR_GREATER
            Write(buffer.AsSpan(index, count));
#else
            WriteCharacters(buffer, index, count);
#endif
        }

#if NET6_0_OR_GREATER
        public override void Write(ReadOnlySpan<char> buffer) {
            if (buffer.IsEmpty) return;
            if (buffer.Length >= DirectEncodingThreshold) {
                WriteEncoded(buffer);
                return;
            }
            char[] characters = GetCharacters();
            if (buffer.Length <= characters.Length - _characterCount) {
                buffer.CopyTo(characters.AsSpan(_characterCount));
                _characterCount += buffer.Length;
                return;
            }
            while (!buffer.IsEmpty) {
                if (_characterCount == characters.Length) {
                    EncodeBufferedCharacters(flushEncoder: false);
                }

                int copyCount = Math.Min(characters.Length - _characterCount, buffer.Length);
                buffer.Slice(0, copyCount).CopyTo(characters.AsSpan(_characterCount));
                _characterCount += copyCount;
                buffer = buffer.Slice(copyCount);
            }
        }
#endif

        public override void Flush() {
            EncodeBufferedCharacters(flushEncoder: false);
            FlushBytes();
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
                EncodeBufferedCharacters(flushEncoder: true);
                FlushBytes();
            } finally {
                _characters = null;
                _bytes = null;
                _characterCount = 0;
                _byteCount = 0;
                ArrayPool<char>.Shared.Return(characters, clearArray: true);
                ArrayPool<byte>.Shared.Return(bytes, clearArray: true);
                try {
                    if (!_leaveOpen) _stream.Dispose();
                } finally {
                    base.Dispose(disposing);
                }
            }
        }

        private char[] GetCharacters()
            => _characters ?? throw new ObjectDisposedException(nameof(PooledUtf8TextWriter));

#if !NET6_0_OR_GREATER
        private void WriteCharacters(char[] source, int sourceIndex, int count) {
            while (count > 0) {
                char[] characters = GetCharacters();
                if (_characterCount == characters.Length) {
                    EncodeBufferedCharacters(flushEncoder: false);
                }

                int copyCount = Math.Min(characters.Length - _characterCount, count);
                Array.Copy(source, sourceIndex, characters, _characterCount, copyCount);
                _characterCount += copyCount;
                sourceIndex += copyCount;
                count -= copyCount;
            }
        }
#endif

        private void EncodeBufferedCharacters(bool flushEncoder) {
            char[] characters = GetCharacters();
            if (_characterCount == 0 && !flushEncoder) return;
            byte[] bytes = _bytes!;
            int offset = 0;
            bool completed;
            do {
                EnsureByteSpace();
                _encoder.Convert(
                    characters, offset, _characterCount - offset,
                    bytes, _byteCount, bytes.Length - _byteCount,
                    flushEncoder, out int usedCharacters, out int usedBytes, out completed);
                offset += usedCharacters;
                _byteCount += usedBytes;
                if (!completed) FlushBytes();
            } while (!completed);
            _characterCount = 0;
        }

#if NET6_0_OR_GREATER
        // Long cell values go directly through the stateful encoder. Small XML
        // fragments still share a character buffer; both paths accumulate bytes
        // before calling the ZIP stream, preserving large compression writes.
        private void WriteEncoded(ReadOnlySpan<char> characters) {
            EncodeBufferedCharacters(flushEncoder: false);
            byte[] bytes = _bytes!;
            // GetBytes avoids Convert's partial-buffer bookkeeping when the
            // destination can hold even the encoding's worst-case fallback.
            // Keep the same encoder so split surrogate pairs retain their state.
            if (characters.Length <= (bytes.Length - _byteCount) / _minimumByteSpace) {
                _byteCount += _encoder.GetBytes(characters, bytes.AsSpan(_byteCount), flush: false);
                return;
            }
            bool completed;
            do {
                EnsureByteSpace();
                _encoder.Convert(characters, bytes.AsSpan(_byteCount), flush: false,
                    out int usedCharacters, out int usedBytes, out completed);
                characters = characters.Slice(usedCharacters);
                _byteCount += usedBytes;
                if (!completed) FlushBytes();
            } while (!completed);
        }
#endif

        private void EnsureByteSpace() {
            if (_bytes!.Length - _byteCount < _minimumByteSpace) FlushBytes();
        }

        private void FlushBytes() {
            if (_byteCount == 0) return;
            int count = _byteCount;
            _byteCount = 0;
            _stream.Write(_bytes!, 0, count);
        }
    }
}
