using System;
using System.IO;
using System.Text;
using System.Buffers;

namespace OfficeIMO.SharedSource.IO {
    /// <summary>
    /// Buffers text and validated UTF-8 bytes in pooled arrays for streaming exports.
    /// </summary>
    internal sealed partial class PooledUtf8TextWriter : TextWriter {
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
#if NET8_0_OR_GREATER
        private bool _encoderMayHavePendingSurrogate;
        private readonly bool _canWriteAsciiBeforeUtf8;
#endif

        internal PooledUtf8TextWriter(Stream stream, Encoding encoding, int bufferSize, bool leaveOpen = false) {
            _stream = stream ?? throw new ArgumentNullException(nameof(stream));
            _encoding = encoding ?? throw new ArgumentNullException(nameof(encoding));
            if (bufferSize <= 0) {
                throw new ArgumentOutOfRangeException(nameof(bufferSize));
            }

            _encoder = encoding.GetEncoder();
#if NET8_0_OR_GREATER
            // Custom encoders and fallbacks can observe flush boundaries. Preserve
            // their existing encoder calls even when the buffered text is ASCII.
            _canWriteAsciiBeforeUtf8 = encoding.GetType() == typeof(UTF8Encoding)
                && (_encoder.Fallback is EncoderReplacementFallback || _encoder.Fallback is EncoderExceptionFallback);
#endif
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

        /// <summary>
        /// Appends three small text fragments with one buffer check, without allocating a
        /// concatenated string. Boundary and long-value writes retain the normal encoder path.
        /// </summary>
        internal void WriteFragments(string first, string second, string third) {
            char[] characters = GetCharacters();
            long length = (long)first.Length + second.Length + third.Length;
            if (length < DirectEncodingThreshold && length <= characters.Length - _characterCount) {
                int offset = _characterCount;
                first.CopyTo(0, characters, offset, first.Length);
                offset += first.Length;
                second.CopyTo(0, characters, offset, second.Length);
                offset += second.Length;
                third.CopyTo(0, characters, offset, third.Length);
                _characterCount += (int)length;
                return;
            }

            Write(first);
            Write(second);
            Write(third);
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
                _stream.Flush();
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
#if NET8_0_OR_GREATER
            if (!flushEncoder && TryEncodeBufferedAscii(characters, bytes)) return;
#endif
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
#if NET8_0_OR_GREATER
            _encoderMayHavePendingSurrogate = !flushEncoder && _characterCount > 0
                && char.IsHighSurrogate(characters[_characterCount - 1]);
#endif
            _characterCount = 0;
        }

#if NET8_0_OR_GREATER
        private bool TryEncodeBufferedAscii(char[] characters, byte[] bytes) {
            // ASCII never changes encoder state. A pending surrogate requires the
            // encoder even when the following buffered characters are all ASCII.
            if (!_encoderMayHavePendingSurrogate && _encoding is UTF8Encoding
                && _characterCount <= bytes.Length - _byteCount
                && System.Text.Ascii.FromUtf16(characters.AsSpan(0, _characterCount),
                    bytes.AsSpan(_byteCount), out int bytesWritten) == OperationStatus.Done) {
                _byteCount += bytesWritten;
                _characterCount = 0;
                return true;
            }
            // A non-ASCII character may leave a tentative prefix. Counts stay
            // unchanged so the encoder overwrites it from the original input.
            return false;
        }
#endif

#if NET6_0_OR_GREATER
        // Long cell values go directly through the stateful encoder. Small XML
        // fragments still share a character buffer; both paths accumulate bytes
        // before calling the ZIP stream, preserving large compression writes.
        private void WriteEncoded(ReadOnlySpan<char> characters) {
            EncodeBufferedCharacters(flushEncoder: false);
#if NET8_0_OR_GREATER
            bool endsWithHighSurrogate = !characters.IsEmpty && char.IsHighSurrogate(characters[characters.Length - 1]);
#endif
            byte[] bytes = _bytes!;
            // GetBytes avoids Convert's partial-buffer bookkeeping when the
            // destination can hold even the encoding's worst-case fallback.
            // Keep the same encoder so split surrogate pairs retain their state.
            if (characters.Length <= (bytes.Length - _byteCount) / _minimumByteSpace) {
                _byteCount += _encoder.GetBytes(characters, bytes.AsSpan(_byteCount), flush: false);
#if NET8_0_OR_GREATER
                _encoderMayHavePendingSurrogate = endsWithHighSurrogate;
#endif
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
#if NET8_0_OR_GREATER
            _encoderMayHavePendingSurrogate = endsWithHighSurrogate;
#endif
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
