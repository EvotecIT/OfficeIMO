namespace OfficeIMO.Yaml;

/// <summary>Unicode encodings a YAML stream may use (YAML 1.2.2 section 5.2).</summary>
internal enum YamlTextEncoding : byte {
    Utf8 = 0,
    Utf16LittleEndian,
    Utf16BigEndian,
    Utf32LittleEndian,
    Utf32BigEndian
}

/// <summary>Detects a stream's encoding from its byte-order mark or its first bytes, as the specification requires.</summary>
internal static class YamlEncodingDetector {
    /// <summary>Detects the encoding from up to four leading bytes and reports the byte-order mark length (0 when absent).</summary>
    public static YamlTextEncoding Detect(ReadOnlySpan<byte> prefix, out int bomLength) {
        int n = prefix.Length;
        byte b0 = n > 0 ? prefix[0] : (byte)0xFF;
        byte b1 = n > 1 ? prefix[1] : (byte)0xFF;
        byte b2 = n > 2 ? prefix[2] : (byte)0xFF;
        byte b3 = n > 3 ? prefix[3] : (byte)0xFF;
        bomLength = 0;
        if (n >= 4 && b0 == 0 && b1 == 0 && b2 == 0xFE && b3 == 0xFF) { bomLength = 4; return YamlTextEncoding.Utf32BigEndian; }
        if (n >= 4 && b0 == 0 && b1 == 0 && b2 == 0) return YamlTextEncoding.Utf32BigEndian;
        if (n >= 4 && b0 == 0xFF && b1 == 0xFE && b2 == 0 && b3 == 0) { bomLength = 4; return YamlTextEncoding.Utf32LittleEndian; }
        if (n >= 4 && b1 == 0 && b2 == 0 && b3 == 0) return YamlTextEncoding.Utf32LittleEndian;
        if (n >= 2 && b0 == 0xFE && b1 == 0xFF) { bomLength = 2; return YamlTextEncoding.Utf16BigEndian; }
        if (n >= 2 && b0 == 0 && b1 != 0) return YamlTextEncoding.Utf16BigEndian;
        if (n >= 2 && b0 == 0xFF && b1 == 0xFE) { bomLength = 2; return YamlTextEncoding.Utf16LittleEndian; }
        if (n >= 2 && b0 != 0 && b1 == 0) return YamlTextEncoding.Utf16LittleEndian;
        if (n >= 3 && b0 == 0xEF && b1 == 0xBB && b2 == 0xBF) { bomLength = 3; return YamlTextEncoding.Utf8; }
        return YamlTextEncoding.Utf8;
    }

    /// <summary>Creates a strict encoding (invalid input throws) without a preamble.</summary>
    public static Encoding CreateEncoding(YamlTextEncoding encoding) => encoding switch {
        YamlTextEncoding.Utf16LittleEndian => new UnicodeEncoding(bigEndian: false, byteOrderMark: false, throwOnInvalidBytes: true),
        YamlTextEncoding.Utf16BigEndian => new UnicodeEncoding(bigEndian: true, byteOrderMark: false, throwOnInvalidBytes: true),
        YamlTextEncoding.Utf32LittleEndian => new UTF32Encoding(bigEndian: false, byteOrderMark: false, throwOnInvalidCharacters: true),
        YamlTextEncoding.Utf32BigEndian => new UTF32Encoding(bigEndian: true, byteOrderMark: false, throwOnInvalidCharacters: true),
        _ => new UTF8Encoding(encoderShouldEmitUTF8Identifier: false, throwOnInvalidBytes: true)
    };

    /// <summary>
    /// Decodes a complete YAML byte buffer once into text, recording the encoding and byte-order mark so a lossless
    /// save can reproduce them. The byte-order mark is not part of the returned text.
    /// </summary>
    public static string Decode(byte[] bytes, int count, out YamlTextEncoding encoding, out bool hasByteOrderMark) {
        encoding = Detect(new ReadOnlySpan<byte>(bytes, 0, Math.Min(count, 4)), out int bomLength);
        hasByteOrderMark = bomLength > 0;
        try {
            return CreateEncoding(encoding).GetString(bytes, bomLength, count - bomLength);
        } catch (DecoderFallbackException ex) {
            throw new YamlException(YamlErrorCodes.InvalidEncoding, "The YAML input is not valid " + encoding + " text.", default, ex);
        }
    }

    /// <summary>Returns the byte-order mark of an encoding.</summary>
    public static byte[] GetByteOrderMark(YamlTextEncoding encoding) => encoding switch {
        YamlTextEncoding.Utf16LittleEndian => new byte[] { 0xFF, 0xFE },
        YamlTextEncoding.Utf16BigEndian => new byte[] { 0xFE, 0xFF },
        YamlTextEncoding.Utf32LittleEndian => new byte[] { 0xFF, 0xFE, 0, 0 },
        YamlTextEncoding.Utf32BigEndian => new byte[] { 0, 0, 0xFE, 0xFF },
        _ => new byte[] { 0xEF, 0xBB, 0xBF }
    };
}

/// <summary>Supplies decoded characters to the scanner's sliding buffer.</summary>
internal abstract class YamlCharSource {
    /// <summary>Reads up to <paramref name="count"/> characters; returns 0 only at the end of input.</summary>
    public abstract int Read(char[] buffer, int offset, int count);

    /// <summary>Gets the input consumed so far in the unit <see cref="YamlReaderOptions.MaxInputBytes"/> uses for this source.</summary>
    public abstract long Consumed { get; }

    /// <summary>Returns pooled buffers. The caller's reader or stream is never disposed.</summary>
    public virtual void Release() { }
}

/// <summary>Reads characters from a caller-owned <see cref="TextReader"/>.</summary>
internal sealed class YamlTextReaderSource : YamlCharSource {
    private readonly TextReader _reader;
    private long _consumed;

    public YamlTextReaderSource(TextReader reader) => _reader = reader;

    public override long Consumed => _consumed;

    public override int Read(char[] buffer, int offset, int count) {
        int n = _reader.Read(buffer, offset, count);
        if (n > 0) _consumed += n;
        return n;
    }
}

/// <summary>Decodes a caller-owned <see cref="Stream"/> after detecting its encoding from the leading bytes.</summary>
internal sealed class YamlStreamSource : YamlCharSource {
    private const int ByteBufferSize = 16 * 1024;
    private readonly Stream _stream;
    private byte[]? _bytes;
    private int _bytePos;
    private int _byteCount;
    private bool _eof;
    private bool _flushed;
    private Decoder? _decoder;
    private long _consumed;

    public YamlStreamSource(Stream stream) {
        _stream = stream;
        _bytes = ArrayPool<byte>.Shared.Rent(ByteBufferSize);
    }

    public YamlTextEncoding Encoding { get; private set; }

    public bool HasByteOrderMark { get; private set; }

    public override long Consumed => _consumed;

    public override int Read(char[] buffer, int offset, int count) {
        byte[] bytes = _bytes ?? throw new ObjectDisposedException(nameof(YamlReader));
        if (_decoder is null) Initialize(bytes);
        Decoder decoder = _decoder!;
        try {
            while (true) {
                if (_bytePos < _byteCount) {
                    decoder.Convert(bytes, _bytePos, _byteCount - _bytePos, buffer, offset, count, false, out int used, out int produced, out _);
                    _bytePos += used;
                    if (produced > 0) return produced;
                    if (used == 0 && _byteCount - _bytePos >= bytes.Length) return 0;
                }
                if (_eof) {
                    if (_flushed) return 0;
                    _flushed = true;
                    decoder.Convert(bytes, _bytePos, _byteCount - _bytePos, buffer, offset, count, true, out int used, out int produced, out _);
                    _bytePos += used;
                    return produced;
                }
                if (_bytePos > 0) {
                    int remaining = _byteCount - _bytePos;
                    if (remaining > 0) Buffer.BlockCopy(bytes, _bytePos, bytes, 0, remaining);
                    _byteCount = remaining;
                    _bytePos = 0;
                }
                int read = _stream.Read(bytes, _byteCount, bytes.Length - _byteCount);
                if (read <= 0) _eof = true;
                else {
                    _byteCount += read;
                    _consumed += read;
                }
            }
        } catch (DecoderFallbackException ex) {
            throw new YamlException(YamlErrorCodes.InvalidEncoding, "The YAML input is not valid " + Encoding + " text.", default, ex);
        }
    }

    private void Initialize(byte[] bytes) {
        while (_byteCount < 4) {
            int read = _stream.Read(bytes, _byteCount, 4 - _byteCount);
            if (read <= 0) { _eof = true; break; }
            _byteCount += read;
            _consumed += read;
        }
        Encoding = YamlEncodingDetector.Detect(new ReadOnlySpan<byte>(bytes, 0, _byteCount), out int bomLength);
        HasByteOrderMark = bomLength > 0;
        _bytePos = bomLength;
        _decoder = YamlEncodingDetector.CreateEncoding(Encoding).GetDecoder();
    }

    public override void Release() {
        byte[]? bytes = _bytes;
        _bytes = null;
        if (bytes is not null) ArrayPool<byte>.Shared.Return(bytes);
    }
}
