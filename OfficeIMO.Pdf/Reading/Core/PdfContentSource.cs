using System.Globalization;

namespace OfficeIMO.Pdf;

/// <summary>
/// Decoded content-stream text held either as the decoded bytes or as an existing Latin-1 string.
/// Both forms read one character per byte, so offsets and tokens are the same for either.
/// </summary>
internal readonly struct PdfContentSource {
    private PdfContentSource(string? text, byte[]? bytes, int length) {
        Text = text;
        Bytes = bytes;
        Length = length;
    }

    internal string? Text { get; }
    internal byte[]? Bytes { get; }
    internal int Length { get; }
    internal bool IsNullOrEmpty => Length == 0;

    public static implicit operator PdfContentSource(string? text) => new(text, null, text?.Length ?? 0);

    /// <summary>Reads an owned decoded buffer; the bytes must not change while the source is in use.</summary>
    internal static PdfContentSource FromBytes(byte[] bytes) {
        Guard.NotNull(bytes, nameof(bytes));
        return new PdfContentSource(null, bytes, bytes.Length);
    }

    internal void ThrowIfNull(string paramName) {
        if (Text is null && Bytes is null) throw new ArgumentNullException(paramName, $"Parameter '{paramName}' cannot be null.");
    }

    /// <summary>Ordinal search for an ASCII value, matching <see cref="string.IndexOf(string, StringComparison)"/>.</summary>
    internal int IndexOf(string value) {
        if (Text is not null) return Text.IndexOf(value, StringComparison.Ordinal);
        if (value.Length == 0) return 0;
        byte[] bytes = Bytes!;
#if NET8_0_OR_GREATER
        if (value.Length <= 64) {
            Span<byte> needle = stackalloc byte[value.Length];
            for (int index = 0; index < value.Length; index++) {
                if (value[index] > 0xFF) return -1;
                needle[index] = (byte)value[index];
            }
            return bytes.AsSpan(0, Length).IndexOf(needle);
        }
#endif
        char first = value[0];
        int last = Length - value.Length;
        for (int index = 0; index <= last; index++) {
            if (bytes[index] != first) continue;
            int matched = 1;
            while (matched < value.Length && bytes[index + matched] == value[matched]) matched++;
            if (matched == value.Length) return index;
        }
        return -1;
    }

    /// <summary>Returns the Latin-1 string form, building it only for byte content.</summary>
    public override string ToString() => Text ?? PdfEncoding.Latin1GetString(Bytes!, 0, Length);
}

/// <summary>Character view over content text, specialised per representation by the generic reader.</summary>
internal interface IPdfContentText {
    int Length { get; }
    char this[int index] { get; }
    string Substring(int start, int length);
    string Substring(int start, int length, System.Threading.CancellationToken cancellationToken);
    /// <summary>The underlying string or byte array, for deferred readers of a known window.</summary>
    object Source { get; }
    bool TryParseDouble(int start, int length, out double value);
    bool TryParseDouble(int start, int length, NumberStyles styles, out double value);
    bool TryParseInt32(int start, int length, out int value);
    /// <summary>First ordinal match of an ASCII <paramref name="value"/> lying wholly inside [start, start + count), or -1.</summary>
    int IndexOf(string value, int start, int count);
    /// <summary>
    /// Last match of an ASCII <paramref name="value"/> lying wholly inside [start, start + count), or -1.
    /// Ignoring case folds ASCII letters only, which is what ordinal ignore-case does for Latin-1 text.
    /// </summary>
    int LastIndexOf(string value, int start, int count, bool ignoreCase);
}

internal readonly struct PdfStringContentText : IPdfContentText {
    private readonly string _text;

    internal PdfStringContentText(string text) => _text = text;

    public int Length => _text.Length;
    public char this[int index] => _text[index];
    public string Substring(int start, int length) => _text.Substring(start, length);
    public string Substring(int start, int length, System.Threading.CancellationToken cancellationToken) =>
        PdfEncoding.StringSliceCancellable(_text, start, length, cancellationToken);
    public object Source => _text;

    public int IndexOf(string value, int start, int count) => _text.IndexOf(value, start, count, StringComparison.Ordinal);

    public int LastIndexOf(string value, int start, int count, bool ignoreCase) =>
        count < value.Length
            ? -1
            : _text.LastIndexOf(value, start + count - 1, count, ignoreCase ? StringComparison.OrdinalIgnoreCase : StringComparison.Ordinal);

    public bool TryParseDouble(int start, int length, out double value) =>
        TryParseDouble(start, length, NumberStyles.Float, out value);

    public bool TryParseDouble(int start, int length, NumberStyles styles, out double value) {
#pragma warning disable CA1846 // Keep netstandard2.0-safe parsing.
        return OfficeIMO.Core.Internal.OfficeRoundTripText.TryParse(
#if NET8_0_OR_GREATER
            _text.AsSpan(start, length),
#else
            _text.Substring(start, length),
#endif
            styles,
            CultureInfo.InvariantCulture,
            out value);
#pragma warning restore CA1846
    }

    public bool TryParseInt32(int start, int length, out int value) {
#pragma warning disable CA1846 // Keep netstandard2.0-safe parsing.
        return int.TryParse(
#if NET8_0_OR_GREATER
            _text.AsSpan(start, length),
#else
            _text.Substring(start, length),
#endif
            NumberStyles.Integer,
            CultureInfo.InvariantCulture,
            out value);
#pragma warning restore CA1846
    }
}

internal readonly struct PdfByteContentText : IPdfContentText {
    private readonly ReadOnlyMemory<byte> _bytes;
    // Owned whole arrays are read directly; tokenizers index every character, and going through the memory's span
    // on each read costs more than the byte itself.
    private readonly byte[]? _array;
    private readonly PdfSeekableCursor? _cursor;

    internal PdfByteContentText(byte[] ownedBytes) : this(ownedBytes, ownedBytes) {
        _array = ownedBytes;
    }

    // Borrowed views each get an identity so reusing a buffer cannot reuse a previous scan's indexes.
    internal PdfByteContentText(ReadOnlyMemory<byte> bytes) : this(bytes, bytes) {
    }

    private PdfByteContentText(ReadOnlyMemory<byte> bytes, object source) {
        _bytes = bytes;
        _array = null;
        _cursor = null;
        Source = source;
    }

    internal PdfByteContentText(PdfSeekableInput input, CancellationToken cancellationToken) {
        _bytes = default;
        _array = null;
        _cursor = new PdfSeekableCursor(input, cancellationToken);
        Source = input;
    }

    public int Length => _array?.Length ?? _cursor?.Input.Length ?? _bytes.Length;
    public char this[int index] => (char)ReadByte(index);
    public object Source { get; }

    internal bool IsSeekable => _cursor is not null;

    internal byte ReadByte(int index) =>
        _array is byte[] array ? array[index] : _cursor is null ? _bytes.Span[index] : _cursor.ReadByte(index);

    internal ReadOnlySpan<byte> ReadRange(int start, int length) =>
        _cursor is null ? _bytes.Span.Slice(start, length) : _cursor.ReadRange(start, length);

    internal byte[] CopyRange(int start, int length, CancellationToken cancellationToken) =>
        _cursor is null ? PdfSyntax.CopyBytes(_bytes, start, length, cancellationToken)
                        : _cursor.Input.ReadAt(start, length, cancellationToken);

    internal int IndexOfByte(byte value, int start, int count) =>
        _cursor is null ? PdfByteSearch.IndexOf(_bytes.Span, value, start, count)
                        : _cursor.IndexOf(value, start, count);

    internal int IndexOfBytes(byte[] value, int start, int count) =>
        _cursor is null ? PdfSyntax.IndexOfAscii(_bytes.Span, value, start, start + count)
                        : _cursor.IndexOf(value, start, count);

    internal void BeginRawSecurityScan() => _cursor?.BeginRawSecurityScan();

    internal bool TryFindRawSecurityMarker(byte[] marker, int start, out int index) {
        index = -1;
        return _cursor?.TryFindRawSecurityMarker(marker, start, out index) == true;
    }

    internal PdfStream CreateStream(PdfDictionary dictionary, int start, int length, bool retainSourceSlices,
        CancellationToken cancellationToken, bool decodingFailed = false, string? decodingError = null) {
        if (_cursor is not null) {
            return new PdfStream(dictionary, _cursor.Input, start, length, decodingFailed, decodingError);
        }

        return retainSourceSlices
            ? PdfStream.FromOwnedSource(dictionary, _bytes, start, length, decodingFailed, decodingError)
            : new PdfStream(dictionary, CopyRange(start, length, cancellationToken), decodingFailed, decodingError);
    }
    public int IndexOf(string value, int start, int count) {
        if (start < 0 || count < 0 || start > Length - count) throw new ArgumentOutOfRangeException(nameof(start));
        if (count < value.Length) return -1;
        if (_cursor is not null) {
            Span<byte> marker = value.Length <= 64 ? stackalloc byte[value.Length] : new byte[value.Length];
            for (int index = 0; index < marker.Length; index++) marker[index] = (byte)value[index];
            return _cursor.IndexOf(marker, start, count);
        }
        ReadOnlySpan<byte> bytes = _bytes.Span;
#if NET8_0_OR_GREATER
        if (value.Length <= 64) {
            Span<byte> needle = stackalloc byte[value.Length];
            for (int index = 0; index < value.Length; index++) needle[index] = (byte)value[index];
            int relative = bytes.Slice(start, count).IndexOf(needle);
            return relative < 0 ? -1 : start + relative;
        }
#endif
        if (value.Length == 0) return start;
        int lastStart = start + count - value.Length;
        // Keywords end in rarer letters than they start with ("obj", "stream"), so candidates are found by the last one.
        int anchor = value.Length - 1;
        byte last = (byte)value[anchor];
        for (int index = start; index <= lastStart; index++) {
            int found = PdfByteSearch.IndexOf(bytes, last, index + anchor, lastStart - index + 1);
            if (found < 0) return -1;
            index = found - anchor;
            int matched = 0;
            while (matched < anchor && bytes[index + matched] == value[matched]) matched++;
            if (matched == anchor) return index;
        }
        return -1;
    }

    public int LastIndexOf(string value, int start, int count, bool ignoreCase) {
        if (start < 0 || count < 0 || start > Length - count) throw new ArgumentOutOfRangeException(nameof(start));
        if (count < value.Length) return -1;
        if (_cursor is not null) return _cursor.LastIndexOf(value, start, count, ignoreCase);
        ReadOnlySpan<byte> bytes = _bytes.Span;
#if NET8_0_OR_GREATER
        if (!ignoreCase && value.Length <= 64) {
            Span<byte> needle = stackalloc byte[value.Length];
            for (int index = 0; index < value.Length; index++) needle[index] = (byte)value[index];
            int relative = bytes.Slice(start, count).LastIndexOf(needle);
            return relative < 0 ? -1 : start + relative;
        }
#endif
        for (int index = start + count - value.Length; index >= start; index--) {
            int matched = 0;
            while (matched < value.Length && SameCharacter(bytes[index + matched], value[matched], ignoreCase)) matched++;
            if (matched == value.Length) return index;
        }
        return -1;
    }

    internal static bool SameCharacter(byte actual, char expected, bool ignoreCase) {
        if (actual == expected) return true;
        if (!ignoreCase) return false;
        char folded = (char)(actual | 0x20);
        return folded == (expected | 0x20) && folded >= 'a' && folded <= 'z';
    }

    public string Substring(int start, int length) {
        if (start < 0 || length < 0 || start > Length - length) throw new ArgumentOutOfRangeException(nameof(start));
        return PdfEncoding.Latin1GetString(ReadRange(start, length));
    }

    public string Substring(int start, int length, System.Threading.CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();
        if (start < 0 || length < 0 || start > Length - length) throw new ArgumentOutOfRangeException(nameof(start));
        if (length == 0) return string.Empty;
        return PdfEncoding.Latin1GetStringCancellable(ReadRange(start, length), cancellationToken);
    }

    public bool TryParseDouble(int start, int length, out double value) =>
        TryParseDouble(start, length, NumberStyles.Float, out value);

    public bool TryParseDouble(int start, int length, NumberStyles styles, out double value) {
#if NET8_0_OR_GREATER
        if (length <= 128) {
            ReadOnlySpan<byte> bytes = ReadRange(start, length);
            Span<char> characters = stackalloc char[length];
            for (int index = 0; index < length; index++) characters[index] = (char)bytes[index];
            return OfficeIMO.Core.Internal.OfficeRoundTripText.TryParse(characters, styles, CultureInfo.InvariantCulture, out value);
        }
#endif
        return OfficeIMO.Core.Internal.OfficeRoundTripText.TryParse(Substring(start, length), styles, CultureInfo.InvariantCulture, out value);
    }

    public bool TryParseInt32(int start, int length, out int value) {
#if NET8_0_OR_GREATER
        if (length <= 128) {
            ReadOnlySpan<byte> bytes = ReadRange(start, length);
            Span<char> characters = stackalloc char[length];
            for (int index = 0; index < length; index++) characters[index] = (char)bytes[index];
            return int.TryParse(characters, NumberStyles.Integer, CultureInfo.InvariantCulture, out value);
        }
#endif
        return int.TryParse(Substring(start, length), NumberStyles.Integer, CultureInfo.InvariantCulture, out value);
    }
}
