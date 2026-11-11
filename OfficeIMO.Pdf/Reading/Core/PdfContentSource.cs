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

    public bool TryParseDouble(int start, int length, out double value) =>
        TryParseDouble(start, length, NumberStyles.Float, out value);

    public bool TryParseDouble(int start, int length, NumberStyles styles, out double value) {
#pragma warning disable CA1846 // Keep netstandard2.0-safe parsing.
        return double.TryParse(
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
    private readonly byte[] _bytes;

    internal PdfByteContentText(byte[] bytes) => _bytes = bytes;

    public int Length => _bytes.Length;

    public char this[int index] => (char)_bytes[index];
    public object Source => _bytes;

    public string Substring(int start, int length) {
        if (start < 0 || length < 0 || start > _bytes.Length - length) throw new ArgumentOutOfRangeException(nameof(start));
        return PdfEncoding.Latin1GetString(_bytes, start, length);
    }

    public string Substring(int start, int length, System.Threading.CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();
        if (start < 0 || length < 0 || start > _bytes.Length - length) throw new ArgumentOutOfRangeException(nameof(start));
        if (length == 0) return string.Empty;
        return PdfEncoding.Latin1GetStringCancellable(_bytes, start, length, cancellationToken);
    }

    public bool TryParseDouble(int start, int length, out double value) =>
        TryParseDouble(start, length, NumberStyles.Float, out value);

    public bool TryParseDouble(int start, int length, NumberStyles styles, out double value) {
#if NET8_0_OR_GREATER
        if (length <= 128) {
            Span<char> characters = stackalloc char[length];
            for (int index = 0; index < length; index++) characters[index] = (char)_bytes[start + index];
            return double.TryParse(characters, styles, CultureInfo.InvariantCulture, out value);
        }
#endif
        return double.TryParse(Substring(start, length), styles, CultureInfo.InvariantCulture, out value);
    }

    public bool TryParseInt32(int start, int length, out int value) {
#if NET8_0_OR_GREATER
        if (length <= 128) {
            Span<char> characters = stackalloc char[length];
            for (int index = 0; index < length; index++) characters[index] = (char)_bytes[start + index];
            return int.TryParse(characters, NumberStyles.Integer, CultureInfo.InvariantCulture, out value);
        }
#endif
        return int.TryParse(Substring(start, length), NumberStyles.Integer, CultureInfo.InvariantCulture, out value);
    }
}
