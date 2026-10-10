#nullable enable

#if NET8_0_OR_GREATER
using System.Numerics;
using System.Runtime.CompilerServices;

namespace OfficeIMO.CSV;

internal sealed partial class CsvDataReader
{
    // Date, boolean and GUID parsers consume characters. Bound their stack-backed
    // transcoding; unusually long scalar text keeps the existing string fallback.
    private const int MaximumStackScalarTextBytes = 128;

    /// <summary>
    /// Borrows scalar text only for the same unconverted string columns exposed by
    /// the public UTF-8 capability. Schema converters, null markers and normalized
    /// parser fallbacks continue through the existing value conversion path.
    /// </summary>
    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    private bool TryGetDirectUtf8Text(int ordinal, out ReadOnlySpan<byte> text)
    {
        if (_useDirectTextSourceStrings && _utf8StreamTextRowSource is not null)
        {
            if (!_hasCurrentTextRow)
            {
                ThrowInvalidTextReaderState();
            }

            return _utf8StreamTextRowSource.TryGetUtf8Text(ordinal, out text);
        }

        text = default;
        return false;
    }

    /// <summary>
    /// Uses the platform UTF-8 numeric parser with the getter's existing styles
    /// and culture, preserving integer width, overflow and decimal precision.
    /// A failed parse leaves error construction and conversion to the old path.
    /// </summary>
    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    private bool TryGetDirectUtf8Number<T>(int ordinal, out T value)
        where T : struct, INumberBase<T>
    {
        if (TryGetDirectUtf8Text(ordinal, out var text))
        {
            return T.TryParse(text, NumberStyles.Any, _culture, out value);
        }

        value = default;
        return false;
    }

    private bool TryGetDirectUtf8DateTime(int ordinal, out DateTime value)
    {
        if (TryGetDirectUtf8Text(ordinal, out var text) && text.Length <= MaximumStackScalarTextBytes)
        {
            Span<char> characters = stackalloc char[text.Length];
            int length = Encoding.UTF8.GetChars(text, characters);
            return CsvDataProjectionConverter.TryParseDateTime(
                characters.Slice(0, length), _culture, _dateTimeFormats, out value);
        }

        value = default;
        return false;
    }

    private bool TryGetDirectUtf8Boolean(int ordinal, out bool value)
    {
        if (TryGetDirectUtf8Text(ordinal, out var text) && text.Length <= MaximumStackScalarTextBytes)
        {
            Span<char> characters = stackalloc char[text.Length];
            int length = Encoding.UTF8.GetChars(text, characters);
            if (bool.TryParse(characters.Slice(0, length), out value))
            {
                return true;
            }

            if (text.Length == 1 && (text[0] == (byte)'0' || text[0] == (byte)'1'))
            {
                value = text[0] == (byte)'1';
                return true;
            }
        }

        value = default;
        return false;
    }

    private bool TryGetDirectUtf8Guid(int ordinal, out Guid value)
    {
        if (TryGetDirectUtf8Text(ordinal, out var text) && text.Length <= MaximumStackScalarTextBytes)
        {
            Span<char> characters = stackalloc char[text.Length];
            int length = Encoding.UTF8.GetChars(text, characters);
            return Guid.TryParse(characters.Slice(0, length), out value);
        }

        value = default;
        return false;
    }
}
#endif
