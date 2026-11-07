#if NET8_0_OR_GREATER
namespace OfficeIMO.CSV;

public sealed partial class CsvRowWriter
{
    private char[]? _propertyCharacters;
    private int _propertyCharacterCount;

    /// <summary>Stages text into the bounded row buffer before any output is delivered.</summary>
    private void StageProjectedSpanText(int ordinal, string? text)
    {
        long worstCaseLength = (long)_propertyCharacterCount + (text?.Length ?? 0) * 2L + 3;
        if (_propertyFallbackColumn >= 0 || worstCaseLength > _propertyCharacters!.Length)
        {
            if (_propertyFallbackColumn < 0)
            {
                _propertyFallbackColumn = ordinal;
                _propertyFallbackPrefixLength = _propertyCharacterCount;
            }
            StageProjectedText(ordinal, text);
            return;
        }
        if (ordinal > 0) _propertyCharacters[_propertyCharacterCount++] = _delimiter;
        if (text != null)
            _propertyCharacterCount += CsvWriter.WriteEscapedDefault(
                _propertyCharacters.AsSpan(_propertyCharacterCount), text, _delimiter);
    }

    /// <summary>Formats a known invariant scalar directly into the staged row's available space.</summary>
    private static void StageProjectedScalarSpan<T>(CsvRowWriter writer, int ordinal, T value)
        where T : ISpanFormattable
    {
        if (writer._propertyFallbackColumn < 0)
        {
            int start = writer._propertyCharacterCount + (ordinal > 0 ? 1 : 0);
            if (start <= writer._propertyCharacters!.Length &&
                value.TryFormat(writer._propertyCharacters.AsSpan(start), out int written,
                    default, CultureInfo.InvariantCulture))
            {
                if (ordinal > 0) writer._propertyCharacters[writer._propertyCharacterCount] = writer._delimiter;
                writer._propertyCharacterCount = start + written;
                return;
            }
            writer._propertyFallbackColumn = ordinal;
            writer._propertyFallbackPrefixLength = writer._propertyCharacterCount;
        }
        StageProjectedScalar(writer, ordinal, value);
    }
}
#endif
