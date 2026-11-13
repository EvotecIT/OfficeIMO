#if NET8_0_OR_GREATER
using System.Diagnostics.CodeAnalysis;
using System.Linq.Expressions;
using System.Reflection;

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
        if (text == null) return;
        // Projected rows are comma-separated; the room checked above holds every escaped form of the text.
        Span<char> destination = _propertyCharacters.AsSpan(_propertyCharacterCount);
        if (CsvWriter.TryCopyPlainText(text, destination, out _))
            _propertyCharacterCount += text.Length;
        else if (text.Length <= CsvWriter.EscapedShortTextLength && CsvWriter.TryWriteEscapedShortText(text, destination, out int written))
            _propertyCharacterCount += written;
        else
            _propertyCharacterCount += CsvWriter.WriteEscapedDefault(destination, text, _delimiter);
    }

    /// <summary>Formats a known invariant scalar directly into the staged row's available space.</summary>
    private static void StageProjectedScalarSpan<T>(CsvRowWriter writer, int ordinal, T value)
        where T : ISpanFormattable, IUtf8SpanFormattable
    {
        if (writer._propertyFallbackColumn < 0)
        {
            int start = writer._propertyCharacterCount + (ordinal > 0 ? 1 : 0);
            if (start <= writer._propertyCharacters!.Length &&
                CsvWriter.TryFormatInvariant(value, writer._propertyCharacters.AsSpan(start), out int written))
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

    /// <summary>Formats a date with the configured format, culture and UTC conversion, as row writing does.</summary>
    private static void StageProjectedFormattedDate(CsvRowWriter writer, int ordinal, DateTime value)
    {
        CsvSaveOptions options = writer._options;
        if (options.UseUtc) value = value.ToUniversalTime();
        string? format = string.IsNullOrEmpty(options.DateTimeFormat) ? null : options.DateTimeFormat;
        if (writer._propertyFallbackColumn < 0)
        {
            int start = writer._propertyCharacterCount + (ordinal > 0 ? 1 : 0);
            if (start <= writer._propertyCharacters!.Length &&
                CsvWriter.TryFormatDate(value, writer._propertyCharacters.AsSpan(start), out int written, format, options.Culture) &&
                !CsvWriter.NeedsQuotes(writer._propertyCharacters.AsSpan(start, written), writer._delimiter))
            {
                if (ordinal > 0) writer._propertyCharacters[writer._propertyCharacterCount] = writer._delimiter;
                writer._propertyCharacterCount = start + written;
                return;
            }
        }
        writer.StageProjectedSpanText(ordinal, value.ToString(format, options.Culture));
    }

    /// <summary>Stages any value with the configured date, UTC and null rules of row writing.</summary>
    private void StageProjectedFormattedValue(int ordinal, object? value) =>
        StageProjectedSpanText(ordinal, CsvWriter.FormatValue(value, _options.Culture, _options.DateTimeFormat, _options.UseUtc, _options.NullValue));

    /// <summary>Stages text, writing the configured null marker for a null value.</summary>
    private void StageProjectedFormattedText(int ordinal, string? value) =>
        StageProjectedSpanText(ordinal, value ?? _options.NullValue);

    /// <summary>
    /// Builds property writes for date, UTC and null options. Dates are formatted into the row buffer, null
    /// values write the null marker, and other scalars keep the default invariant staging; date-time offsets
    /// and untyped values go through the row writer's value formatting.
    /// </summary>
    [RequiresUnreferencedCode("Builds typed property expressions from runtime schema metadata.")]
    private static Expression CreateFormattedPropertySpanWriteExpression(ParameterExpression value,
        Expression defaultText, ParameterExpression writer, ParameterExpression culture, int ordinal, out bool usesTypedValue)
    {
        Expression ordinalConstant = Expression.Constant(ordinal);
        if (value.Type == typeof(string))
        {
            usesTypedValue = false;
            MethodInfo formattedText = typeof(CsvRowWriter).GetMethod(nameof(StageProjectedFormattedText), BindingFlags.NonPublic | BindingFlags.Instance)!;
            return Expression.Call(writer, formattedText, ordinalConstant, value);
        }

        MethodInfo formattedValue =
            typeof(CsvRowWriter).GetMethod(nameof(StageProjectedFormattedValue), BindingFlags.NonPublic | BindingFlags.Instance)!;
        Type scalarType = Nullable.GetUnderlyingType(value.Type) ?? value.Type;
        MethodInfo? scalarMethod = scalarType == typeof(DateTime)
            ? typeof(CsvRowWriter).GetMethod(nameof(StageProjectedFormattedDate), BindingFlags.NonPublic | BindingFlags.Static)
            : scalarType == typeof(DateTimeOffset) ? null : CsvWriter.GetPropertyScalarStageMethod(scalarType, ProjectedStaging.Characters);
        usesTypedValue = scalarMethod != null;
        if (scalarMethod == null)
        {
            return Expression.Call(writer, formattedValue, ordinalConstant, Expression.Convert(value, typeof(object)));
        }

        Expression scalarValue = scalarType == value.Type ? value : Expression.Property(value, "Value");
        Expression scalar = Expression.Call(scalarMethod, writer, ordinalConstant, scalarValue);
        if (scalarType != typeof(DateTime))
        {
            Expression invariant = Expression.ReferenceEqual(culture,
                Expression.Property(null, typeof(CultureInfo), nameof(CultureInfo.InvariantCulture)));
            scalar = Expression.IfThenElse(invariant, scalar,
                Expression.Call(writer, formattedValue, ordinalConstant, Expression.Convert(scalarValue, typeof(object))));
        }

        if (scalarType == value.Type) return scalar;
        Expression nullMarker = Expression.Call(writer, formattedValue, ordinalConstant, Expression.Constant(null, typeof(object)));
        return Expression.IfThenElse(Expression.Property(value, "HasValue"), scalar, nullMarker);
    }
}
#endif
