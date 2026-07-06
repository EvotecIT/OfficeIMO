#nullable enable

using System.Globalization;
using System.Text;

namespace OfficeIMO.CSV;

internal static class CsvWriter
{
#if NET8_0_OR_GREATER
    private static readonly System.Buffers.SearchValues<char> DefaultCommaQuoteCharacters =
        System.Buffers.SearchValues.Create(new[] { '"', ',', '\r', '\n' });
#endif

    public static void Write(TextWriter writer, CsvDocument document, CsvSaveOptions options)
    {
        var delimiter = options.Delimiter;
        var culture = options.Culture;
        var includeHeader = options.IncludeHeader;
        var newLine = options.NewLine;
        var formulaInjectionPolicy = options.FormulaInjectionPolicy;
        var quoteMode = options.QuoteMode;
        var quoteFields = CreateQuoteFieldSet(options.QuoteFields);

        if (includeHeader && document.Header.Count > 0)
        {
            WriteRecord(writer, document.Header, delimiter, newLine, CultureInfo.InvariantCulture, formulaInjectionPolicy, quoteMode, quoteFields, document.Header);
        }

        foreach (var row in document.AsEnumerable())
        {
            WriteRecord(writer, row.Values, delimiter, newLine, culture, formulaInjectionPolicy, quoteMode, quoteFields, document.Header);
        }
    }

    internal static HashSet<string>? CreateQuoteFieldSet(string[]? quoteFields)
    {
        if (quoteFields == null || quoteFields.Length == 0)
        {
            return null;
        }

        var set = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var field in quoteFields)
        {
            if (!string.IsNullOrWhiteSpace(field))
            {
                set.Add(field);
            }
        }

        return set.Count == 0 ? null : set;
    }

    internal static void WriteRecord(
        TextWriter writer,
        IEnumerable<object?> values,
        char delimiter,
        string newLine,
        CultureInfo culture,
        CsvFormulaInjectionPolicy formulaInjectionPolicy,
        CsvQuoteMode quoteMode,
        ISet<string>? quoteFields,
        IReadOnlyList<string>? fieldNames)
    {
        var first = true;
        var index = 0;
        foreach (var value in values)
        {
            if (!first)
            {
                writer.Write(delimiter);
            }
            else
            {
                first = false;
            }

            var text = FormatValue(value, culture);
            if (formulaInjectionPolicy == CsvFormulaInjectionPolicy.Escape)
            {
                text = ApplyFormulaInjectionPolicy(text, formulaInjectionPolicy);
            }

            WriteEscaped(writer, text, delimiter, quoteMode, ShouldQuoteField(quoteFields, fieldNames, index));
            index++;
        }

        writer.Write(newLine);
    }

    internal static void WriteRecord<T>(
        TextWriter writer,
        IReadOnlyList<T> values,
        char delimiter,
        string newLine,
        CultureInfo culture,
        CsvFormulaInjectionPolicy formulaInjectionPolicy,
        CsvQuoteMode quoteMode,
        ISet<string>? quoteFields,
        IReadOnlyList<string>? fieldNames)
    {
        for (var i = 0; i < values.Count; i++)
        {
            if (i > 0)
            {
                writer.Write(delimiter);
            }

            var text = FormatValue(values[i], culture);
            if (formulaInjectionPolicy == CsvFormulaInjectionPolicy.Escape)
            {
                text = ApplyFormulaInjectionPolicy(text, formulaInjectionPolicy);
            }

            WriteEscaped(writer, text, delimiter, quoteMode, ShouldQuoteField(quoteFields, fieldNames, i));
        }

        writer.Write(newLine);
    }

    internal static void WriteRecordBuffered<T>(
        TextWriter writer,
        StringBuilder buffer,
        IReadOnlyList<T> values,
        char delimiter,
        string newLine,
        CultureInfo culture,
        CsvFormulaInjectionPolicy formulaInjectionPolicy,
        CsvQuoteMode quoteMode,
        ISet<string>? quoteFields,
        IReadOnlyList<string>? fieldNames)
    {
        if (buffer == null)
        {
            throw new ArgumentNullException(nameof(buffer));
        }

        buffer.Clear();
        for (var i = 0; i < values.Count; i++)
        {
            if (i > 0)
            {
                buffer.Append(delimiter);
            }

            AppendEscapedValue(buffer, values[i], delimiter, culture, formulaInjectionPolicy, quoteMode, ShouldQuoteField(quoteFields, fieldNames, i));
        }

        WriteBufferedRecordLine(writer, buffer, newLine);
    }

    internal static void WriteRecordBuffered(
        TextWriter writer,
        StringBuilder buffer,
        object?[] values,
        char delimiter,
        string newLine,
        CultureInfo culture,
        CsvFormulaInjectionPolicy formulaInjectionPolicy,
        CsvQuoteMode quoteMode,
        ISet<string>? quoteFields,
        IReadOnlyList<string>? fieldNames)
    {
        if (buffer == null)
        {
            throw new ArgumentNullException(nameof(buffer));
        }

        buffer.Clear();
        for (var i = 0; i < values.Length; i++)
        {
            if (i > 0)
            {
                buffer.Append(delimiter);
            }

            AppendEscapedValue(buffer, values[i], delimiter, culture, formulaInjectionPolicy, quoteMode, ShouldQuoteField(quoteFields, fieldNames, i));
        }

        WriteBufferedRecordLine(writer, buffer, newLine);
    }

    internal static void WriteRecordBufferedDefault(
        TextWriter writer,
        StringBuilder buffer,
        object?[] values,
        char delimiter,
        string newLine,
        CultureInfo culture)
    {
        if (buffer == null)
        {
            throw new ArgumentNullException(nameof(buffer));
        }

        buffer.Clear();
        for (var i = 0; i < values.Length; i++)
        {
            if (i > 0)
            {
                buffer.Append(delimiter);
            }

            AppendEscapedValueDefault(buffer, values[i], delimiter, culture);
        }

        WriteBufferedRecordLine(writer, buffer, newLine);
    }

    internal static void WriteRecordBufferedDefault(
        TextWriter writer,
        StringBuilder buffer,
        string?[] values,
        char delimiter,
        string newLine)
    {
        if (buffer == null)
        {
            throw new ArgumentNullException(nameof(buffer));
        }

        buffer.Clear();
        for (var i = 0; i < values.Length; i++)
        {
            if (i > 0)
            {
                buffer.Append(delimiter);
            }

            AppendEscapedTextDefault(buffer, values[i], delimiter);
        }

        WriteBufferedRecordLine(writer, buffer, newLine);
    }

    internal static void WriteRecordBufferedDefault<TState>(
        TextWriter writer,
        StringBuilder buffer,
        int valueCount,
        TState state,
        Func<TState, int, object?> valueAccessor,
        char delimiter,
        string newLine,
        CultureInfo culture)
    {
        if (buffer == null)
        {
            throw new ArgumentNullException(nameof(buffer));
        }

        if (valueAccessor == null)
        {
            throw new ArgumentNullException(nameof(valueAccessor));
        }

        buffer.Clear();
        for (var i = 0; i < valueCount; i++)
        {
            if (i > 0)
            {
                buffer.Append(delimiter);
            }

            AppendEscapedValueDefault(buffer, valueAccessor(state, i), delimiter, culture);
        }

        WriteBufferedRecordLine(writer, buffer, newLine);
    }

    internal static void WriteRecordBufferedDefault<TState>(
        TextWriter writer,
        StringBuilder buffer,
        int valueCount,
        TState state,
        Func<TState, int, string?> valueAccessor,
        char delimiter,
        string newLine)
    {
        if (buffer == null)
        {
            throw new ArgumentNullException(nameof(buffer));
        }

        if (valueAccessor == null)
        {
            throw new ArgumentNullException(nameof(valueAccessor));
        }

        buffer.Clear();
        for (var i = 0; i < valueCount; i++)
        {
            if (i > 0)
            {
                buffer.Append(delimiter);
            }

            AppendEscapedTextDefault(buffer, valueAccessor(state, i), delimiter);
        }

        WriteBufferedRecordLine(writer, buffer, newLine);
    }

    internal static void WriteRecordDefault(
        TextWriter writer,
        string?[] values,
        char delimiter,
        string newLine)
    {
        for (var i = 0; i < values.Length; i++)
        {
            if (i > 0)
            {
                writer.Write(delimiter);
            }

            WriteEscapedDefault(writer, values[i], delimiter);
        }

        writer.Write(newLine);
    }

    internal static void WriteRecordDefaultAdaptive(
        TextWriter writer,
        StringBuilder buffer,
        string?[] values,
        char delimiter,
        string newLine)
    {
        if (!TextRowNeedsEscaping(values, delimiter))
        {
            WritePlainTextRecord(writer, values, delimiter, newLine);
            return;
        }

        WriteRecordBufferedDefault(writer, buffer, values, delimiter, newLine);
    }

    internal static void WriteRecordBufferedAlwaysQuoted(
        TextWriter writer,
        StringBuilder buffer,
        object?[] values,
        char delimiter,
        string newLine,
        CultureInfo culture)
    {
        if (buffer == null)
        {
            throw new ArgumentNullException(nameof(buffer));
        }

        buffer.Clear();
        for (var i = 0; i < values.Length; i++)
        {
            if (i > 0)
            {
                buffer.Append(delimiter);
            }

            AppendAlwaysQuotedValue(buffer, values[i], culture);
        }

        WriteBufferedRecordLine(writer, buffer, newLine);
    }

    internal static void WriteRecordBufferedAlwaysQuoted(
        TextWriter writer,
        StringBuilder buffer,
        string?[] values,
        char delimiter,
        string newLine)
    {
        if (buffer == null)
        {
            throw new ArgumentNullException(nameof(buffer));
        }

        buffer.Clear();
        for (var i = 0; i < values.Length; i++)
        {
            if (i > 0)
            {
                buffer.Append(delimiter);
            }

            AppendAlwaysQuotedTextValue(buffer, values[i]);
        }

        WriteBufferedRecordLine(writer, buffer, newLine);
    }

    internal static void WriteRecordBufferedAlwaysQuoted<TState>(
        TextWriter writer,
        StringBuilder buffer,
        int valueCount,
        TState state,
        Func<TState, int, string?> valueAccessor,
        char delimiter,
        string newLine)
    {
        if (buffer == null)
        {
            throw new ArgumentNullException(nameof(buffer));
        }

        if (valueAccessor == null)
        {
            throw new ArgumentNullException(nameof(valueAccessor));
        }

        buffer.Clear();
        for (var i = 0; i < valueCount; i++)
        {
            if (i > 0)
            {
                buffer.Append(delimiter);
            }

            AppendAlwaysQuotedTextValue(buffer, valueAccessor(state, i));
        }

        WriteBufferedRecordLine(writer, buffer, newLine);
    }

    internal static void WriteRecordBufferedAlwaysQuoted(
        TextWriter writer,
        StringBuilder buffer,
        IReadOnlyList<string> values,
        char delimiter,
        string newLine,
        CultureInfo culture)
    {
        if (buffer == null)
        {
            throw new ArgumentNullException(nameof(buffer));
        }

        buffer.Clear();
        for (var i = 0; i < values.Count; i++)
        {
            if (i > 0)
            {
                buffer.Append(delimiter);
            }

            AppendAlwaysQuotedValue(buffer, values[i], culture);
        }

        WriteBufferedRecordLine(writer, buffer, newLine);
    }


    internal static void WriteRecord<TState>(
        TextWriter writer,
        int valueCount,
        TState state,
        Func<TState, int, object?> valueAccessor,
        char delimiter,
        string newLine,
        CultureInfo culture,
        CsvFormulaInjectionPolicy formulaInjectionPolicy,
        CsvQuoteMode quoteMode,
        ISet<string>? quoteFields,
        IReadOnlyList<string>? fieldNames)
    {
        if (valueAccessor == null)
        {
            throw new ArgumentNullException(nameof(valueAccessor));
        }

        for (var i = 0; i < valueCount; i++)
        {
            if (i > 0)
            {
                writer.Write(delimiter);
            }

            var text = FormatValue(valueAccessor(state, i), culture);
            if (formulaInjectionPolicy == CsvFormulaInjectionPolicy.Escape)
            {
                text = ApplyFormulaInjectionPolicy(text, formulaInjectionPolicy);
            }

            WriteEscaped(writer, text, delimiter, quoteMode, ShouldQuoteField(quoteFields, fieldNames, i));
        }

        writer.Write(newLine);
    }

    internal static void WriteRecordBuffered<TState>(
        TextWriter writer,
        StringBuilder buffer,
        int valueCount,
        TState state,
        Func<TState, int, object?> valueAccessor,
        char delimiter,
        string newLine,
        CultureInfo culture,
        CsvFormulaInjectionPolicy formulaInjectionPolicy,
        CsvQuoteMode quoteMode,
        ISet<string>? quoteFields,
        IReadOnlyList<string>? fieldNames)
    {
        if (buffer == null)
        {
            throw new ArgumentNullException(nameof(buffer));
        }

        if (valueAccessor == null)
        {
            throw new ArgumentNullException(nameof(valueAccessor));
        }

        buffer.Clear();
        for (var i = 0; i < valueCount; i++)
        {
            if (i > 0)
            {
                buffer.Append(delimiter);
            }

            AppendEscapedValue(buffer, valueAccessor(state, i), delimiter, culture, formulaInjectionPolicy, quoteMode, ShouldQuoteField(quoteFields, fieldNames, i));
        }

        WriteBufferedRecordLine(writer, buffer, newLine);
    }

    internal static void WriteTextRecordBuffered<TState>(
        TextWriter writer,
        StringBuilder buffer,
        int valueCount,
        TState state,
        Func<TState, int, string?> valueAccessor,
        char delimiter,
        string newLine,
        CultureInfo culture,
        CsvFormulaInjectionPolicy formulaInjectionPolicy,
        CsvQuoteMode quoteMode,
        ISet<string>? quoteFields,
        IReadOnlyList<string>? fieldNames)
    {
        if (buffer == null)
        {
            throw new ArgumentNullException(nameof(buffer));
        }

        if (valueAccessor == null)
        {
            throw new ArgumentNullException(nameof(valueAccessor));
        }

        buffer.Clear();
        for (var i = 0; i < valueCount; i++)
        {
            if (i > 0)
            {
                buffer.Append(delimiter);
            }

            AppendEscapedValue(buffer, valueAccessor(state, i), delimiter, culture, formulaInjectionPolicy, quoteMode, ShouldQuoteField(quoteFields, fieldNames, i));
        }

        WriteBufferedRecordLine(writer, buffer, newLine);
    }

    private static string FormatValue(object? value, CultureInfo culture)
    {
        if (value is null)
        {
            return string.Empty;
        }

        if (value is IFormattable formattable)
        {
            return formattable.ToString(null, culture);
        }

        return value.ToString() ?? string.Empty;
    }

    private static bool IsKnownCsvSafeSpanFormattedValue(object value)
    {
        return value is decimal
            or int
            or DateTime
            or double
            or long
            or DateTimeOffset
            or Guid
            or TimeSpan
            or float
            or byte
            or sbyte
            or short
            or ushort
            or uint
            or ulong;
    }

    private static void WriteBufferedRecordLine(TextWriter writer, StringBuilder buffer, string newLine)
    {
#if NET6_0_OR_GREATER
        if (string.Equals(newLine, writer.NewLine, StringComparison.Ordinal))
        {
            writer.WriteLine(buffer);
            return;
        }

        writer.Write(buffer);
        writer.Write(newLine);
#else
        if (string.Equals(newLine, writer.NewLine, StringComparison.Ordinal))
        {
            writer.WriteLine(buffer.ToString());
            return;
        }

        buffer.Append(newLine);
        writer.Write(buffer.ToString());
#endif
    }

    private static void AppendEscapedValue(
        StringBuilder buffer,
        object? value,
        char delimiter,
        CultureInfo culture,
        CsvFormulaInjectionPolicy formulaInjectionPolicy,
        CsvQuoteMode quoteMode,
        bool forceQuote)
    {
        if (value is null)
        {
            if (quoteMode == CsvQuoteMode.Always || forceQuote)
            {
                buffer.Append("\"\"");
            }

            return;
        }

        if (value is string text)
        {
            if (formulaInjectionPolicy == CsvFormulaInjectionPolicy.Escape)
            {
                text = ApplyFormulaInjectionPolicy(text, formulaInjectionPolicy);
            }

            WriteEscaped(buffer, text, delimiter, quoteMode, forceQuote);
            return;
        }

#if NET6_0_OR_GREATER
        if (value is bool boolValue)
        {
            AppendEscapedSpan(buffer, boolValue ? "True" : "False", delimiter, prefixApostrophe: false, quoteMode, forceQuote);
            return;
        }

        if (value is ISpanFormattable spanFormattable)
        {
            Span<char> destination = stackalloc char[128];
            if (spanFormattable.TryFormat(destination, out var charsWritten, default, culture))
            {
                var formattedSpan = destination[..charsWritten];
                var prefixApostrophe = formulaInjectionPolicy == CsvFormulaInjectionPolicy.Escape && StartsWithFormulaTrigger(formattedSpan);
                AppendEscapedSpan(buffer, formattedSpan, delimiter, prefixApostrophe, quoteMode, forceQuote);
                return;
            }
        }
#endif

        var formatted = FormatValue(value, culture);
        if (formulaInjectionPolicy == CsvFormulaInjectionPolicy.Escape)
        {
            formatted = ApplyFormulaInjectionPolicy(formatted, formulaInjectionPolicy);
        }

        WriteEscaped(buffer, formatted, delimiter, quoteMode, forceQuote);
    }

    private static void AppendEscapedValueDefault(
        StringBuilder buffer,
        object? value,
        char delimiter,
        CultureInfo culture)
    {
        if (value is null)
        {
            return;
        }

        if (value is string text)
        {
            WriteEscapedDefault(buffer, text, delimiter);
            return;
        }

#if NET6_0_OR_GREATER
        if (value is bool boolValue)
        {
            WriteEscapedDefault(buffer, boolValue ? "True" : "False", delimiter);
            return;
        }

        if (value is ISpanFormattable spanFormattable)
        {
            Span<char> destination = stackalloc char[128];
            if (spanFormattable.TryFormat(destination, out var charsWritten, default, culture))
            {
                var formatted = destination[..charsWritten];
                if (delimiter == ',' && ReferenceEquals(culture, CultureInfo.InvariantCulture) && IsKnownCsvSafeSpanFormattedValue(value))
                {
                    buffer.Append(formatted);
                }
                else
                {
                    AppendEscapedSpanDefault(buffer, formatted, delimiter);
                }

                return;
            }
        }
#endif

        WriteEscapedDefault(buffer, FormatValue(value, culture), delimiter);
    }

    private static void AppendAlwaysQuotedValue(
        StringBuilder buffer,
        object? value,
        CultureInfo culture)
    {
        buffer.Append('"');
        if (value is null)
        {
            buffer.Append('"');
            return;
        }

        if (value is string text)
        {
            AppendQuotedText(buffer, text);
            return;
        }

#if NET6_0_OR_GREATER
        if (value is bool boolValue)
        {
            buffer.Append(boolValue ? "True\"" : "False\"");
            return;
        }

        if (value is ISpanFormattable spanFormattable)
        {
            Span<char> destination = stackalloc char[128];
            if (spanFormattable.TryFormat(destination, out var charsWritten, default, culture))
            {
                AppendQuotedSpan(buffer, destination[..charsWritten]);
                return;
            }
        }
#endif

        AppendQuotedText(buffer, FormatValue(value, culture));
    }

    private static void AppendAlwaysQuotedTextValue(StringBuilder buffer, string? value)
    {
        buffer.Append('"');
        if (value == null)
        {
            buffer.Append('"');
            return;
        }

        AppendQuotedText(buffer, value);
    }

    private static void AppendQuotedText(StringBuilder buffer, string text)
    {
        if (text.IndexOf('"') < 0)
        {
            buffer.Append(text);
            buffer.Append('"');
            return;
        }

        foreach (var ch in text)
        {
            if (ch == '\"')
            {
                buffer.Append("\"\"");
            }
            else
            {
                buffer.Append(ch);
            }
        }

        buffer.Append('"');
    }

#if NET6_0_OR_GREATER
    private static void AppendQuotedSpan(StringBuilder buffer, ReadOnlySpan<char> text)
    {
        if (text.IndexOf('"') < 0)
        {
            buffer.Append(text);
            buffer.Append('"');
            return;
        }

        foreach (var ch in text)
        {
            if (ch == '\"')
            {
                buffer.Append("\"\"");
            }
            else
            {
                buffer.Append(ch);
            }
        }

        buffer.Append('"');
    }
#endif

    private static bool ShouldQuoteField(ISet<string>? quoteFields, IReadOnlyList<string>? fieldNames, int index)
    {
        return quoteFields != null &&
            fieldNames != null &&
            index >= 0 &&
            index < fieldNames.Count &&
            quoteFields.Contains(fieldNames[index]);
    }

    private static string ApplyFormulaInjectionPolicy(string text, CsvFormulaInjectionPolicy policy)
    {
        if (policy != CsvFormulaInjectionPolicy.Escape || !StartsWithFormulaTrigger(text))
        {
            return text;
        }

        return "'" + text;
    }

    private static bool StartsWithFormulaTrigger(string text)
    {
        if (string.IsNullOrEmpty(text))
        {
            return false;
        }

        var index = 0;
        while (index < text.Length && text[index] == ' ')
        {
            index++;
        }

        if (index >= text.Length)
        {
            return false;
        }

        return text[index] == '='
            || text[index] == '+'
            || text[index] == '-'
            || text[index] == '@'
            || text[index] == '\t'
            || text[index] == '\r'
            || text[index] == '\n';
    }

#if NET6_0_OR_GREATER
    private static bool StartsWithFormulaTrigger(ReadOnlySpan<char> text)
    {
        if (text.IsEmpty)
        {
            return false;
        }

        var index = 0;
        while (index < text.Length && text[index] == ' ')
        {
            index++;
        }

        if (index >= text.Length)
        {
            return false;
        }

        return text[index] == '='
            || text[index] == '+'
            || text[index] == '-'
            || text[index] == '@'
            || text[index] == '\t'
            || text[index] == '\r'
            || text[index] == '\n';
    }
#endif

    private static void WriteEscaped(TextWriter writer, string text, char delimiter, CsvQuoteMode quoteMode, bool forceQuote)
    {
        if (quoteMode == CsvQuoteMode.Never)
        {
            writer.Write(text);
            return;
        }

        var needsQuotes = quoteMode == CsvQuoteMode.Always || forceQuote || NeedsQuotes(text, delimiter);
        if (!needsQuotes)
        {
            writer.Write(text);
            return;
        }

        writer.Write('"');
        if (text.IndexOf('"') < 0)
        {
            writer.Write(text);
            writer.Write('"');
            return;
        }

        foreach (var ch in text)
        {
            if (ch == '\"')
            {
                writer.Write("\"\"");
            }
            else
            {
                writer.Write(ch);
            }
        }

        writer.Write('"');
    }

    private static void WriteEscaped(StringBuilder writer, string text, char delimiter, CsvQuoteMode quoteMode, bool forceQuote)
    {
        if (quoteMode == CsvQuoteMode.Never)
        {
            writer.Append(text);
            return;
        }

        var needsQuotes = quoteMode == CsvQuoteMode.Always || forceQuote || NeedsQuotes(text, delimiter);
        if (!needsQuotes)
        {
            writer.Append(text);
            return;
        }

        writer.Append('"');
        if (text.IndexOf('"') < 0)
        {
            writer.Append(text);
            writer.Append('"');
            return;
        }

        foreach (var ch in text)
        {
            if (ch == '\"')
            {
                writer.Append("\"\"");
            }
            else
            {
                writer.Append(ch);
            }
        }

        writer.Append('"');
    }

    private static void WriteEscapedDefault(StringBuilder writer, string text, char delimiter)
    {
        var specialIndex = IndexOfCsvSpecial(text, delimiter);
        if (specialIndex < 0)
        {
            writer.Append(text);
            return;
        }

        writer.Append('"');
        if (text.IndexOf('"', specialIndex) < 0)
        {
            writer.Append(text);
            writer.Append('"');
            return;
        }

        var segmentStart = 0;
        for (var i = specialIndex; i < text.Length; i++)
        {
            if (text[i] == '"')
            {
                writer.Append(text, segmentStart, i - segmentStart);
                writer.Append("\"\"");
                segmentStart = i + 1;
            }
        }

        if (segmentStart < text.Length)
        {
            writer.Append(text, segmentStart, text.Length - segmentStart);
        }

        writer.Append('"');
    }

    private static void WriteEscapedDefault(TextWriter writer, string? text, char delimiter)
    {
        if (text == null)
        {
            return;
        }

        var specialIndex = IndexOfCsvSpecial(text, delimiter);
        if (specialIndex < 0)
        {
            writer.Write(text);
            return;
        }

        WriteEscapedDefault(writer, text, specialIndex);
    }

    private static void WriteEscapedDefault(TextWriter writer, string text, int specialIndex)
    {
        writer.Write('"');
        if (text.IndexOf('"', specialIndex) < 0)
        {
            writer.Write(text);
            writer.Write('"');
            return;
        }

        var segmentStart = 0;
        for (var i = specialIndex; i < text.Length; i++)
        {
            if (text[i] == '"')
            {
                WriteTextSegment(writer, text, segmentStart, i - segmentStart);
                writer.Write("\"\"");
                segmentStart = i + 1;
            }
        }

        if (segmentStart < text.Length)
        {
            WriteTextSegment(writer, text, segmentStart, text.Length - segmentStart);
        }

        writer.Write('"');
    }

    private static void WriteTextSegment(TextWriter writer, string text, int start, int length)
    {
        if (length <= 0)
        {
            return;
        }

#if NET6_0_OR_GREATER
        writer.Write(text.AsSpan(start, length));
#else
        writer.Write(text.Substring(start, length));
#endif
    }

    private static void AppendEscapedTextDefault(StringBuilder writer, string? text, char delimiter)
    {
        if (text == null)
        {
            return;
        }

        WriteEscapedDefault(writer, text, delimiter);
    }

    private static bool TextRowNeedsEscaping(string?[] values, char delimiter)
    {
        for (var i = 0; i < values.Length; i++)
        {
            var text = values[i];
            if (text != null && IndexOfCsvSpecial(text, delimiter) >= 0)
            {
                return true;
            }
        }

        return false;
    }

    private static void WritePlainTextRecord(TextWriter writer, string?[] values, char delimiter, string newLine)
    {
        for (var i = 0; i < values.Length; i++)
        {
            if (i > 0)
            {
                writer.Write(delimiter);
            }

            if (values[i] != null)
            {
                writer.Write(values[i]);
            }
        }

        writer.Write(newLine);
    }

    private static int IndexOfCsvSpecial(string text, char delimiter)
    {
#if NET8_0_OR_GREATER
        if (delimiter == ',')
        {
            return text.AsSpan().IndexOfAny(DefaultCommaQuoteCharacters);
        }
#endif

        for (var i = 0; i < text.Length; i++)
        {
            var ch = text[i];
            if (ch == '"' || ch == '\n' || ch == '\r' || ch == delimiter)
            {
                return i;
            }
        }

        return -1;
    }

#if NET6_0_OR_GREATER
    private static void AppendEscapedSpan(StringBuilder writer, ReadOnlySpan<char> text, char delimiter, bool prefixApostrophe, CsvQuoteMode quoteMode, bool forceQuote)
    {
        if (quoteMode == CsvQuoteMode.Never)
        {
            if (prefixApostrophe)
            {
                writer.Append('\'');
            }

            writer.Append(text);
            return;
        }

        var needsQuotes = quoteMode == CsvQuoteMode.Always || forceQuote || NeedsQuotes(text, delimiter);
        if (!needsQuotes)
        {
            if (prefixApostrophe)
            {
                writer.Append('\'');
            }

            writer.Append(text);
            return;
        }

        writer.Append('"');
        if (prefixApostrophe)
        {
            writer.Append('\'');
        }

        if (text.IndexOf('"') < 0)
        {
            writer.Append(text);
            writer.Append('"');
            return;
        }

        foreach (var ch in text)
        {
            if (ch == '\"')
            {
                writer.Append("\"\"");
            }
            else
            {
                writer.Append(ch);
            }
        }

        writer.Append('"');
    }

    private static void AppendEscapedSpanDefault(StringBuilder writer, ReadOnlySpan<char> text, char delimiter)
    {
        if (!NeedsQuotes(text, delimiter))
        {
            writer.Append(text);
            return;
        }

        writer.Append('"');
        if (text.IndexOf('"') < 0)
        {
            writer.Append(text);
            writer.Append('"');
            return;
        }

        foreach (var ch in text)
        {
            if (ch == '\"')
            {
                writer.Append("\"\"");
            }
            else
            {
                writer.Append(ch);
            }
        }

        writer.Append('"');
    }
#endif

    private static bool NeedsQuotes(string text, char delimiter)
    {
#if NET8_0_OR_GREATER
        if (delimiter == ',')
        {
            return text.AsSpan().IndexOfAny(DefaultCommaQuoteCharacters) >= 0;
        }
#endif

        foreach (var ch in text)
        {
            if (ch == '"' || ch == '\n' || ch == '\r' || ch == delimiter)
            {
                return true;
            }
        }

        return false;
    }

#if NET6_0_OR_GREATER
    private static bool NeedsQuotes(ReadOnlySpan<char> text, char delimiter)
    {
#if NET8_0_OR_GREATER
        if (delimiter == ',')
        {
            return text.IndexOfAny(DefaultCommaQuoteCharacters) >= 0;
        }
#endif

        foreach (var ch in text)
        {
            if (ch == '"' || ch == '\n' || ch == '\r' || ch == delimiter)
            {
                return true;
            }
        }

        return false;
    }
#endif
}
