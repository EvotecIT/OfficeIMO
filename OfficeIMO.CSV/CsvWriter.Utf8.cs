#if NET8_0_OR_GREATER
namespace OfficeIMO.CSV;

internal static partial class CsvWriter
{
    internal static void WriteUtf8Record(
        PooledUtf8TextWriter writer,
        ReadOnlyMemory<byte>?[] values,
        ReadOnlySpan<byte> delimiter,
        CsvSaveOptions options,
        ISet<string>? quoteFields,
        IReadOnlyList<string> fieldNames)
    {
        for (int i = 0; i < values.Length; i++)
        {
            if (i != 0) writer.WriteUtf8(delimiter);
            bool forceQuote = ShouldQuoteField(quoteFields, fieldNames, i);
            if (values[i] is { } value)
            {
                WriteEscapedUtf8(writer, value.Span, delimiter, options, forceQuote);
            }
            else if (options.NullValue is { } nullText)
            {
                WriteEscapedUtf8(writer, writer.Encoding.GetBytes(nullText), delimiter, options, forceQuote);
            }
            else if (options.ResolvedQuoteMode < CsvQuoteMode.Never && (options.ResolvedQuoteMode == CsvQuoteMode.Always || forceQuote))
            {
                writer.Write("\"\"");
            }
        }
        writer.Write(options.NewLine);
    }

    private static void WriteEscapedUtf8(PooledUtf8TextWriter writer, ReadOnlySpan<byte> text, ReadOnlySpan<byte> delimiter, CsvSaveOptions options, bool forceQuote)
    {
        bool prefixApostrophe = options.FormulaInjectionPolicy == CsvFormulaInjectionPolicy.Escape && StartsWithFormulaTrigger(text);
        CsvQuoteMode quoteMode = options.ResolvedQuoteMode;
        if (quoteMode >= CsvQuoteMode.Never)
        {
            if (prefixApostrophe) writer.Write('\'');
            WriteUnquotedUtf8(writer, text, delimiter, quoteMode);
            return;
        }

        bool quote = quoteMode == CsvQuoteMode.Always || forceQuote || NeedsUtf8Quotes(text, delimiter, prefixApostrophe);
        if (quote) writer.Write('"');
        if (prefixApostrophe) writer.Write('\'');

        if (quote)
        {
            int start = 0;
            for (int i = 0; i < text.Length; i++)
            {
                if (text[i] != (byte)'"') continue;
                writer.WriteUtf8(text.Slice(start, i - start));
                writer.Write("\"\"");
                start = i + 1;
            }
            writer.WriteUtf8(text.Slice(start));
            writer.Write('"');
        }
        else writer.WriteUtf8(text);
    }

    private static void WriteUnquotedUtf8(PooledUtf8TextWriter writer, ReadOnlySpan<byte> text, ReadOnlySpan<byte> delimiter, CsvQuoteMode quoteMode)
    {
        if (quoteMode == BackslashEscapedQuoteMode)
        {
            int start = 0;
            for (int i = 0; i < text.Length; i++)
            {
                string? escape = text[i] switch
                {
                    (byte)'\\' => "\\\\",
                    (byte)'\t' => "\\t",
                    (byte)'\n' => "\\n",
                    (byte)'\r' => "\\r",
                    _ => null
                };
                if (escape is null) continue;
                WriteUnquotedUtf8Segment(writer, text.Slice(start, i - start), delimiter, quoteMode);
                writer.Write(escape);
                start = i + 1;
            }

            WriteUnquotedUtf8Segment(writer, text.Slice(start), delimiter, quoteMode);
            return;
        }

        WriteUnquotedUtf8Segment(writer, text, delimiter, quoteMode);
    }

    private static void WriteUnquotedUtf8Segment(PooledUtf8TextWriter writer, ReadOnlySpan<byte> text, ReadOnlySpan<byte> delimiter, CsvQuoteMode quoteMode)
    {
        if (quoteMode != LiteralUnquotedQuoteMode && !IsUnquotedUtf8Segment(text, delimiter))
            throw CreateUnrepresentableFieldException(quoteMode);
        writer.WriteUtf8(text);
    }

    private static bool IsUnquotedUtf8Segment(ReadOnlySpan<byte> text, ReadOnlySpan<byte> delimiter) =>
        text.IndexOfAny((byte)'\r', (byte)'\n') < 0 && text.IndexOf(delimiter) < 0;

    /// <summary>
    /// Rejects a row that unquoted output cannot represent before any of it is written, so a failed row leaves no
    /// partial record behind.
    /// </summary>
    internal static void ThrowIfUnrepresentableUtf8(ReadOnlyMemory<byte>?[] values, ReadOnlySpan<byte> delimiter, CsvSaveOptions options, Encoding encoding)
    {
        CsvQuoteMode quoteMode = options.ResolvedQuoteMode;
        if (quoteMode < CsvQuoteMode.Never || quoteMode == LiteralUnquotedQuoteMode) return;
        byte[]? nullText = null;
        for (int i = 0; i < values.Length; i++)
        {
            ReadOnlySpan<byte> text;
            if (values[i] is { } value) text = value.Span;
            else if (options.NullValue is { } nullValue) text = nullText ??= encoding.GetBytes(nullValue);
            else continue;
            if (!IsUnquotedUtf8(text, delimiter, quoteMode)) throw CreateUnrepresentableFieldException(quoteMode);
        }
    }

    private static bool IsUnquotedUtf8(ReadOnlySpan<byte> text, ReadOnlySpan<byte> delimiter, CsvQuoteMode quoteMode)
    {
        if (quoteMode != BackslashEscapedQuoteMode) return IsUnquotedUtf8Segment(text, delimiter);
        int start = 0;
        for (int i = 0; i < text.Length; i++)
        {
            if (text[i] is not ((byte)'\\' or (byte)'\t' or (byte)'\n' or (byte)'\r')) continue;
            if (!IsUnquotedUtf8Segment(text.Slice(start, i - start), delimiter)) return false;
            start = i + 1;
        }

        return IsUnquotedUtf8Segment(text.Slice(start), delimiter);
    }

    private static bool NeedsUtf8Quotes(ReadOnlySpan<byte> text, ReadOnlySpan<byte> delimiter, bool prefixApostrophe) =>
        text.IndexOfAny((byte)'"', (byte)'\r', (byte)'\n') >= 0
        || text.IndexOf(delimiter) >= 0
        || (prefixApostrophe && delimiter[0] == (byte)'\'' && text.StartsWith(delimiter.Slice(1)));

    private static bool StartsWithFormulaTrigger(ReadOnlySpan<byte> text)
    {
        int index = 0;
        while (index < text.Length && text[index] == (byte)' ') index++;
        if (index == text.Length) return false;
        return text[index] is (byte)'=' or (byte)'+' or (byte)'-' or (byte)'@' or (byte)'\t' or (byte)'\r' or (byte)'\n';
    }
}
#endif
