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
            WriteUnquotedUtf8(writer, text, quoteMode);
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

    // Fields reach here after ThrowIfUnrepresentableUtf8 accepted the row.
    private static void WriteUnquotedUtf8(PooledUtf8TextWriter writer, ReadOnlySpan<byte> text, CsvQuoteMode quoteMode)
    {
        if (quoteMode != BackslashEscapedQuoteMode)
        {
            writer.WriteUtf8(text);
            return;
        }

        while (true)
        {
            int index = text.IndexOfAny(BackslashEscapes);
            if (index < 0) break;
            writer.WriteUtf8(text.Slice(0, index));
            writer.Write(text[index] switch
            {
                (byte)'\\' => "\\\\",
                (byte)'\t' => "\\t",
                (byte)'\n' => "\\n",
                _ => "\\r"
            });
            text = text.Slice(index + 1);
        }

        writer.WriteUtf8(text);
    }

    private static ReadOnlySpan<byte> BackslashEscapes => "\\\t\n\r"u8;

    /// <summary>
    /// Rejects a row that unquoted output cannot represent before any of it is written, so a failed row leaves no
    /// partial record behind.
    /// </summary>
    internal static void ThrowIfUnrepresentableUtf8(ReadOnlyMemory<byte>?[] values, ReadOnlySpan<byte> delimiter, CsvSaveOptions options, Encoding encoding)
    {
        CsvQuoteMode quoteMode = options.ResolvedQuoteMode;
        if (quoteMode < CsvQuoteMode.Never || quoteMode == LiteralUnquotedQuoteMode) return;
        // Backslash escaping writes tab and line breaks as escapes, so only a delimiter made of other bytes can collide.
        bool backslash = quoteMode == BackslashEscapedQuoteMode;
        if (backslash && delimiter.IndexOfAny(BackslashEscapes) >= 0) return;
        byte[]? nullText = null;
        for (int i = 0; i < values.Length; i++)
        {
            ReadOnlySpan<byte> text;
            if (values[i] is { } value) text = value.Span;
            else if (options.NullValue is { } nullValue) text = nullText ??= encoding.GetBytes(nullValue);
            else continue;
            bool collides = backslash ? text.IndexOf(delimiter) >= 0
                : delimiter.Length == 1 ? text.IndexOfAny((byte)'\r', (byte)'\n', delimiter[0]) >= 0
                : text.IndexOfAny((byte)'\r', (byte)'\n') >= 0 || text.IndexOf(delimiter) >= 0;
            if (collides) throw CreateUnrepresentableFieldException(quoteMode);
        }
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
