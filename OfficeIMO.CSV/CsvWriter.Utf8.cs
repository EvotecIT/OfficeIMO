#if NET8_0_OR_GREATER
#nullable enable

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
            else if (options.QuoteMode != CsvQuoteMode.Never && (options.QuoteMode == CsvQuoteMode.Always || forceQuote))
            {
                writer.Write("\"\"");
            }
        }
        writer.Write(options.NewLine);
    }

    private static void WriteEscapedUtf8(PooledUtf8TextWriter writer, ReadOnlySpan<byte> text, ReadOnlySpan<byte> delimiter, CsvSaveOptions options, bool forceQuote)
    {
        bool prefixApostrophe = options.FormulaInjectionPolicy == CsvFormulaInjectionPolicy.Escape && StartsWithFormulaTrigger(text);
        bool quote = options.QuoteMode != CsvQuoteMode.Never
            && (options.QuoteMode == CsvQuoteMode.Always || forceQuote || NeedsUtf8Quotes(text, delimiter, prefixApostrophe));
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
