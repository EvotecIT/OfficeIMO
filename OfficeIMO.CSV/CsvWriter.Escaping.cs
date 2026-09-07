#nullable enable

using System.Text;

namespace OfficeIMO.CSV;

internal static partial class CsvWriter
{
    private static void WriteEscaped(TextWriter writer, string text, char delimiter, CsvQuoteMode quoteMode, bool forceQuote)
    {
        if (quoteMode == CsvQuoteMode.Never || !(quoteMode == CsvQuoteMode.Always || forceQuote || NeedsQuotes(text, delimiter)))
        {
            writer.Write(text);
            return;
        }

        writer.Write('"');
        WriteQuotedText(writer, text);
    }

    private static void WriteEscaped(TextWriter writer, string text, string delimiter, CsvQuoteMode quoteMode, bool forceQuote)
    {
        if (quoteMode == CsvQuoteMode.Never || !(quoteMode == CsvQuoteMode.Always || forceQuote || NeedsQuotes(text, delimiter)))
        {
            writer.Write(text);
            return;
        }

        writer.Write('"');
        WriteQuotedText(writer, text);
    }

    private static void WriteEscaped(StringBuilder writer, string text, char delimiter, CsvQuoteMode quoteMode, bool forceQuote)
    {
        if (quoteMode == CsvQuoteMode.Never || !(quoteMode == CsvQuoteMode.Always || forceQuote || NeedsQuotes(text, delimiter)))
        {
            writer.Append(text);
            return;
        }

        writer.Append('"');
        AppendQuotedText(writer, text);
    }

    private static void WriteEscaped(StringBuilder writer, string text, string delimiter, CsvQuoteMode quoteMode, bool forceQuote)
    {
        if (quoteMode == CsvQuoteMode.Never || !(quoteMode == CsvQuoteMode.Always || forceQuote || NeedsQuotes(text, delimiter)))
        {
            writer.Append(text);
            return;
        }

        writer.Append('"');
        AppendQuotedText(writer, text);
    }

    private static void WriteEscapedDefault(StringBuilder writer, string text, char delimiter)
    {
        int specialIndex = IndexOfCsvSpecial(text, delimiter);
        if (specialIndex < 0)
        {
            writer.Append(text);
            return;
        }

        writer.Append('"');
        AppendQuotedText(writer, text, specialIndex);
    }

    private static void WriteEscapedDefault(TextWriter writer, string? text, char delimiter)
    {
        if (text == null) return;
        int specialIndex = IndexOfCsvSpecial(text, delimiter);
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
        WriteQuotedText(writer, text, specialIndex);
    }

    // These helpers receive an already-open quote and write its escaped contents
    // and closing quote. IndexOf skips ordinary runs without per-character writes.
    private static void AppendQuotedText(StringBuilder buffer, string text, int searchStart = 0)
    {
        int quote = text.IndexOf('"', searchStart);
        if (quote < 0)
        {
            buffer.Append(text);
            buffer.Append('"');
            return;
        }

        int start = 0;
        do
        {
            buffer.Append(text, start, quote - start);
            buffer.Append("\"\"");
            start = quote + 1;
            quote = text.IndexOf('"', start);
            if (quote >= 0 && quote - start < 16)
            {
                // Repeated tiny searches cost more than character writes when
                // quotes cluster, as they do in JSON stored inside a CSV field.
                AppendDenseQuotedText(buffer, text, start);
                return;
            }
        } while (quote >= 0);

        buffer.Append(text, start, text.Length - start);
        buffer.Append('"');
    }

    // Keep the dense loop separate from the bulk-copy path so its per-character
    // StringBuilder appends can be optimized without that path's larger body.
    [System.Runtime.CompilerServices.MethodImpl(System.Runtime.CompilerServices.MethodImplOptions.NoInlining)]
    private static void AppendDenseQuotedText(StringBuilder buffer, string text, int start)
    {
#if NET6_0_OR_GREATER
        foreach (char character in text.AsSpan(start))
        {
            if (character == '"') buffer.Append("\"\"");
            else buffer.Append(character);
        }
#else
        for (int index = start; index < text.Length; index++)
        {
            char character = text[index];
            if (character == '"') buffer.Append("\"\"");
            else buffer.Append(character);
        }
#endif

        buffer.Append('"');
    }

    private static void WriteQuotedText(TextWriter writer, string text, int searchStart = 0)
    {
        int quote = text.IndexOf('"', searchStart);
        if (quote < 0)
        {
            writer.Write(text);
            writer.Write('"');
            return;
        }

        int start = 0;
        do
        {
            WriteTextSegment(writer, text, start, quote - start);
            writer.Write("\"\"");
            start = quote + 1;
            quote = text.IndexOf('"', start);
            if (quote >= 0 && quote - start < 16)
            {
#if NET6_0_OR_GREATER
                foreach (char character in text.AsSpan(start))
                {
                    if (character == '"') writer.Write("\"\"");
                    else writer.Write(character);
                }
#else
                for (int index = start; index < text.Length; index++)
                {
                    char character = text[index];
                    if (character == '"') writer.Write("\"\"");
                    else writer.Write(character);
                }
#endif

                writer.Write('"');
                return;
            }
        } while (quote >= 0);

        WriteTextSegment(writer, text, start, text.Length - start);
        writer.Write('"');
    }

    private static void WriteTextSegment(TextWriter writer, string text, int start, int length)
    {
        if (length <= 0) return;
#if NET6_0_OR_GREATER
        writer.Write(text.AsSpan(start, length));
#else
        writer.Write(text.Substring(start, length));
#endif
    }

    private static void AppendEscapedTextDefault(StringBuilder writer, string? text, char delimiter)
    {
        if (text != null) WriteEscapedDefault(writer, text, delimiter);
    }

    internal static bool TextRowNeedsEscaping(string?[] values, char delimiter)
    {
        for (int i = 0; i < values.Length; i++)
        {
            string? text = values[i];
            if (text != null && IndexOfCsvSpecial(text, delimiter) >= 0) return true;
        }

        return false;
    }

    internal static void WritePlainTextRecordBuffered(TextWriter writer, StringBuilder buffer, string?[] values, char delimiter, string newLine)
    {
        if (buffer == null) throw new ArgumentNullException(nameof(buffer));
        buffer.Clear();
        for (int i = 0; i < values.Length; i++)
        {
            if (i > 0) buffer.Append(delimiter);
            if (values[i] != null) buffer.Append(values[i]);
        }

        WriteBufferedRecordLine(writer, buffer, newLine);
    }

    private static int IndexOfCsvSpecial(string text, char delimiter)
    {
#if NET8_0_OR_GREATER
        if (delimiter == ',') return text.AsSpan().IndexOfAny(DefaultCommaQuoteCharacters);
#endif
        for (int i = 0; i < text.Length; i++)
        {
            char ch = text[i];
            if (ch == '"' || ch == '\n' || ch == '\r' || ch == delimiter) return i;
        }

        return -1;
    }

#if NET6_0_OR_GREATER
    private static void AppendEscapedSpan(StringBuilder writer, ReadOnlySpan<char> text, char delimiter, bool prefixApostrophe, CsvQuoteMode quoteMode, bool forceQuote)
    {
        if (quoteMode == CsvQuoteMode.Never || !(quoteMode == CsvQuoteMode.Always || forceQuote || NeedsQuotes(text, delimiter)))
        {
            if (prefixApostrophe) writer.Append('\'');
            writer.Append(text);
            return;
        }

        writer.Append('"');
        if (prefixApostrophe) writer.Append('\'');
        AppendQuotedSpan(writer, text);
    }

    private static void AppendEscapedSpanDefault(StringBuilder writer, ReadOnlySpan<char> text, char delimiter)
    {
        if (!NeedsQuotes(text, delimiter))
        {
            writer.Append(text);
            return;
        }

        writer.Append('"');
        AppendQuotedSpan(writer, text);
    }

    private static void AppendQuotedSpan(StringBuilder buffer, ReadOnlySpan<char> text)
    {
        int quote = text.IndexOf('"');
        while (quote >= 0)
        {
            buffer.Append(text.Slice(0, quote));
            buffer.Append("\"\"");
            text = text.Slice(quote + 1);
            quote = text.IndexOf('"');
            if (quote >= 0 && quote < 16)
            {
                foreach (char character in text)
                {
                    if (character == '"') buffer.Append("\"\"");
                    else buffer.Append(character);
                }

                buffer.Append('"');
                return;
            }
        }

        buffer.Append(text);
        buffer.Append('"');
    }
#endif

    private static bool NeedsQuotes(string text, char delimiter) => IndexOfCsvSpecial(text, delimiter) >= 0;

    private static bool NeedsQuotes(string text, string delimiter) =>
        text.IndexOf(delimiter, StringComparison.Ordinal) >= 0 ||
        text.IndexOf('"') >= 0 || text.IndexOf('\n') >= 0 || text.IndexOf('\r') >= 0;

#if NET6_0_OR_GREATER
    private static bool NeedsQuotes(ReadOnlySpan<char> text, char delimiter)
    {
#if NET8_0_OR_GREATER
        if (delimiter == ',') return text.IndexOfAny(DefaultCommaQuoteCharacters) >= 0;
#endif
        foreach (char ch in text)
        {
            if (ch == '"' || ch == '\n' || ch == '\r' || ch == delimiter) return true;
        }

        return false;
    }
#endif
}
