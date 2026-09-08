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

#if NET6_0_OR_GREATER
    // Each caller supplies at most 256 source characters and reserves enough
    // room to double every quote. The field's closing quote follows all chunks.
    private static int EscapeQuotedChunk(ReadOnlySpan<char> text, Span<char> escaped)
    {
        int count = 0;
        foreach (char character in text)
        {
            escaped[count++] = character;
            if (character == '"') escaped[count++] = '"';
        }
        return count;
    }
#endif

    [System.Runtime.CompilerServices.MethodImpl(System.Runtime.CompilerServices.MethodImplOptions.NoInlining)]
    private static void AppendDenseQuotedText(StringBuilder buffer, string text, int start)
    {
#if NET6_0_OR_GREATER
        if (text.Length - start > 256)
        {
            AppendDenseQuotedSpan(buffer, text.AsSpan(start));
            return;
        }
        if (TryAppendQuoteRun(buffer, text.AsSpan(start))) return;
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

#if NET6_0_OR_GREATER
    [System.Runtime.CompilerServices.MethodImpl(System.Runtime.CompilerServices.MethodImplOptions.NoInlining)]
    private static void AppendDenseQuotedSpan(StringBuilder buffer, ReadOnlySpan<char> remaining)
    {
        if (TryAppendQuoteRun(buffer, remaining)) return;
        AppendQuotedChunks(buffer, remaining);
    }

    [System.Runtime.CompilerServices.MethodImpl(System.Runtime.CompilerServices.MethodImplOptions.AggressiveInlining)]
    private static bool TryAppendQuoteRun(StringBuilder buffer, ReadOnlySpan<char> remaining)
    {
#if NET8_0_OR_GREATER
        if (!remaining.IsEmpty && remaining[0] == '"' && remaining.IndexOfAnyExcept('"') < 0)
        {
            buffer.Append('"', checked(remaining.Length * 2));
            buffer.Append('"');
            return true;
        }
#endif
        return false;
    }

    // Allocate once per field, in a separate call from the writer's row loop.
    private static void AppendQuotedChunks(StringBuilder buffer, ReadOnlySpan<char> remaining)
    {
        Span<char> escaped = stackalloc char[512];
        while (!remaining.IsEmpty)
        {
            int length = Math.Min(remaining.Length, 256);
            int written = EscapeQuotedChunk(remaining.Slice(0, length), escaped);
            buffer.Append(escaped.Slice(0, written));
            remaining = remaining.Slice(length);
        }
        buffer.Append('"');
    }
#endif

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
                WriteDenseQuotedText(writer, text, start);
                return;
            }
        } while (quote >= 0);

        WriteTextSegment(writer, text, start, text.Length - start);
        writer.Write('"');
    }

    [System.Runtime.CompilerServices.MethodImpl(System.Runtime.CompilerServices.MethodImplOptions.NoInlining)]
    private static void WriteDenseQuotedText(TextWriter writer, string text, int start)
    {
#if NET6_0_OR_GREATER
        ReadOnlySpan<char> remaining = text.AsSpan(start);
        Span<char> escaped = stackalloc char[Math.Min(remaining.Length, 256) * 2];
        while (!remaining.IsEmpty)
        {
            int length = Math.Min(remaining.Length, 256);
            int written = EscapeQuotedChunk(remaining.Slice(0, length), escaped);
            writer.Write(escaped.Slice(0, written));
            remaining = remaining.Slice(length);
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
                if (text.Length > 256)
                {
                    AppendDenseQuotedSpan(buffer, text);
                    return;
                }
                if (TryAppendQuoteRun(buffer, text)) return;
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
