#nullable enable

namespace OfficeIMO.CSV;

internal static partial class CsvParser
{
#if NET8_0_OR_GREATER
    /// <summary>Decodes a complete quoted field once, without an intermediate builder or character array.</summary>
    private static QuotedRecordParseResult TryReadStandardQuotedFieldDirect(
        string text, ref int index, bool trim, char delimiter, out string value)
    {
        int start = index + 1;
        int closingQuote = FindClosingQuote(text, start, out int escapedQuotes);
        value = string.Empty;
        if (closingQuote < 0) return QuotedRecordParseResult.Incomplete;

        index = closingQuote + 1;
        if (trim)
        {
            while (index < text.Length && text[index] != delimiter && char.IsWhiteSpace(text[index])) index++;
        }
        if (index < text.Length && text[index] != delimiter) return QuotedRecordParseResult.Invalid;

        int length = closingQuote - start;
        value = escapedQuotes == 0
            ? text.Substring(start, length)
            : string.Create(length - escapedQuotes, (text, start, length, escapedQuotes), static (destination, state) =>
                CopyUnescapedQuotedSegment(state.text.AsSpan(state.start, state.length), state.escapedQuotes, destination));
        return QuotedRecordParseResult.Complete;
    }
#endif
}
