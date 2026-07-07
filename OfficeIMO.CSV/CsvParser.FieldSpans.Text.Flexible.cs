#nullable enable

namespace OfficeIMO.CSV;

internal static partial class CsvParser
{
#if NET8_0_OR_GREATER
    private static bool HasUnexpectedQuoteInTextUnquotedField(ReadOnlySpan<char> text, char delimiter, int position)
    {
        var fieldStart = position;
        for (var i = position; i < text.Length; i++)
        {
            var value = text[i];
            if (value == '\r' || value == '\n')
            {
                return false;
            }

            if (value == delimiter)
            {
                fieldStart = i + 1;
                continue;
            }

            if (value == '"')
            {
                return i != fieldStart;
            }
        }

        return false;
    }

    private static int ReadFlexibleTextRecordFieldSpans<TVisitor>(
        ReadOnlySpan<char> text,
        char delimiter,
        bool trim,
        bool emitFields,
        int recordIndex,
        ref int position,
        ref TVisitor fieldVisitor,
        out int firstFieldLength)
        where TVisitor : struct, ICsvFieldSpanVisitor
    {
        var recordStart = position;
        var scan = recordStart;
        var fields = new List<string>(16);

        while (scan <= text.Length)
        {
            var lineEnd = scan;
            while (lineEnd < text.Length && text[lineEnd] != '\r' && text[lineEnd] != '\n')
            {
                lineEnd++;
            }

            var candidate = text.Slice(recordStart, lineEnd - recordStart).ToString();
            if (TryParseQuotedRecord(candidate, delimiter, trim, strictQuotes: false, lineNumber: 0, fields))
            {
                firstFieldLength = fields.Count == 0 ? 0 : fields[0].Length;
                if (emitFields)
                {
                    for (var i = 0; i < fields.Count; i++)
                    {
                        fieldVisitor.VisitFieldValue(recordIndex, i, fields[i]);
                    }
                }

                position = lineEnd;
                if (position < text.Length)
                {
                    ConsumeTextLineSeparator(text, ref position);
                }

                return fields.Count;
            }

            if (lineEnd >= text.Length)
            {
                throw new CsvParseException("Unterminated quoted field.", 0);
            }

            scan = lineEnd;
            ConsumeTextLineSeparator(text, ref scan);
        }

        throw new CsvParseException("Unterminated quoted field.", 0);
    }
#endif
}
