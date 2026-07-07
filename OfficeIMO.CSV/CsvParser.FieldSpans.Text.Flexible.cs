#nullable enable

namespace OfficeIMO.CSV;

internal static partial class CsvParser
{
#if NET8_0_OR_GREATER
    private static bool ShouldUseFlexibleTextRecordParsing(ReadOnlySpan<char> text, char delimiter, bool trim, int position)
    {
        var index = position;
        var fieldStart = position;
        while (index < text.Length)
        {
            var value = text[index];
            if (value == '\r' || value == '\n')
            {
                return false;
            }

            if (value == delimiter)
            {
                index++;
                fieldStart = index;
                continue;
            }

            if (value == '"')
            {
                if (index != fieldStart)
                {
                    return true;
                }

                index++;
                while (index < text.Length)
                {
                    var quoteOffset = text[index..].IndexOf('"');
                    if (quoteOffset < 0)
                    {
                        return false;
                    }

                    index += quoteOffset;
                    if (index + 1 < text.Length && text[index + 1] == '"')
                    {
                        index += 2;
                        continue;
                    }

                    index++;
                    break;
                }

                if (trim)
                {
                    while (index < text.Length &&
                        text[index] != delimiter &&
                        text[index] != '\r' &&
                        text[index] != '\n' &&
                        char.IsWhiteSpace(text[index]))
                    {
                        index++;
                    }
                }

                if (index < text.Length &&
                    text[index] != delimiter &&
                    text[index] != '\r' &&
                    text[index] != '\n')
                {
                    return true;
                }

                if (index < text.Length && text[index] == delimiter)
                {
                    index++;
                    fieldStart = index;
                }

                continue;
            }

            index++;
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
