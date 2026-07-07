#nullable enable

using System.Text;

namespace OfficeIMO.CSV;

internal static partial class CsvParser
{
    private enum QuotedRecordParseResult
    {
        Complete,
        Invalid,
        Incomplete
    }

    private static bool TryParseQuotedRecord(string text, char delimiter, bool trim, bool strictQuotes, int lineNumber, out string[] fields)
    {
        fields = Array.Empty<string>();
        if (text.Length > 0 && text[0] == '"' && TryParseStrictQuotedRecord(text, delimiter, trim, out fields))
        {
            return true;
        }

        var standardResult = TryParseStandardQuotedRecord(text, delimiter, trim, out fields);
        if (standardResult == QuotedRecordParseResult.Complete)
        {
            return true;
        }

        if (standardResult == QuotedRecordParseResult.Incomplete)
        {
            return false;
        }

        if (strictQuotes)
        {
            throw new CsvParseException("Invalid quoted field.", lineNumber);
        }

        var buffer = new StringBuilder();
        var parsedFields = new List<string>(16);
        var inQuotes = false;
        var fieldWasQuoted = false;
        var afterClosingQuote = false;

        for (var i = 0; i < text.Length; i++)
        {
            var c = text[i];

            if (inQuotes)
            {
                if (c == '"')
                {
                    if (i + 1 < text.Length && text[i + 1] == '"')
                    {
                        i++;
                        buffer.Append('"');
                    }
                    else
                    {
                        inQuotes = false;
                        afterClosingQuote = true;
                    }
                }
                else
                {
                    buffer.Append(c);
                }

                continue;
            }

            if (c == '"')
            {
                if (afterClosingQuote)
                {
                    buffer.Append(c);
                    afterClosingQuote = false;
                    continue;
                }

                if (trim && IsWhitespaceOnly(buffer))
                {
                    buffer.Clear();
                }

                inQuotes = true;
                fieldWasQuoted = true;
                continue;
            }

            if (c == delimiter)
            {
                AddQuotedField(parsedFields, buffer, trim, ref fieldWasQuoted);
                afterClosingQuote = false;
                continue;
            }

            if (afterClosingQuote && char.IsWhiteSpace(c) && trim)
            {
                continue;
            }

            afterClosingQuote = false;
            buffer.Append(c);
        }

        if (inQuotes)
        {
            return false;
        }

        AddQuotedField(parsedFields, buffer, trim, ref fieldWasQuoted);
        fields = parsedFields.ToArray();
        return true;
    }

    private static bool TryParseQuotedRecord(string text, char delimiter, bool trim, bool strictQuotes, int lineNumber, List<string> fields)
    {
        fields.Clear();
        var standardResult = TryParseStandardQuotedRecord(text, delimiter, trim, fields);
        if (standardResult == QuotedRecordParseResult.Complete)
        {
            return true;
        }

        fields.Clear();
        if (standardResult == QuotedRecordParseResult.Incomplete)
        {
            return false;
        }

        if (strictQuotes)
        {
            throw new CsvParseException("Invalid quoted field.", lineNumber);
        }

        return TryParseFlexibleQuotedRecord(text, delimiter, trim, fields);
    }

    private static bool TryParseQuotedRecordLenient(string text, char delimiter, bool trim, out string[] fields) =>
        TryParseQuotedRecord(text, delimiter, trim, strictQuotes: false, lineNumber: 0, out fields);

    private static bool TryParseQuotedRecordLenient(string text, char delimiter, bool trim, List<string> fields) =>
        TryParseQuotedRecord(text, delimiter, trim, strictQuotes: false, lineNumber: 0, fields);

    private static bool TryParseFlexibleQuotedRecord(string text, char delimiter, bool trim, List<string> parsedFields)
    {
        var buffer = new StringBuilder();
        var inQuotes = false;
        var fieldWasQuoted = false;
        var afterClosingQuote = false;

        for (var i = 0; i < text.Length; i++)
        {
            var c = text[i];

            if (inQuotes)
            {
                if (c == '"')
                {
                    if (i + 1 < text.Length && text[i + 1] == '"')
                    {
                        i++;
                        buffer.Append('"');
                    }
                    else
                    {
                        inQuotes = false;
                        afterClosingQuote = true;
                    }
                }
                else
                {
                    buffer.Append(c);
                }

                continue;
            }

            if (c == '"')
            {
                if (afterClosingQuote)
                {
                    buffer.Append(c);
                    afterClosingQuote = false;
                    continue;
                }

                if (trim && IsWhitespaceOnly(buffer))
                {
                    buffer.Clear();
                }

                inQuotes = true;
                fieldWasQuoted = true;
                continue;
            }

            if (c == delimiter)
            {
                AddQuotedField(parsedFields, buffer, trim, ref fieldWasQuoted);
                afterClosingQuote = false;
                continue;
            }

            if (afterClosingQuote && char.IsWhiteSpace(c) && trim)
            {
                continue;
            }

            afterClosingQuote = false;
            buffer.Append(c);
        }

        if (inQuotes)
        {
            parsedFields.Clear();
            return false;
        }

        AddQuotedField(parsedFields, buffer, trim, ref fieldWasQuoted);
        return true;
    }

#if NET8_0_OR_GREATER
    private static bool TryParseQuotedRecordContinuations(
        CsvLineReader reader,
        Queue<CsvLine> pendingLines,
        string firstLine,
        string firstLineSeparator,
        char delimiter,
        bool trim,
        bool strictQuotes,
        List<string> parsedFields,
        ref int lineNumber)
    {
        parsedFields.Clear();
        var result = TryParseStandardQuotedRecordContinuations(
            reader,
            pendingLines,
            firstLine,
            firstLineSeparator,
            delimiter,
            trim,
            parsedFields,
            ref lineNumber);

        if (result == QuotedRecordParseResult.Invalid && strictQuotes)
        {
            throw new CsvParseException("Invalid quoted field.", lineNumber);
        }

        return result == QuotedRecordParseResult.Complete;
    }

    private static QuotedRecordParseResult TryParseStandardQuotedRecordContinuations(
        CsvLineReader reader,
        Queue<CsvLine> pendingLines,
        string firstLine,
        string firstLineSeparator,
        char delimiter,
        bool trim,
        List<string> parsedFields,
        ref int lineNumber)
    {
        var line = firstLine;
        var lineSeparator = firstLineSeparator;
        var index = 0;

        while (index < line.Length)
        {
            if (line[index] == delimiter)
            {
                parsedFields.Add(string.Empty);
                index++;
                if (index == line.Length)
                {
                    parsedFields.Add(string.Empty);
                }

                continue;
            }

            if (line[index] == '"')
            {
                var quotedResult = TryReadStandardQuotedFieldContinuations(
                    reader,
                    pendingLines,
                    ref line,
                    ref lineSeparator,
                    ref index,
                    trim,
                    delimiter,
                    out var quotedValue,
                    ref lineNumber);
                if (quotedResult != QuotedRecordParseResult.Complete)
                {
                    parsedFields.Clear();
                    return quotedResult;
                }

                parsedFields.Add(quotedValue);
            }
            else
            {
                var start = index;
                var specialIndex = line.AsSpan(index).IndexOfAny(delimiter, '"');
                if (specialIndex >= 0 && line[index + specialIndex] == '"')
                {
                    parsedFields.Clear();
                    return QuotedRecordParseResult.Invalid;
                }

                index = specialIndex >= 0 ? index + specialIndex : line.Length;

                parsedFields.Add(GetUnquotedField(line, start, index - start, trim));
            }

            if (index == line.Length)
            {
                return QuotedRecordParseResult.Complete;
            }

            if (line[index] != delimiter)
            {
                parsedFields.Clear();
                return QuotedRecordParseResult.Invalid;
            }

            index++;
            if (index == line.Length)
            {
                parsedFields.Add(string.Empty);
            }
        }

        return QuotedRecordParseResult.Complete;
    }

    private static QuotedRecordParseResult TryReadStandardQuotedFieldContinuations(
        CsvLineReader reader,
        Queue<CsvLine> pendingLines,
        ref string line,
        ref string lineSeparator,
        ref int index,
        bool trim,
        char delimiter,
        out string value,
        ref int lineNumber)
    {
        index++;
        var start = index;
        StringBuilder? builder = null;

        while (true)
        {
            while (index < line.Length)
            {
                var quoteIndex = line.IndexOf('"', index);
                if (quoteIndex < 0)
                {
                    break;
                }

                index = quoteIndex;
                if (index + 1 < line.Length && line[index + 1] == '"')
                {
                    builder ??= new StringBuilder();
                    builder.Append(line, start, index - start);
                    builder.Append('"');
                    index += 2;
                    start = index;
                    continue;
                }

                value = builder is null
                    ? line.Substring(start, index - start)
                    : AppendAndGetString(builder, line, start, index - start);
                index++;

                if (trim)
                {
                    while (index < line.Length && line[index] != delimiter && char.IsWhiteSpace(line[index]))
                    {
                        index++;
                    }
                }

                if (index < line.Length && line[index] != delimiter)
                {
                    value = string.Empty;
                    return QuotedRecordParseResult.Invalid;
                }

                return QuotedRecordParseResult.Complete;
            }

            var currentLine = line;
            var currentStart = start;
            var currentSeparator = lineSeparator;
            var next = ReadLineWithSeparator(reader, pendingLines, out lineSeparator);
            if (next == null)
            {
                value = string.Empty;
                return QuotedRecordParseResult.Incomplete;
            }

            line = next;
            index = 0;
            start = 0;
            lineNumber++;
            if (builder is null &&
                TryCompleteSingleContinuationQuotedField(
                    currentLine,
                    currentStart,
                    currentSeparator,
                    line,
                    trim,
                    delimiter,
                    ref index,
                    out value))
            {
                return QuotedRecordParseResult.Complete;
            }

            builder ??= new StringBuilder(currentLine.Length - currentStart + currentSeparator.Length + line.Length + 128);
            builder.Append(currentLine, currentStart, currentLine.Length - currentStart);
            builder.Append(currentSeparator);
        }
    }

    private static bool TryParseFlexibleQuotedRecordContinuations(
        CsvLineReader reader,
        Queue<CsvLine> pendingLines,
        string firstLine,
        string firstLineSeparator,
        char delimiter,
        bool trim,
        List<string> parsedFields,
        ref int lineNumber)
    {
        parsedFields.Clear();
        var buffer = new StringBuilder(firstLine.Length + firstLineSeparator.Length + 128);
        var inQuotes = false;
        var fieldWasQuoted = false;
        var afterClosingQuote = false;
        var line = firstLine;
        var lineSeparator = firstLineSeparator;

        while (true)
        {
            ParseQuotedLineSegment(line, delimiter, trim, parsedFields, buffer, ref inQuotes, ref fieldWasQuoted, ref afterClosingQuote);
            if (!inQuotes)
            {
                AddQuotedField(parsedFields, buffer, trim, ref fieldWasQuoted);
                return true;
            }

            buffer.Append(lineSeparator);
            var next = ReadLineWithSeparator(reader, pendingLines, out lineSeparator);
            if (next == null)
            {
                parsedFields.Clear();
                return false;
            }

            line = next;
            lineNumber++;
        }
    }

    private static void ParseQuotedLineSegment(
        string text,
        char delimiter,
        bool trim,
        List<string> parsedFields,
        StringBuilder buffer,
        ref bool inQuotes,
        ref bool fieldWasQuoted,
        ref bool afterClosingQuote)
    {
        for (var i = 0; i < text.Length; i++)
        {
            var c = text[i];

            if (inQuotes)
            {
                if (c == '"')
                {
                    if (i + 1 < text.Length && text[i + 1] == '"')
                    {
                        i++;
                        buffer.Append('"');
                    }
                    else
                    {
                        inQuotes = false;
                        afterClosingQuote = true;
                    }
                }
                else
                {
                    buffer.Append(c);
                }

                continue;
            }

            if (c == '"')
            {
                if (afterClosingQuote)
                {
                    buffer.Append(c);
                    afterClosingQuote = false;
                    continue;
                }

                if (trim && IsWhitespaceOnly(buffer))
                {
                    buffer.Clear();
                }

                inQuotes = true;
                fieldWasQuoted = true;
                continue;
            }

            if (c == delimiter)
            {
                AddQuotedField(parsedFields, buffer, trim, ref fieldWasQuoted);
                afterClosingQuote = false;
                continue;
            }

            if (afterClosingQuote && char.IsWhiteSpace(c) && trim)
            {
                continue;
            }

            afterClosingQuote = false;
            buffer.Append(c);
        }
    }
#endif

    private static QuotedRecordParseResult TryParseStandardQuotedRecord(string text, char delimiter, bool trim, out string[] fields)
    {
        var parsedFields = new List<string>(16);
        var result = TryParseStandardQuotedRecord(text, delimiter, trim, parsedFields);
        fields = result == QuotedRecordParseResult.Complete ? parsedFields.ToArray() : Array.Empty<string>();
        return result;
    }

    private static QuotedRecordParseResult TryParseStandardQuotedRecord(string text, char delimiter, bool trim, List<string> parsedFields)
    {
        var index = 0;

        while (index < text.Length)
        {
            if (text[index] == delimiter)
            {
                parsedFields.Add(string.Empty);
                index++;
                if (index == text.Length)
                {
                    parsedFields.Add(string.Empty);
                }

                continue;
            }

            if (text[index] == '"')
            {
                var quotedResult = TryReadStandardQuotedField(text, ref index, trim, delimiter, out var quotedValue);
                if (quotedResult != QuotedRecordParseResult.Complete)
                {
                    return quotedResult;
                }

                parsedFields.Add(quotedValue);
            }
            else
            {
                var start = index;
                var delimiterIndex = text.IndexOf(delimiter, index);
                var quoteIndex = text.IndexOf('"', index);
                if (quoteIndex >= 0 && (delimiterIndex < 0 || quoteIndex < delimiterIndex))
                {
                    return QuotedRecordParseResult.Invalid;
                }

                index = delimiterIndex >= 0 ? delimiterIndex : text.Length;

                parsedFields.Add(GetUnquotedField(text, start, index - start, trim));
            }

            if (index == text.Length)
            {
                return QuotedRecordParseResult.Complete;
            }

            if (text[index] != delimiter)
            {
                return QuotedRecordParseResult.Invalid;
            }

            index++;
            if (index == text.Length)
            {
                parsedFields.Add(string.Empty);
            }
        }

        return QuotedRecordParseResult.Complete;
    }

    private static QuotedRecordParseResult TryReadStandardQuotedField(string text, ref int index, bool trim, char delimiter, out string value)
    {
        index++;
        var start = index;
        StringBuilder? builder = null;

        while (index < text.Length)
        {
            var quoteIndex = text.IndexOf('"', index);
            if (quoteIndex < 0)
            {
                break;
            }

            index = quoteIndex;
            if (index + 1 < text.Length && text[index + 1] == '"')
            {
                builder ??= new StringBuilder();
                builder.Append(text, start, index - start);
                builder.Append('"');
                index += 2;
                start = index;
                continue;
            }

            value = builder is null
                ? text.Substring(start, index - start)
                : AppendAndGetString(builder, text, start, index - start);
            index++;

            if (trim)
            {
                while (index < text.Length && text[index] != delimiter && char.IsWhiteSpace(text[index]))
                {
                    index++;
                }
            }

            if (index < text.Length && text[index] != delimiter)
            {
                value = string.Empty;
                return QuotedRecordParseResult.Invalid;
            }

            return QuotedRecordParseResult.Complete;
        }

        value = string.Empty;
        return QuotedRecordParseResult.Incomplete;
    }

    private static string AppendAndGetString(StringBuilder builder, string text, int start, int count)
    {
        if (count > 0)
        {
            builder.Append(text, start, count);
        }

        return builder.ToString();
    }

#if NET8_0_OR_GREATER
    private static bool TryCompleteSingleContinuationQuotedField(
        string firstLine,
        int firstStart,
        string separator,
        string secondLine,
        bool trim,
        char delimiter,
        ref int index,
        out string value)
    {
        var quoteIndex = FindClosingQuote(secondLine, 0, out var secondEscapedQuoteCount);
        if (quoteIndex < 0)
        {
            value = string.Empty;
            return false;
        }

        var afterQuote = quoteIndex + 1;
        if (trim)
        {
            while (afterQuote < secondLine.Length &&
                   secondLine[afterQuote] != delimiter &&
                   char.IsWhiteSpace(secondLine[afterQuote]))
            {
                afterQuote++;
            }
        }

        if (afterQuote < secondLine.Length && secondLine[afterQuote] != delimiter)
        {
            value = string.Empty;
            return false;
        }

        var firstLength = firstLine.Length - firstStart;
        var firstEscapedQuoteCount = CountEscapedQuotes(firstLine, firstStart, firstLine.Length);
        value = string.Create(
            firstLength - firstEscapedQuoteCount + separator.Length + quoteIndex - secondEscapedQuoteCount,
            (firstLine, firstStart, firstLength, firstEscapedQuoteCount, separator, secondLine, quoteIndex, secondEscapedQuoteCount),
            static (destination, state) =>
            {
                var position = 0;
                position += CopyUnescapedQuotedSegment(
                    state.firstLine.AsSpan(state.firstStart, state.firstLength),
                    state.firstEscapedQuoteCount,
                    destination);
                state.separator.AsSpan().CopyTo(destination[position..]);
                position += state.separator.Length;
                CopyUnescapedQuotedSegment(
                    state.secondLine.AsSpan(0, state.quoteIndex),
                    state.secondEscapedQuoteCount,
                    destination[position..]);
            });

        index = afterQuote;
        return true;
    }

    private static int FindClosingQuote(string text, int start, out int escapedQuoteCount)
    {
        escapedQuoteCount = 0;
        var index = start;
        while (index < text.Length)
        {
            var quoteIndex = text.IndexOf('"', index);
            if (quoteIndex < 0)
            {
                return -1;
            }

            if (quoteIndex + 1 < text.Length && text[quoteIndex + 1] == '"')
            {
                escapedQuoteCount++;
                index = quoteIndex + 2;
                continue;
            }

            return quoteIndex;
        }

        return -1;
    }

    private static int CountEscapedQuotes(string text, int start, int end)
    {
        var count = 0;
        var index = start;
        while (index < end)
        {
            var quoteIndex = text.IndexOf('"', index, end - index);
            if (quoteIndex < 0 || quoteIndex + 1 >= end || text[quoteIndex + 1] != '"')
            {
                return count;
            }

            count++;
            index = quoteIndex + 2;
        }

        return count;
    }

    private static int CopyUnescapedQuotedSegment(ReadOnlySpan<char> source, int escapedQuoteCount, Span<char> destination)
    {
        if (escapedQuoteCount == 0)
        {
            source.CopyTo(destination);
            return source.Length;
        }

        var readIndex = 0;
        var writeIndex = 0;
        while (readIndex < source.Length)
        {
            var quoteOffset = source[readIndex..].IndexOf('"');
            if (quoteOffset < 0)
            {
                source[readIndex..].CopyTo(destination[writeIndex..]);
                writeIndex += source.Length - readIndex;
                break;
            }

            if (quoteOffset > 0)
            {
                source.Slice(readIndex, quoteOffset).CopyTo(destination[writeIndex..]);
                writeIndex += quoteOffset;
                readIndex += quoteOffset;
            }

            destination[writeIndex++] = '"';
            readIndex += 2;
        }

        return writeIndex;
    }
#endif

    private static bool IsWhitespaceOnly(StringBuilder buffer)
    {
        for (var i = 0; i < buffer.Length; i++)
        {
            if (!char.IsWhiteSpace(buffer[i]))
            {
                return false;
            }
        }

        return true;
    }

    private static bool TryParseStrictQuotedRecord(string text, char delimiter, bool trim, out string[] fields)
    {
        if (text.Length == 0)
        {
            fields = new[] { string.Empty };
            return true;
        }

        var fieldCount = 1;
        for (var i = 0; i < text.Length; i++)
        {
            if (text[i] == delimiter)
            {
                fieldCount++;
            }
        }

        fields = new string[fieldCount];

        var index = 0;
        var fieldIndex = 0;
        while (index < text.Length)
        {
            if (text[index] != '"')
            {
                fields = Array.Empty<string>();
                return false;
            }

            index++;
            var start = index;
            while (index < text.Length && text[index] != '"')
            {
                index++;
            }

            if (index >= text.Length)
            {
                fields = Array.Empty<string>();
                return false;
            }

            if (index + 1 < text.Length && text[index + 1] == '"')
            {
                fields = Array.Empty<string>();
                return false;
            }

            var value = text.Substring(start, index - start);
            fields[fieldIndex++] = value;
            index++;

            if (index == text.Length)
            {
                return fieldIndex == fields.Length;
            }

            if (text[index] != delimiter)
            {
                fields = Array.Empty<string>();
                return false;
            }

            index++;
            if (index == text.Length)
            {
                fields = Array.Empty<string>();
                return false;
            }
        }

        return fieldIndex == fields.Length;
    }

    private static void AddQuotedField(List<string> fields, StringBuilder buffer, bool trim, ref bool fieldWasQuoted)
    {
        var value = buffer.ToString();
        if (trim && !fieldWasQuoted)
        {
            value = value.Trim();
        }

        fields.Add(value);
        buffer.Clear();
        fieldWasQuoted = false;
    }
}
