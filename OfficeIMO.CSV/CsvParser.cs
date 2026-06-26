#nullable enable

using System.Text;

namespace OfficeIMO.CSV;

internal static class CsvParser
{
    private readonly struct CsvLine
    {
        public CsvLine(string text, string separator)
        {
            Text = text;
            Separator = separator;
        }

        public string Text { get; }

        public string Separator { get; }
    }

    internal readonly struct CsvParsedRecord
    {
        public CsvParsedRecord(IReadOnlyList<string> values, bool startsWithCommentCharacter)
        {
            Values = values;
            StartsWithCommentCharacter = startsWithCommentCharacter;
        }

        public IReadOnlyList<string> Values { get; }

        public bool StartsWithCommentCharacter { get; }
    }

    public static IEnumerable<string[]> Parse(TextReader reader, CsvLoadOptions options)
    {
        return ParseLineOrQuoted(reader, options);
    }

    public static void ReadRecords(TextReader reader, CsvLoadOptions options, Action<string[]> recordAction)
    {
        if (recordAction == null)
        {
            throw new ArgumentNullException(nameof(recordAction));
        }

        ReadLineOrQuoted(reader, options, recordAction);
    }

    internal static IEnumerable<CsvParsedRecord> ParseWithMetadata(TextReader reader, CsvLoadOptions options)
    {
        return ParseLineOrQuotedWithMetadata(reader, options);
    }

    internal static void ReadRecordsWithMetadata(TextReader reader, CsvLoadOptions options, Action<CsvParsedRecord> recordAction)
    {
        if (recordAction == null)
        {
            throw new ArgumentNullException(nameof(recordAction));
        }

        foreach (var record in ParseLineOrQuotedWithMetadata(reader, options))
        {
            recordAction(record);
        }
    }

    public static void ReadRecordsReusable(TextReader reader, CsvLoadOptions options, Action<IReadOnlyList<string>> recordAction)
    {
        if (recordAction == null)
        {
            throw new ArgumentNullException(nameof(recordAction));
        }

        ReadLineOrQuotedReusable(reader, options, recordAction);
    }

    internal static void ReadRecordsReusableWithMetadata(TextReader reader, CsvLoadOptions options, Action<CsvParsedRecord> recordAction)
    {
        if (recordAction == null)
        {
            throw new ArgumentNullException(nameof(recordAction));
        }

        ReadLineOrQuotedReusableWithMetadata(reader, options, recordAction);
    }

    private static void ReadLineOrQuoted(TextReader reader, CsvLoadOptions options, Action<string[]> recordAction)
    {
        var delimiter = options.Delimiter;
        var trim = options.TrimWhitespace;
        var allowEmpty = options.AllowEmptyLines;
        var lineNumber = 1;
        var emittedRecordCount = 0;
        var pendingLines = new Queue<CsvLine>();

        while (ReadLineWithSeparator(reader, pendingLines, out var lineSeparator) is { } line)
        {
            var startsWithCommentCharacter = IsRawCommentLine(line, options);
            if (TrySkipCommentRecordBeforeParsing(reader, pendingLines, startsWithCommentCharacter, line, lineSeparator, options, emittedRecordCount, ref lineNumber))
            {
                lineNumber++;
                continue;
            }

            if (TrySplitUnquotedRecord(line, delimiter, trim, out var record))
            {
                if (!ShouldSkipCommentRecord(startsWithCommentCharacter, line, options, emittedRecordCount) &&
                    ShouldEmitRecord(record, allowEmpty))
                {
                    recordAction(record);
                    emittedRecordCount++;
                }

                lineNumber++;
                continue;
            }

            string[] fields;
            if (!TryParseQuotedRecord(line, delimiter, trim, out fields))
            {
                var logicalRecord = new StringBuilder(line);
                var pendingSeparator = lineSeparator;
                while (true)
                {
                    var next = ReadLineWithSeparator(reader, pendingLines, out var nextSeparator);
                    if (next == null)
                    {
                        throw new CsvParseException("Unterminated quoted field.", lineNumber);
                    }

                    logicalRecord.Append(pendingSeparator);
                    logicalRecord.Append(next);
                    lineNumber++;

                    if (TryParseQuotedRecord(logicalRecord.ToString(), delimiter, trim, out fields))
                    {
                        break;
                    }

                    pendingSeparator = nextSeparator;
                }
            }

            if (ShouldEmitRecord(fields, allowEmpty))
            {
                if (!ShouldSkipCommentRecord(startsWithCommentCharacter, line, options, emittedRecordCount))
                {
                    recordAction(fields);
                    emittedRecordCount++;
                }
            }

            lineNumber++;
        }
    }

    private static IEnumerable<string[]> ParseLineOrQuoted(TextReader reader, CsvLoadOptions options)
    {
        foreach (var record in ParseLineOrQuotedWithMetadata(reader, options))
        {
            if (record.Values is string[] fields)
            {
                yield return fields;
            }
        }
    }

    private static IEnumerable<CsvParsedRecord> ParseLineOrQuotedWithMetadata(TextReader reader, CsvLoadOptions options)
    {
        var delimiter = options.Delimiter;
        var trim = options.TrimWhitespace;
        var allowEmpty = options.AllowEmptyLines;
        var lineNumber = 1;
        var emittedRecordCount = 0;
        var pendingLines = new Queue<CsvLine>();

        while (ReadLineWithSeparator(reader, pendingLines, out var lineSeparator) is { } line)
        {
            var startsWithCommentCharacter = IsRawCommentLine(line, options);
            if (TrySkipCommentRecordBeforeParsing(reader, pendingLines, startsWithCommentCharacter, line, lineSeparator, options, emittedRecordCount, ref lineNumber))
            {
                lineNumber++;
                continue;
            }

            if (TrySplitUnquotedRecord(line, delimiter, trim, out var record))
            {
                if (!ShouldSkipCommentRecord(startsWithCommentCharacter, line, options, emittedRecordCount) &&
                    ShouldEmitRecord(record, allowEmpty))
                {
                    yield return new CsvParsedRecord(record, startsWithCommentCharacter);
                    emittedRecordCount++;
                }

                lineNumber++;
                continue;
            }

            string[] fields;
            if (!TryParseQuotedRecord(line, delimiter, trim, out fields))
            {
                var logicalRecord = new StringBuilder(line);
                var pendingSeparator = lineSeparator;
                while (true)
                {
                    var next = ReadLineWithSeparator(reader, pendingLines, out var nextSeparator);
                    if (next == null)
                    {
                        throw new CsvParseException("Unterminated quoted field.", lineNumber);
                    }

                    logicalRecord.Append(pendingSeparator);
                    logicalRecord.Append(next);
                    lineNumber++;

                    if (TryParseQuotedRecord(logicalRecord.ToString(), delimiter, trim, out fields))
                    {
                        break;
                    }

                    pendingSeparator = nextSeparator;
                }
            }

            if (ShouldEmitRecord(fields, allowEmpty))
            {
                if (!ShouldSkipCommentRecord(startsWithCommentCharacter, line, options, emittedRecordCount))
                {
                    yield return new CsvParsedRecord(fields, startsWithCommentCharacter);
                    emittedRecordCount++;
                }
            }

            lineNumber++;
        }
    }

    private static void ReadLineOrQuotedReusable(TextReader reader, CsvLoadOptions options, Action<IReadOnlyList<string>> recordAction)
    {
        ReadLineOrQuotedReusableWithMetadata(reader, options, record => recordAction(record.Values));
    }

    private static void ReadLineOrQuotedReusableWithMetadata(TextReader reader, CsvLoadOptions options, Action<CsvParsedRecord> recordAction)
    {
        var delimiter = options.Delimiter;
        var trim = options.TrimWhitespace;
        var allowEmpty = options.AllowEmptyLines;
        var lineNumber = 1;
        var reusableRecord = new List<string>(16);
        var emittedRecordCount = 0;
        var pendingLines = new Queue<CsvLine>();

        while (ReadLineWithSeparator(reader, pendingLines, out var lineSeparator) is { } line)
        {
            var startsWithCommentCharacter = IsRawCommentLine(line, options);
            if (TrySkipCommentRecordBeforeParsing(reader, pendingLines, startsWithCommentCharacter, line, lineSeparator, options, emittedRecordCount, ref lineNumber))
            {
                lineNumber++;
                continue;
            }

            if (TrySplitUnquotedRecord(line, delimiter, trim, reusableRecord))
            {
                if (!ShouldSkipCommentRecord(startsWithCommentCharacter, line, options, emittedRecordCount) &&
                    ShouldEmitRecord(reusableRecord, allowEmpty))
                {
                    recordAction(new CsvParsedRecord(reusableRecord, startsWithCommentCharacter));
                    emittedRecordCount++;
                }

                lineNumber++;
                continue;
            }

            string[] fields;
            if (!TryParseQuotedRecord(line, delimiter, trim, out fields))
            {
                var logicalRecord = new StringBuilder(line);
                var pendingSeparator = lineSeparator;
                while (true)
                {
                    var next = ReadLineWithSeparator(reader, pendingLines, out var nextSeparator);
                    if (next == null)
                    {
                        throw new CsvParseException("Unterminated quoted field.", lineNumber);
                    }

                    logicalRecord.Append(pendingSeparator);
                    logicalRecord.Append(next);
                    lineNumber++;

                    if (TryParseQuotedRecord(logicalRecord.ToString(), delimiter, trim, out fields))
                    {
                        break;
                    }

                    pendingSeparator = nextSeparator;
                }
            }

            if (ShouldEmitRecord(fields, allowEmpty))
            {
                if (!ShouldSkipCommentRecord(startsWithCommentCharacter, line, options, emittedRecordCount))
                {
                    recordAction(new CsvParsedRecord(fields, startsWithCommentCharacter));
                    emittedRecordCount++;
                }
            }

            lineNumber++;
        }
    }

    private static bool TrySplitUnquotedRecord(string line, char delimiter, bool trim, out string[] fields)
    {
        var fieldCount = 1;
        for (var i = 0; i < line.Length; i++)
        {
            var value = line[i];
            if (value == '"')
            {
                fields = Array.Empty<string>();
                return false;
            }

            if (value == delimiter)
            {
                fieldCount++;
            }
        }

        fields = new string[fieldCount];
        var fieldIndex = 0;
        var start = 0;
        for (var i = 0; i < line.Length; i++)
        {
            if (line[i] != delimiter)
            {
                continue;
            }

            fields[fieldIndex] = GetUnquotedField(line, start, i - start, trim);
            fieldIndex++;
            start = i + 1;
        }

        fields[fieldIndex] = GetUnquotedField(line, start, line.Length - start, trim);
        return true;
    }

    private static bool TrySplitUnquotedRecord(string line, char delimiter, bool trim, List<string> fields)
    {
        fields.Clear();
        var start = 0;
        for (var i = 0; i < line.Length; i++)
        {
            var value = line[i];
            if (value == '"')
            {
                fields.Clear();
                return false;
            }

            if (value != delimiter)
            {
                continue;
            }

            fields.Add(GetUnquotedField(line, start, i - start, trim));
            start = i + 1;
        }

        fields.Add(GetUnquotedField(line, start, line.Length - start, trim));
        return true;
    }

    private static string GetUnquotedField(string line, int start, int length, bool trim)
    {
        if (length == 0)
        {
            return string.Empty;
        }

        if (!trim)
        {
            return line.Substring(start, length);
        }

        var end = start + length - 1;
        while (start <= end && char.IsWhiteSpace(line[start]))
        {
            start++;
        }

        while (end >= start && char.IsWhiteSpace(line[end]))
        {
            end--;
        }

        if (end < start)
        {
            return string.Empty;
        }

        return line.Substring(start, end - start + 1);
    }

    private static bool TryParseQuotedRecord(string text, char delimiter, bool trim, out string[] fields)
    {
        fields = Array.Empty<string>();
        if (text.Length > 0 && text[0] == '"' && TryParseStrictQuotedRecord(text, delimiter, trim, out fields))
        {
            return true;
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
                AddField(parsedFields, buffer, trim, ref fieldWasQuoted);
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

        AddField(parsedFields, buffer, trim, ref fieldWasQuoted);
        fields = parsedFields.ToArray();
        return true;
    }

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

    private static bool ShouldSkipCommentRecord(bool startsWithCommentCharacter, string firstLine, CsvLoadOptions options, int emittedRecordCount)
    {
        if (!options.SkipCommentRows || !startsWithCommentCharacter)
        {
            return false;
        }

        return !CanReadW3CFieldsHeader(options, emittedRecordCount) || !IsW3CFieldsLine(firstLine, options);
    }

    private static bool TrySkipCommentRecordBeforeParsing(TextReader reader, Queue<CsvLine> pendingLines, bool startsWithCommentCharacter, string firstLine, string firstLineSeparator, CsvLoadOptions options, int emittedRecordCount, ref int lineNumber)
    {
        if (!ShouldSkipCommentRecordBeforeParsing(startsWithCommentCharacter, firstLine, options, emittedRecordCount))
        {
            return false;
        }

        if (TryParseQuotedRecord(firstLine, options.Delimiter, options.TrimWhitespace, out _))
        {
            return true;
        }

        var logicalRecord = new StringBuilder(firstLine);
        var pendingSeparator = firstLineSeparator;
        while (true)
        {
            var next = ReadLineWithSeparator(reader, pendingLines, out var nextSeparator);
            if (next == null)
            {
                return true;
            }

            var candidate = string.Concat(logicalRecord.ToString(), pendingSeparator, next);
            if (TryParseQuotedRecord(candidate, options.Delimiter, options.TrimWhitespace, out _))
            {
                lineNumber++;
                return true;
            }

            pendingLines.Enqueue(new CsvLine(next, nextSeparator));
            return true;
        }
    }

    private static bool ShouldSkipCommentRecordBeforeParsing(bool startsWithCommentCharacter, string firstLine, CsvLoadOptions options, int emittedRecordCount)
    {
        if (!startsWithCommentCharacter)
        {
            return false;
        }

        var canReadW3CFieldsHeader = CanReadW3CFieldsHeader(options, emittedRecordCount) && IsW3CFieldsLine(firstLine, options);
        if (canReadW3CFieldsHeader)
        {
            return false;
        }

        return options.SkipCommentRows ||
            (options.HasHeaderRow &&
                options.Header is null &&
                options.SkipCommentRowsBeforeHeader &&
                emittedRecordCount <= GetParserInitialRecordsToSkip(options));
    }

    private static bool IsRawCommentLine(string line, CsvLoadOptions options) =>
        line.Length > 0 && line[0] == options.CommentCharacter;

    private static bool CanReadW3CFieldsHeader(CsvLoadOptions options, int emittedRecordCount) =>
        emittedRecordCount <= GetParserInitialRecordsToSkip(options) &&
        options.HasHeaderRow &&
        options.Header is null &&
        options.RecognizeW3CFieldsHeader;

    private static int GetParserInitialRecordsToSkip(CsvLoadOptions options)
    {
        if (options.SkipInitialRecords < 0)
        {
            throw new ArgumentOutOfRangeException(nameof(options), "SkipInitialRecords cannot be negative.");
        }

        return options.SkipInitialRecords;
    }

    private static bool IsW3CFieldsLine(string line, CsvLoadOptions options) =>
        options.RecognizeW3CFieldsHeader && line.StartsWith("#Fields:", StringComparison.OrdinalIgnoreCase);

    private static string? ReadLineWithSeparator(TextReader reader, out string separator)
    {
        separator = string.Empty;
        var builder = new StringBuilder();
        while (true)
        {
            var value = reader.Read();
            if (value < 0)
            {
                return builder.Length == 0 ? null : builder.ToString();
            }

            var ch = (char)value;
            if (ch == '\r')
            {
                if (reader.Peek() == '\n')
                {
                    reader.Read();
                    separator = "\r\n";
                }
                else
                {
                    separator = "\r";
                }

                return builder.ToString();
            }

            if (ch == '\n')
            {
                separator = "\n";
                return builder.ToString();
            }

            builder.Append(ch);
        }
    }

    private static string? ReadLineWithSeparator(TextReader reader, Queue<CsvLine> pendingLines, out string separator)
    {
        if (pendingLines.Count > 0)
        {
            var pending = pendingLines.Dequeue();
            separator = pending.Separator;
            return pending.Text;
        }

        return ReadLineWithSeparator(reader, out separator);
    }

    private static void AddField(List<string> fields, StringBuilder buffer, bool trim, ref bool fieldWasQuoted)
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

    private static bool ShouldEmitRecord(IReadOnlyList<string> fields, bool allowEmpty)
    {
        return fields.Count != 0 &&
            (allowEmpty || fields.Count != 1 || fields[0].Length != 0);
    }
}
