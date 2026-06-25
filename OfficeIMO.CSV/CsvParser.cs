#nullable enable

using System.Text;

namespace OfficeIMO.CSV;

internal static class CsvParser
{
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

        while (ReadLineWithSeparator(reader, out _) is { } line)
        {
            if (ShouldSkipCommentLine(line, options))
            {
                lineNumber++;
                continue;
            }

            if (TrySplitUnquotedRecord(line, delimiter, trim, out var record))
            {
                if (ShouldEmitRecord(record, allowEmpty))
                {
                    recordAction(record);
                }

                lineNumber++;
                continue;
            }

            string[] fields;
            if (!TryParseQuotedRecord(line, delimiter, trim, out fields))
            {
                var logicalRecord = new StringBuilder(line);
                while (true)
                {
                    var next = ReadLineWithSeparator(reader, out var separator);
                    if (next == null)
                    {
                        throw new CsvParseException("Unterminated quoted field.", lineNumber);
                    }

                    logicalRecord.Append(separator);
                    logicalRecord.Append(next);
                    lineNumber++;

                    if (TryParseQuotedRecord(logicalRecord.ToString(), delimiter, trim, out fields))
                    {
                        break;
                    }
                }
            }

            if (ShouldEmitRecord(fields, allowEmpty))
            {
                recordAction(fields);
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

        while (ReadLineWithSeparator(reader, out _) is { } line)
        {
            var startsWithCommentCharacter = IsRawCommentLine(line, options);
            if (ShouldSkipCommentLine(line, options))
            {
                lineNumber++;
                continue;
            }

            if (TrySplitUnquotedRecord(line, delimiter, trim, out var record))
            {
                if (ShouldEmitRecord(record, allowEmpty))
                {
                    yield return new CsvParsedRecord(record, startsWithCommentCharacter);
                }

                lineNumber++;
                continue;
            }

            string[] fields;
            if (!TryParseQuotedRecord(line, delimiter, trim, out fields))
            {
                var logicalRecord = new StringBuilder(line);
                while (true)
                {
                    var next = ReadLineWithSeparator(reader, out var separator);
                    if (next == null)
                    {
                        throw new CsvParseException("Unterminated quoted field.", lineNumber);
                    }

                    logicalRecord.Append(separator);
                    logicalRecord.Append(next);
                    lineNumber++;

                    if (TryParseQuotedRecord(logicalRecord.ToString(), delimiter, trim, out fields))
                    {
                        break;
                    }
                }
            }

            if (ShouldEmitRecord(fields, allowEmpty))
            {
                yield return new CsvParsedRecord(fields, startsWithCommentCharacter);
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

        while (ReadLineWithSeparator(reader, out _) is { } line)
        {
            var startsWithCommentCharacter = IsRawCommentLine(line, options);
            if (ShouldSkipCommentLine(line, options))
            {
                lineNumber++;
                continue;
            }

            if (TrySplitUnquotedRecord(line, delimiter, trim, reusableRecord))
            {
                if (ShouldEmitRecord(reusableRecord, allowEmpty))
                {
                    recordAction(new CsvParsedRecord(reusableRecord, startsWithCommentCharacter));
                }

                lineNumber++;
                continue;
            }

            string[] fields;
            if (!TryParseQuotedRecord(line, delimiter, trim, out fields))
            {
                var logicalRecord = new StringBuilder(line);
                while (true)
                {
                    var next = ReadLineWithSeparator(reader, out var separator);
                    if (next == null)
                    {
                        throw new CsvParseException("Unterminated quoted field.", lineNumber);
                    }

                    logicalRecord.Append(separator);
                    logicalRecord.Append(next);
                    lineNumber++;

                    if (TryParseQuotedRecord(logicalRecord.ToString(), delimiter, trim, out fields))
                    {
                        break;
                    }
                }
            }

            if (ShouldEmitRecord(fields, allowEmpty))
            {
                recordAction(new CsvParsedRecord(fields, startsWithCommentCharacter));
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

                if (IsWhitespaceOnly(buffer))
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

            if (afterClosingQuote && char.IsWhiteSpace(c))
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

    private static bool ShouldSkipCommentLine(string line, CsvLoadOptions options)
    {
        if (!options.SkipCommentRows || !IsRawCommentLine(line, options))
        {
            return false;
        }

        return !IsW3CFieldsLine(line, options);
    }

    private static bool IsRawCommentLine(string line, CsvLoadOptions options) =>
        line.Length > 0 && line[0] == options.CommentCharacter;

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
