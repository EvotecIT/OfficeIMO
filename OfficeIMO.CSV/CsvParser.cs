#nullable enable

using System.Text;

namespace OfficeIMO.CSV;

internal static class CsvParser
{
    public static IEnumerable<string[]> Parse(TextReader reader, CsvLoadOptions options)
    {
        return ParseLineOrQuoted(reader, options);
    }

    private static IEnumerable<string[]> ParseLineOrQuoted(TextReader reader, CsvLoadOptions options)
    {
        var delimiter = options.Delimiter;
        var trim = options.TrimWhitespace;
        var allowEmpty = options.AllowEmptyLines;
        var lineNumber = 1;

        while (reader.ReadLine() is { } line)
        {
            if (ShouldSkipCommentLine(line, options))
            {
                lineNumber++;
                continue;
            }

            if (line.IndexOf('"') < 0)
            {
                var record = SplitUnquotedRecord(line, delimiter, trim);
                if (ShouldEmitRecord(record, allowEmpty))
                {
                    yield return record;
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
                    var next = reader.ReadLine();
                    if (next == null)
                    {
                        throw new CsvParseException("Unterminated quoted field.", lineNumber);
                    }

                    logicalRecord.Append('\n');
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
                yield return fields;
            }

            lineNumber++;
        }
    }

    private static string[] SplitUnquotedRecord(string line, char delimiter, bool trim)
    {
        var fieldCount = 1;
        for (var i = 0; i < line.Length; i++)
        {
            if (line[i] == delimiter)
            {
                fieldCount++;
            }
        }

        var fields = new string[fieldCount];
        var fieldIndex = 0;
        var start = 0;
        while (true)
        {
            var index = line.IndexOf(delimiter, start);
            if (index < 0)
            {
                fields[fieldIndex] = GetUnquotedField(line, start, line.Length - start, trim);
                return fields;
            }

            fields[fieldIndex] = GetUnquotedField(line, start, index - start, trim);
            fieldIndex++;
            start = index + 1;
        }
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
                inQuotes = true;
                fieldWasQuoted = true;
                continue;
            }

            if (c == delimiter)
            {
                AddField(parsedFields, buffer, trim, ref fieldWasQuoted);
                continue;
            }

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
            fields[fieldIndex++] = trim ? value.Trim() : value;
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

    private static IEnumerable<string[]> ParseCharacterByCharacter(TextReader reader, CsvLoadOptions options)
    {
        var delimiter = options.Delimiter;
        var trim = options.TrimWhitespace;
        var allowEmpty = options.AllowEmptyLines;

        var buffer = new StringBuilder();
        var fields = new List<string>();
        var lineNumber = 1;
        var inQuotes = false;
        var fieldWasQuoted = false;

        while (true)
        {
            var ch = reader.Read();
            var endOfFile = ch == -1;
            if (endOfFile)
            {
                if (inQuotes)
                {
                    throw new CsvParseException("Unterminated quoted field.", lineNumber);
                }

                AddField(fields, buffer, trim, ref fieldWasQuoted);
                if (ShouldEmitRecord(fields, allowEmpty))
                {
                    yield return fields.ToArray();
                }

                fields.Clear();
                yield break;
            }

            var c = (char)ch;

            if (inQuotes)
            {
                if (c == '"')
                {
                    var next = reader.Peek();
                    if (next == '"')
                    {
                        reader.Read();
                        buffer.Append('"');
                    }
                    else
                    {
                        inQuotes = false;
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
                inQuotes = true;
                fieldWasQuoted = true;
                continue;
            }

            if (c == delimiter)
            {
                AddField(fields, buffer, trim, ref fieldWasQuoted);
                continue;
            }

            if (c == '\n' || c == '\r')
            {
                if (c == '\r' && reader.Peek() == '\n')
                {
                    reader.Read();
                }

                AddField(fields, buffer, trim, ref fieldWasQuoted);
                if (ShouldEmitRecord(fields, allowEmpty))
                {
                    yield return fields.ToArray();
                }

                fields.Clear();
                lineNumber++;
                continue;
            }

            buffer.Append(c);
        }
    }

    private static bool ShouldSkipCommentLine(string line, CsvLoadOptions options)
    {
        if (!options.SkipCommentRows || line.Length == 0 || line[0] != options.CommentCharacter)
        {
            return false;
        }

        return !IsW3CFieldsLine(line, options);
    }

    private static bool IsW3CFieldsLine(string line, CsvLoadOptions options) =>
        options.RecognizeW3CFieldsHeader && line.StartsWith("#Fields:", StringComparison.OrdinalIgnoreCase);

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
        if (fields.Count == 0)
        {
            return allowEmpty;
        }

        if (!allowEmpty && AllFieldsEmpty(fields))
        {
            return false;
        }

        return true;
    }

    private static bool AllFieldsEmpty(IReadOnlyList<string> fields)
    {
        for (var i = 0; i < fields.Count; i++)
        {
            if (!string.IsNullOrEmpty(fields[i]))
            {
                return false;
            }
        }

        return true;
    }
}
