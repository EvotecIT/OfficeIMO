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

    private static bool TryParseQuotedRecord(string text, char delimiter, bool trim, out string[] fields)
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

    private static bool TryParseQuotedRecord(string text, char delimiter, bool trim, List<string> fields)
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

        return TryParseFlexibleQuotedRecord(text, delimiter, trim, fields);
    }

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
                while (index < text.Length && text[index] != delimiter)
                {
                    if (text[index] == '"')
                    {
                        return QuotedRecordParseResult.Invalid;
                    }

                    index++;
                }

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
            if (text[index] != '"')
            {
                index++;
                continue;
            }

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
