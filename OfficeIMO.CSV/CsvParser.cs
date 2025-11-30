#nullable enable

using System.Text;

namespace OfficeIMO.CSV;

internal static class CsvParser
{
    public static IEnumerable<string[]> Parse(TextReader reader, CsvLoadOptions options)
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
                foreach (var record in EmitRecord(fields, allowEmpty))
                {
                    yield return record;
                }

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
                foreach (var record in EmitRecord(fields, allowEmpty))
                {
                    yield return record;
                }

                lineNumber++;
                continue;
            }

            buffer.Append(c);
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

    private static IEnumerable<string[]> EmitRecord(List<string> fields, bool allowEmpty)
    {
        if (fields.Count == 0)
        {
            if (allowEmpty)
            {
                yield return Array.Empty<string>();
            }

            yield break;
        }

        if (!allowEmpty && fields.All(string.IsNullOrEmpty))
        {
            fields.Clear();
            yield break;
        }

        var record = fields.ToArray();
        fields.Clear();
        yield return record;
    }
}
