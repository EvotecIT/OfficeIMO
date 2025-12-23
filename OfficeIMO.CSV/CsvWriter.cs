#nullable enable

using System.Globalization;
using System.Text;

namespace OfficeIMO.CSV;

internal static class CsvWriter
{
    public static void Write(TextWriter writer, CsvDocument document, CsvSaveOptions options)
    {
        var delimiter = options.Delimiter;
        var culture = options.Culture;
        var includeHeader = options.IncludeHeader;
        var newLine = options.NewLine;

        if (includeHeader && document.Header.Count > 0)
        {
            WriteRecord(writer, document.Header, delimiter, newLine);
        }

        foreach (var row in document.AsEnumerable())
        {
            WriteRecord(writer, row.Values, delimiter, newLine, culture);
        }
    }

    private static void WriteRecord(TextWriter writer, IReadOnlyList<string> header, char delimiter, string newLine)
    {
        WriteRecord(writer, header.Cast<object?>(), delimiter, newLine, CultureInfo.InvariantCulture);
    }

    private static void WriteRecord(TextWriter writer, IEnumerable<object?> values, char delimiter, string newLine, CultureInfo culture)
    {
        var first = true;
        foreach (var value in values)
        {
            if (!first)
            {
                writer.Write(delimiter);
            }
            else
            {
                first = false;
            }

            var text = FormatValue(value, culture);
            WriteEscaped(writer, text, delimiter);
        }

        writer.Write(newLine);
    }

    private static string FormatValue(object? value, CultureInfo culture)
    {
        if (value is null)
        {
            return string.Empty;
        }

        if (value is IFormattable formattable)
        {
            return formattable.ToString(null, culture);
        }

        return value.ToString() ?? string.Empty;
    }

    private static void WriteEscaped(TextWriter writer, string text, char delimiter)
    {
        var needsQuotes = text.IndexOfAny(new[] { '\"', '\n', '\r', delimiter }) >= 0;
        if (!needsQuotes)
        {
            writer.Write(text);
            return;
        }

        writer.Write('"');
        foreach (var ch in text)
        {
            if (ch == '\"')
            {
                writer.Write("\"\"");
            }
            else
            {
                writer.Write(ch);
            }
        }

        writer.Write('"');
    }
}
