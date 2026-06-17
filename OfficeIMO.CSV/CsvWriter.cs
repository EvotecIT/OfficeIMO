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
        var formulaInjectionPolicy = options.FormulaInjectionPolicy;

        if (includeHeader && document.Header.Count > 0)
        {
            WriteRecord(writer, document.Header, delimiter, newLine, CultureInfo.InvariantCulture, formulaInjectionPolicy);
        }

        foreach (var row in document.AsEnumerable())
        {
            WriteRecord(writer, row.Values, delimiter, newLine, culture, formulaInjectionPolicy);
        }
    }

    private static void WriteRecord(
        TextWriter writer,
        IEnumerable<object?> values,
        char delimiter,
        string newLine,
        CultureInfo culture,
        CsvFormulaInjectionPolicy formulaInjectionPolicy)
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

            var text = ApplyFormulaInjectionPolicy(FormatValue(value, culture), formulaInjectionPolicy);
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

    private static string ApplyFormulaInjectionPolicy(string text, CsvFormulaInjectionPolicy policy)
    {
        if (policy != CsvFormulaInjectionPolicy.Escape || !StartsWithFormulaTrigger(text))
        {
            return text;
        }

        return "'" + text;
    }

    private static bool StartsWithFormulaTrigger(string text)
    {
        if (string.IsNullOrEmpty(text))
        {
            return false;
        }

        var index = 0;
        while (index < text.Length && text[index] == ' ')
        {
            index++;
        }

        if (index >= text.Length)
        {
            return false;
        }

        return text[index] == '='
            || text[index] == '+'
            || text[index] == '-'
            || text[index] == '@'
            || text[index] == '\t'
            || text[index] == '\r'
            || text[index] == '\n';
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
