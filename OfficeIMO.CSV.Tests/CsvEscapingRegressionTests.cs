using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using OfficeIMO.CSV;
using Xunit;

namespace OfficeIMO.CSV.Tests;

public class CsvEscapingRegressionTests {
    [Theory]
    [InlineData(",", CsvQuoteMode.AsNeeded)]
    [InlineData(";", CsvQuoteMode.AsNeeded)]
    [InlineData("||", CsvQuoteMode.AsNeeded)]
    [InlineData(",", CsvQuoteMode.Always)]
    [InlineData("||", CsvQuoteMode.Always)]
    [InlineData(",", CsvQuoteMode.Never)]
    public void TextAndObjectWritersPreserveEscapingAcrossLongSegments(string delimiter, CsvQuoteMode mode) {
        string[] values = Values().ToArray();
        var options = new CsvSaveOptions {
            DelimiterText = delimiter,
            IncludeHeader = false,
            NewLine = "\n",
            QuoteMode = mode,
            FormulaInjectionPolicy = CsvFormulaInjectionPolicy.Escape
        };
        var document = new CsvDocument().WithHeader("Notes");
        using var textOutput = new StringWriter();
        using var objectOutput = new StringWriter();
        using (var textWriter = new CsvRowWriter(textOutput, options, leaveOpen: true))
        using (var objectWriter = new CsvRowWriter(objectOutput, options, leaveOpen: true)) {
            foreach (string value in values) {
                document.AddRow(value);
                textWriter.WriteTextRow(new[] { "Notes" }, new[] { value });
                objectWriter.WriteRow(new[] { "Notes" }, new object?[] { value });
            }
        }
        string expected = string.Concat(values.Select(value => Quote(value, delimiter, mode) + "\n"));
        Assert.Equal(expected, document.ToString(options));
        Assert.Equal(expected, textOutput.ToString());
        Assert.Equal(expected, objectOutput.ToString());
    }

    private static IEnumerable<string> Values() {
        yield return "";
        yield return "\"\"\"\"";
        yield return new string('"', 4096);
        yield return string.Concat(Enumerable.Repeat("\"key\":\"value\",", 300));
        yield return "  =SUM(1;2), \"quoted\"";
        yield return "Łódź 😀,;||\r\n\t";
        foreach (int length in new[] { 0, 15, 16, 31, 32, 63, 64, 255, 256, 257, 511, 512, 513, 4096, 32768 }) {
            yield return new string('a', length) + "\"";
            yield return "\"" + new string('b', length) + "\"\"suffix";
            yield return string.Concat(Enumerable.Repeat("\"a", length / 2)) + "\"Łódź\ud83d\ude80";
            yield return new string('"', length);
        }
    }

    private static string Quote(string value, string delimiter, CsvQuoteMode mode) {
        // The corpus has one formula-leading value; its escaped text remains quoted
        // according to the same public mode as every other value.
        if (value.StartsWith("  =", StringComparison.Ordinal)) value = "'" + value;
        bool quoted = mode == CsvQuoteMode.Always || mode == CsvQuoteMode.AsNeeded &&
            (value.Contains(delimiter) || value.IndexOfAny(new[] { '"', '\r', '\n' }) >= 0);
        return quoted ? "\"" + value.Replace("\"", "\"\"") + "\"" : value;
    }
}
