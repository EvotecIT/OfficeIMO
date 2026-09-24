using System;
using System.Data;
using System.Globalization;
using System.IO;
using System.Linq;
using OfficeIMO.CSV;
using Xunit;

namespace OfficeIMO.CSV.Tests;

public class CsvDataReaderWriterRegressionTests
{
    [Theory]
    [InlineData('T', "\"True\"TAlpha\nFalseTBeta\n")]
    [InlineData('F', "TrueFAlpha\n\"False\"FBeta\n")]
    public void WriteDataReader_QuotesBooleanWhenDelimiterAppearsInLiteral(char delimiter, string expected)
    {
        var table = new DataTable();
        table.Columns.Add("Enabled", typeof(bool));
        table.Columns.Add("Name", typeof(string));
        table.Rows.Add(true, "Alpha");
        table.Rows.Add(false, "Beta");

        using var reader = table.CreateDataReader();
        using var writer = new StringWriter(CultureInfo.InvariantCulture);

        CsvDocument.WriteDataReader(
            writer,
            reader,
            new CsvSaveOptions { Delimiter = delimiter, IncludeHeader = false, NewLine = "\n" });

        Assert.Equal(expected, writer.ToString());
    }

    [Theory]
    [InlineData(CsvQuoteMode.AsNeeded)]
    [InlineData(CsvQuoteMode.Always)]
    public void WriteDataReader_MultiCharacterDelimiterPreservesTypedValuesAndFormulaPolicy(CsvQuoteMode quoteMode)
    {
        using var table = new DataTable();
        table.Columns.Add("Json", typeof(string));
        table.Columns.Add("Count", typeof(int));
        table.Columns.Add("Amount", typeof(decimal));
        table.Columns.Add("Created", typeof(DateTime)).DateTimeMode = DataSetDateTime.Utc;
        table.Columns.Add("Enabled", typeof(bool));
        table.Columns.Add("Missing", typeof(string));
        table.Rows.Add("{\"text\":\"A||B\"}", -12, 1.25m,
            new DateTime(2026, 9, 24, 10, 0, 0, DateTimeKind.Utc), true, DBNull.Value);

        using var reader = table.CreateDataReader();
        using var writer = new StringWriter(CultureInfo.InvariantCulture);
        CsvDocument.WriteDataReader(writer, reader, new CsvSaveOptions {
            DelimiterText = "||",
            NewLine = "\n",
            DateTimeFormat = "O",
            UseUtc = true,
            NullValue = "=empty",
            FormulaInjectionPolicy = CsvFormulaInjectionPolicy.Escape,
            QuoteMode = quoteMode
        });

        string output = writer.ToString();
        string[] fields = {
            "{\"text\":\"A||B\"}", "-12", "1.25", "2026-09-24T10:00:00.0000000Z", "True", "'=empty"
        };
        string expected = "Json||Count||Amount||Created||Enabled||Missing\n" +
            string.Join("||", fields.Select(field =>
                quoteMode == CsvQuoteMode.Always || field.Contains("||") || field.IndexOf('"') >= 0
                    ? "\"" + field.Replace("\"", "\"\"") + "\""
                    : field)) + "\n";
        if (quoteMode == CsvQuoteMode.Always) {
            expected = "\"Json\"||\"Count\"||\"Amount\"||\"Created\"||\"Enabled\"||\"Missing\"\n" +
                expected.Substring(expected.IndexOf('\n') + 1);
        }
        Assert.Equal(expected, output);
    }

    [Fact]
    public void WriteDataReader_MultiCharacterDelimiterPreservesLongCustomDateFormat()
    {
        string format = new string('y', 160);
        var value = new DateTime(2026, 9, 24, 10, 0, 0, DateTimeKind.Utc);
        using var table = new DataTable();
        table.Columns.Add("Created", typeof(DateTime)).DateTimeMode = DataSetDateTime.Utc;
        table.Rows.Add(value);
        using var reader = table.CreateDataReader();
        using var writer = new StringWriter(CultureInfo.InvariantCulture);

        CsvDocument.WriteDataReader(writer, reader, new CsvSaveOptions {
            DelimiterText = "||", NewLine = "\n", DateTimeFormat = format
        });

        Assert.Equal("Created\n" + value.ToString(format, CultureInfo.InvariantCulture) + "\n", writer.ToString());
    }

    [Fact]
    public void WriteDataReader_NeverQuoteLeavesSelectedNullFieldUnquoted()
    {
        using var table = new DataTable();
        table.Columns.Add("Missing", typeof(string));
        table.Columns.Add("Next", typeof(string));
        table.Rows.Add(DBNull.Value, "value");
        using var reader = table.CreateDataReader();
        using var writer = new StringWriter(CultureInfo.InvariantCulture);

        CsvDocument.WriteDataReader(writer, reader, new CsvSaveOptions {
            DelimiterText = "||", IncludeHeader = false, NewLine = "\n",
            QuoteMode = CsvQuoteMode.Never, QuoteFields = new[] { "Missing" }
        });

        Assert.Equal("||value\n", writer.ToString());
    }
}
