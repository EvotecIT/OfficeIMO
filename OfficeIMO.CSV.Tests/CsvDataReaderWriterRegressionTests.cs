using System;
using System.Data;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Threading;
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

    [Theory]
    [InlineData(false, false, false)]
    [InlineData(true, true, false)]
    [InlineData(true, false, false)]
    [InlineData(false, false, true)]
    public void WriteDataReader_CancellationKeepsCompletedRowsFromBatchedPaths(bool textDelimiter, bool formatted, bool alwaysQuoted)
    {
        using var cancellation = new CancellationTokenSource();
        using var reader = new ThrowingGetValuesDataReader(
            new[] { "Name" },
            new[] { new object?[] { "Alpha" }, new object?[] { "Beta" } },
            afterRead: index => { if (index == 1) cancellation.Cancel(); },
            supportGetValues: textDelimiter);
        using var writer = new StringWriter(CultureInfo.InvariantCulture);
        var options = new CsvSaveOptions {
            NewLine = "\n",
            DelimiterText = textDelimiter ? "||" : ",",
            DateTimeFormat = formatted ? "O" : null,
            QuoteMode = alwaysQuoted ? CsvQuoteMode.Always : CsvQuoteMode.AsNeeded
        };

        Assert.Throws<OperationCanceledException>(() =>
            CsvDocument.WriteDataReader(writer, reader, options, cancellation.Token));

        Assert.Equal(alwaysQuoted ? "\"Name\"\n\"Alpha\"\n" : "Name\nAlpha\n", writer.ToString());
    }

    [Theory]
    [InlineData(false, false, false, false)]
    [InlineData(false, false, true, false)]
    [InlineData(true, true, false, false)]
    [InlineData(true, true, true, false)]
    [InlineData(true, false, false, false)]
    [InlineData(true, false, true, false)]
    [InlineData(false, false, false, true)]
    [InlineData(false, false, true, true)]
    public void WriteDataReader_FormattingFailureDoesNotWritePartialBufferedRow(bool textDelimiter, bool formatted, bool supportGetValues, bool alwaysQuoted)
    {
        using var reader = new ThrowingGetValuesDataReader(
            new[] { "Name", "Value" },
            new[] {
                new object?[] { "Alpha", "One" },
                new object?[] { "Beta", new ThrowingCsvValue() }
            },
            supportGetValues: supportGetValues);
        using var writer = new StringWriter(CultureInfo.InvariantCulture);

        string delimiter = textDelimiter ? "||" : ",";
        Assert.Throws<InvalidOperationException>(() => CsvDocument.WriteDataReader(
            writer, reader,
            new CsvSaveOptions { DelimiterText = delimiter, DateTimeFormat = formatted ? "O" : null, NewLine = "\n",
                QuoteMode = alwaysQuoted ? CsvQuoteMode.Always : CsvQuoteMode.AsNeeded }));

        Assert.Equal(alwaysQuoted ? "\"Name\",\"Value\"\n\"Alpha\",\"One\"\n"
            : $"Name{delimiter}Value\nAlpha{delimiter}One\n", writer.ToString());
    }

    [Theory]
    [InlineData(false, false, false)]
    [InlineData(true, true, false)]
    [InlineData(true, false, false)]
    [InlineData(false, false, true)]
    public void WriteDataReader_CancellationAfterLargeRowKeepsCompletedRow(bool textDelimiter, bool formatted, bool alwaysQuoted)
    {
        using var cancellation = new CancellationTokenSource();
        string largeValue = new string('x', textDelimiter ? 9_000 : 40_000);
        using var reader = new ThrowingGetValuesDataReader(
            new[] { "Name", "Value" },
            new[] {
                new object?[] { "Alpha", "One" },
                new object?[] { "Beta", new CancelingCsvValue(cancellation, largeValue) }
            },
            supportGetValues: true);
        using var writer = new StringWriter(CultureInfo.InvariantCulture);

        string delimiter = textDelimiter ? "||" : ",";
        Assert.Throws<OperationCanceledException>(() => CsvDocument.WriteDataReader(
            writer, reader,
            new CsvSaveOptions { DelimiterText = delimiter, DateTimeFormat = formatted ? "O" : null, NewLine = "\n",
                QuoteMode = alwaysQuoted ? CsvQuoteMode.Always : CsvQuoteMode.AsNeeded },
            cancellation.Token));

        Assert.Equal(alwaysQuoted ? "\"Name\",\"Value\"\n\"Alpha\",\"One\"\n\"Beta\",\"" + largeValue + "\"\n"
            : $"Name{delimiter}Value\nAlpha{delimiter}One\nBeta{delimiter}" + largeValue + "\n", writer.ToString());
    }

    [Fact]
    public void WriteDataReader_ReusedWriterPreservesExistingTextAndCompletedRowsAfterFailure()
    {
        using var writer = new StringWriter(CultureInfo.InvariantCulture);
        writer.Write("existing text\n");
        using var csv = new CsvRowWriter(writer, new CsvSaveOptions { NewLine = "\n" }, leaveOpen: true);
        string[] headers = { "Name", "Value" };
        using var first = new ThrowingGetValuesDataReader(headers,
            new[] { new object?[] { "First", "Before" } });
        using var failed = new ThrowingGetValuesDataReader(headers,
            new[] { new object?[] { "Alpha", "One" }, new object?[] { "Beta", new ThrowingCsvValue() } });
        using var last = new ThrowingGetValuesDataReader(headers,
            new[] { new object?[] { "Last", "After" } });

        csv.WriteDataReader(first);
        Assert.Throws<InvalidOperationException>(() => csv.WriteDataReader(failed));
        csv.WriteDataReader(last);

        Assert.Equal("existing text\nName,Value\nFirst,Before\nAlpha,One\nLast,After\n", writer.ToString());
    }

    private sealed class CancelingCsvValue
    {
        private readonly CancellationTokenSource _cancellation;
        private readonly string _value;

        internal CancelingCsvValue(CancellationTokenSource cancellation, string value)
        {
            _cancellation = cancellation;
            _value = value;
        }

        public override string ToString()
        {
            _cancellation.Cancel();
            return _value;
        }
    }

    private sealed class ThrowingCsvValue
    {
        public override string ToString() => throw new InvalidOperationException("Value formatting failed.");
    }
}
