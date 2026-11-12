using System.Data;
using System.Globalization;
using System.Text;
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

    [Theory]
    [InlineData("yyyy-MM-dd, HH:mm", false, "n,a")]
    [InlineData("O", true, null)]
    [InlineData("", true, "NULL")]
    [InlineData(null, false, "\"none\"")]
    [InlineData("dddd dddd dddd dddd dddd dddd dddd dddd dddd dddd dddd dddd dddd yyyy", false, null)]
    public void WriteDataReader_FormattedValuesMatchRowWriting(string? format, bool useUtc, string? nullValue)
    {
        using var table = new DataTable();
        table.Columns.Add("Id", typeof(int));
        table.Columns.Add("Name", typeof(string));
        table.Columns.Add("Created", typeof(DateTime));
        table.Columns.Add("Offset", typeof(DateTimeOffset));
        table.Columns.Add("Any", typeof(object));
        table.Columns.Add("Amount", typeof(decimal));
        table.Rows.Add(1, "a,b", new DateTime(2026, 3, 4, 5, 6, 7, DateTimeKind.Local),
            new DateTimeOffset(2026, 1, 2, 3, 4, 5, TimeSpan.FromHours(2)), new DateTime(2025, 12, 31, 23, 59, 59, DateTimeKind.Utc), 1.5m);
        table.Rows.Add(DBNull.Value, DBNull.Value, DBNull.Value, DBNull.Value, DBNull.Value, DBNull.Value);
        table.Rows.Add(3, "plain", new DateTime(2024, 2, 29), new DateTimeOffset(2024, 2, 29, 0, 0, 0, TimeSpan.Zero), "text", -2m);
        var options = new CsvSaveOptions { NewLine = "\n", DateTimeFormat = format, UseUtc = useUtc, NullValue = nullValue };

        using var reader = table.CreateDataReader();
        using var typed = new StringWriter(CultureInfo.InvariantCulture);
        CsvDocument.WriteDataReader(typed, reader, options);

        using var rows = new StringWriter(CultureInfo.InvariantCulture);
        using (var writer = new CsvRowWriter(rows, options, leaveOpen: true))
        {
            string[] columns = table.Columns.Cast<DataColumn>().Select(column => column.ColumnName).ToArray();
            foreach (DataRow row in table.Rows) writer.WriteRow(columns, row.ItemArray);
        }

        Assert.Equal(rows.ToString(), typed.ToString());
    }

    public sealed class FormattedObjectRow
    {
        public int Id { get; set; }
        public string? Name { get; set; }
        public DateTime Created { get; set; }
        public DateTime? Due { get; set; }
        public DateTimeOffset Offset { get; set; }
        public object? Any { get; set; }
        public double Amount { get; set; }
        public int? Count { get; set; }
    }

    [Theory]
    [InlineData("yyyy-MM-dd, HH:mm", false, "n,a", "", false)]
    [InlineData("O", true, null, "", true)]
    [InlineData("", true, "NULL", "de-DE", false)]
    [InlineData(null, false, "\"none\"", "", true)]
    [InlineData("dd.MM.yyyy", false, null, "de-DE", true)]
    [InlineData("dddd dddd dddd dddd dddd dddd dddd dddd dddd dddd dddd dddd dddd yyyy", false, null, "", false)]
    public void WriteObjects_FormattedValuesMatchRowWriting(string? format, bool useUtc, string? nullValue, string cultureName, bool utf8)
    {
        var rows = new List<FormattedObjectRow>();
        for (int index = 0; index < 400; index++)
        {
            bool empty = index % 5 == 3;
            rows.Add(new FormattedObjectRow {
                Id = index,
                Name = empty ? null : index % 7 == 0 ? "a,\"b\"" : "name" + index,
                Created = new DateTime(2026, 3, 4, 5, 6, 7, DateTimeKind.Local).AddMinutes(index * 37),
                Due = empty ? null : new DateTime(2024, 2, 29, 0, 0, 0, DateTimeKind.Utc).AddHours(index),
                Offset = new DateTimeOffset(2026, 1, 2, 3, 4, 5, TimeSpan.FromHours(2)).AddDays(index),
                Any = empty ? null : index % 2 == 0 ? new DateTime(2025, 12, 31, 23, 59, 59, DateTimeKind.Utc) : "text",
                Amount = index * 1.25,
                Count = empty ? null : index
            });
        }
        var options = new CsvSaveOptions {
            NewLine = "\n", DateTimeFormat = format, UseUtc = useUtc, NullValue = nullValue,
            Culture = cultureName.Length == 0 ? CultureInfo.InvariantCulture : CultureInfo.GetCultureInfo(cultureName)
        };

        string projected;
        if (utf8)
        {
            using var stream = new MemoryStream();
            using (var streamWriter = new StreamWriter(stream, new UTF8Encoding(false), 1024, leaveOpen: true))
                CsvDocument.WriteObjects(streamWriter, rows, options);
            projected = Encoding.UTF8.GetString(stream.ToArray());
        }
        else
        {
            using var text = new StringWriter(CultureInfo.InvariantCulture);
            CsvDocument.WriteObjects(text, rows, options);
            projected = text.ToString();
        }

        using var expected = new StringWriter(CultureInfo.InvariantCulture);
        using (var writer = new CsvRowWriter(expected, options, leaveOpen: true))
        {
            string[] columns = { "Id", "Name", "Created", "Due", "Offset", "Any", "Amount", "Count" };
            foreach (var row in rows)
                writer.WriteRow(columns, new object?[] { row.Id, row.Name, row.Created, row.Due, row.Offset, row.Any, row.Amount, row.Count });
        }

        Assert.Equal(expected.ToString(), projected);
    }

    [Theory]
    [InlineData(',', "", false)]
    [InlineData(';', "", false)]
    [InlineData('\t', "de-DE", false)]
    [InlineData(',', "de-DE", false)]
    [InlineData('|', "fr-FR", false)]
    [InlineData(',', "", true)]
    [InlineData(';', "de-DE", true)]
    public void WriteDataReader_TypedRecordsMatchRowWriting(char delimiter, string cultureName, bool formatted)
    {
        string[] pieces = { "a", "Zż", " ", ",", ";", "\t", "|", "\"", "\r", "\n", "\r\n", "😀", "\uD800", "\uDC00", "1.5", "x" };
        var random = new Random(20261111 + delimiter + cultureName.Length + (formatted ? 7 : 0));
        string Text() => string.Concat(Enumerable.Range(0, random.Next(0, 12)).Select(_ => pieces[random.Next(pieces.Length)]));
        using var table = new DataTable();
        table.Columns.Add("Text", typeof(string));
        table.Columns.Add("Int", typeof(int));
        table.Columns.Add("Long", typeof(long));
        table.Columns.Add("Double", typeof(double));
        table.Columns.Add("Decimal", typeof(decimal));
        table.Columns.Add("Bool", typeof(bool));
        table.Columns.Add("Date", typeof(DateTime));
        table.Columns.Add("Offset", typeof(DateTimeOffset));
        table.Columns.Add("Guid", typeof(Guid));
        table.Columns.Add("Any", typeof(object));
        for (int row = 0; row < 3000; row++)
        {
            object?[] values = {
                Text(), random.Next(int.MinValue, int.MaxValue), ((long)random.Next() << 32) | (uint)random.Next(), random.NextDouble() * Math.Pow(10, random.Next(-5, 12)),
                random.Next(-10_000_000, 10_000_000) / 1000m, random.Next(2) == 0, new DateTime(2020, 1, 1).AddSeconds(random.Next()),
                new DateTimeOffset(2021, 2, 3, 4, 5, 6, TimeSpan.FromHours(random.Next(-12, 13))), new Guid(random.Next(), 1, 2, new byte[8]),
                (row % 4) switch { 0 => Text(), 1 => random.Next(), 2 => new DateTime(2022, 5, 6, 7, 8, 9), _ => new string('ż', row % 97 == 3 ? 9_000 : 3) }
            };
            for (int column = 0; column < values.Length; column++) if (random.Next(9) == 0) values[column] = DBNull.Value;
            table.Rows.Add(values);
        }
        var options = new CsvSaveOptions {
            Delimiter = delimiter, NewLine = "\r\n",
            Culture = cultureName.Length == 0 ? CultureInfo.InvariantCulture : CultureInfo.GetCultureInfo(cultureName),
            DateTimeFormat = formatted ? "yyyy-MM-dd, HH:mm" : null, UseUtc = formatted, NullValue = formatted ? "n,a" : null
        };

        using var rows = new StringWriter(CultureInfo.InvariantCulture);
        using (var writer = new CsvRowWriter(rows, options, leaveOpen: true))
        {
            string[] columns = table.Columns.Cast<DataColumn>().Select(column => column.ColumnName).ToArray();
            foreach (DataRow row in table.Rows) writer.WriteRow(columns, row.ItemArray);
        }
        using var text = new StringWriter(CultureInfo.InvariantCulture);
        using (var reader = table.CreateDataReader()) CsvDocument.WriteDataReader(text, reader, options);
        using var stream = new MemoryStream();
        using (var reader = table.CreateDataReader()) CsvDocument.WriteDataReader(stream, reader, options);

        Assert.Equal(rows.ToString(), text.ToString());
        Assert.Equal(new UTF8Encoding(false).GetBytes(rows.ToString()), stream.ToArray());
    }

    [Theory]
    [InlineData("Mixed", false)]
    [InlineData("Mixed", true)]
    [InlineData("Text", false)]
    [InlineData("Scalar", false)]
    public void WriteDataReader_InvariantRecordsMatchRowWriting(string layout, bool rejectTypedRead)
    {
        // A lone surrogate is invalid UTF-16; UTF-8 streams carry the replacement character the encoder gives it.
        string[] pieces = { "a", "Zż", " ", ",", ";", "\"", "\r", "\n", "\r\n", "😀", "\uD800", "1.5", "x", "plain" };
        var random = new Random(20261114);
        // Mostly short text, with some past the short-text escaping length.
        string Text() => string.Concat(Enumerable.Range(0, random.Next(0, random.Next(8) == 0 ? 40 : 10)).Select(_ => pieces[random.Next(pieces.Length)]));
        string[] headers = layout switch {
            "Text" => new[] { "Code", "Name", "Notes" },
            "Scalar" => new[] { "Int", "Long", "Double", "Decimal", "Bool", "Date" },
            _ => new[] { "Text", "Int", "Long", "Double", "Decimal", "Bool", "Date" }
        };
        object? Value(string header) => header switch {
            "Int" => random.Next(int.MinValue, int.MaxValue),
            "Long" => ((long)random.Next() << 32) | (uint)random.Next(),
            "Double" => random.NextDouble() * Math.Pow(10, random.Next(-8, 20)),
            "Decimal" => random.Next(-10_000_000, 10_000_000) / 1000m,
            "Bool" => random.Next(2) == 0,
            "Date" => new DateTime(2020, 1, 1).AddSeconds(random.Next()),
            _ => Text()
        };
        var rows = new object?[2500][];
        for (int row = 0; row < rows.Length; row++)
        {
            rows[row] = headers.Select(Value).ToArray();
            if (row > 0)
            {
                for (int column = 0; column < headers.Length; column++) if (random.Next(9) == 0) rows[row][column] = DBNull.Value;
            }
        }
        if (rejectTypedRead) rows[1700][1] = 5L;
        var options = new CsvSaveOptions { NewLine = "\r\n" };

        using var expected = new StringWriter(CultureInfo.InvariantCulture);
        using (var writer = new CsvRowWriter(expected, options, leaveOpen: true))
        {
            foreach (object?[] row in rows) writer.WriteRow(headers, row);
        }
        using var text = new StringWriter(CultureInfo.InvariantCulture);
        using (var reader = new ThrowingGetValuesDataReader(headers, rows)) CsvDocument.WriteDataReader(text, reader, options);
        using var stream = new MemoryStream();
        using (var reader = new ThrowingGetValuesDataReader(headers, rows)) CsvDocument.WriteDataReader(stream, reader, options, leaveOpen: true);
        using var table = new DataTable();
        for (int column = 0; column < headers.Length; column++) table.Columns.Add(headers[column], rows[0][column]!.GetType());
        foreach (object?[] row in rows) table.Rows.Add(row);
        using var tableStream = new MemoryStream();
        using (var reader = table.CreateDataReader()) CsvDocument.WriteDataReader(tableStream, reader, options, leaveOpen: true);

        Assert.Equal(expected.ToString(), text.ToString());
        Assert.Equal(new UTF8Encoding(false).GetBytes(expected.ToString()), stream.ToArray());
        Assert.Equal(new UTF8Encoding(false).GetBytes(expected.ToString()), tableStream.ToArray());
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void WriteDataReader_StreamKeepsOnlyCompletedRowsAfterFailure(bool cancel)
    {
        using var cancellation = new CancellationTokenSource();
        string largeValue = new string('x', 40_000);
        using var reader = new ThrowingGetValuesDataReader(
            new[] { "Name", "Value" },
            new[] {
                new object?[] { "Alpha", "One" },
                new object?[] { "Beta", cancel ? new CancelingCsvValue(cancellation, largeValue) : new ThrowingCsvValue() },
                new object?[] { "Gamma", "Two" }
            },
            supportGetValues: true);
        using var stream = new MemoryStream();

        Action write = () => CsvDocument.WriteDataReader(stream, reader, new CsvSaveOptions { NewLine = "\n" }, leaveOpen: true, cancellation.Token);
        if (cancel) Assert.Throws<OperationCanceledException>(write);
        else Assert.Throws<InvalidOperationException>(write);

        Assert.Equal(cancel ? "Name,Value\nAlpha,One\nBeta," + largeValue + "\n" : "Name,Value\nAlpha,One\n",
            Encoding.UTF8.GetString(stream.ToArray()));
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
    [InlineData(true, false, true)]
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
    [InlineData(true, false, false, true)]
    [InlineData(true, false, true, true)]
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

        Assert.Equal(alwaysQuoted ? $"\"Name\"{delimiter}\"Value\"\n\"Alpha\"{delimiter}\"One\"\n"
            : $"Name{delimiter}Value\nAlpha{delimiter}One\n", writer.ToString());
    }

    [Theory]
    [InlineData(false, false, false)]
    [InlineData(true, true, false)]
    [InlineData(true, false, false)]
    [InlineData(false, false, true)]
    [InlineData(true, false, true)]
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

        Assert.Equal(alwaysQuoted ? $"\"Name\"{delimiter}\"Value\"\n\"Alpha\"{delimiter}\"One\"\n\"Beta\"{delimiter}\"" + largeValue + "\"\n"
            : $"Name{delimiter}Value\nAlpha{delimiter}One\nBeta{delimiter}" + largeValue + "\n", writer.ToString());
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void WriteDataReader_ReusedWriterPreservesExistingTextAndCompletedRowsAfterFailure(bool textDelimiter)
    {
        using var writer = new StringWriter(CultureInfo.InvariantCulture);
        writer.Write("existing text\n");
        string delimiter = textDelimiter ? "||" : ",";
        using var csv = new CsvRowWriter(writer, new CsvSaveOptions { NewLine = "\n", DelimiterText = delimiter }, leaveOpen: true);
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

        Assert.Equal($"existing text\nName{delimiter}Value\nFirst{delimiter}Before\nAlpha{delimiter}One\nLast{delimiter}After\n", writer.ToString());
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
