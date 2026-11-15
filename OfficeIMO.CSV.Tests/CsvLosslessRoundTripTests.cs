using System.Data;
using System.Globalization;
using System.Text;
using Xunit;

namespace OfficeIMO.CSV.Tests;

/// <summary>
/// Values written with the default save options read back unchanged through a typed schema, on every write route,
/// under a culture whose decimal separator is a comma.
/// </summary>
public sealed class CsvLosslessRoundTripTests
{
    private static readonly (string Name, Type Type, object Value)[] Columns =
    {
        ("Unspecified", typeof(DateTime), new DateTime(2024, 2, 29, 13, 14, 15, DateTimeKind.Unspecified).AddTicks(1234567)),
        ("Utc", typeof(DateTime), new DateTime(1999, 12, 31, 23, 59, 59, DateTimeKind.Utc).AddTicks(9999999)),
        ("Midnight", typeof(DateTime), new DateTime(2024, 1, 2)),
        ("Offset", typeof(DateTimeOffset), new DateTimeOffset(2024, 6, 1, 8, 30, 0, TimeSpan.FromMinutes(345)).AddTicks(1000001)),
        ("Duration", typeof(TimeSpan), new TimeSpan(1, 2, 3, 4).Add(TimeSpan.FromTicks(5678901))),
        ("Pi", typeof(double), Math.PI),
        ("DoubleMax", typeof(double), double.MaxValue),
        ("Sum", typeof(double), 0.1 + 0.2),
        ("SingleMax", typeof(float), float.MaxValue),
        ("Single", typeof(float), 0.1f),
        ("DecimalMax", typeof(decimal), decimal.MaxValue),
        ("Money", typeof(decimal), -1234567890.1234567890m),
        ("Binary", typeof(byte[]), new byte[] { 0, 1, 44, 34, 250, 255 }),
        ("Id", typeof(Guid), new Guid("0f8fad5b-d9cb-469f-a165-70867728950e")),
        ("Text", typeof(string), "a, \"quoted\"\r\nline")
    };

    public static TheoryData<string> WriteRoutes => new() { "DataReader", "DataTable", "Rows", "Document" };

    [Theory]
    [MemberData(nameof(WriteRoutes))]
    public void DefaultOptions_ReadBackEveryValueExactly_UnderPolishCulture(string route)
    {
        WithCulture("pl-PL", () =>
        {
            string csv = Write(route, CreateTable(includeNullRow: false), new CsvSaveOptions());

            foreach (List<object[]> rows in ReadBoth(csv, new CsvReaderOptions()))
            {
                object[] row = Assert.Single(rows);
                for (int ordinal = 0; ordinal < Columns.Length; ordinal++)
                {
                    Assert.Equal(Columns[ordinal].Value, row[ordinal]);
                }

                Assert.Equal(((DateTimeOffset)Columns[3].Value).Offset, ((DateTimeOffset)row[3]).Offset);
                Assert.Equal(DateTimeKind.Utc, ((DateTime)row[1]).Kind);
            }
        });
    }

    public static TheoryData<string> TextRoutes => new() { "TextWriter", "Utf8Stream" };

    /// <summary>
    /// The exact default text, through the UTF-16 writer and the UTF-8 stream writer that writes typed rows of
    /// dates and numbers as bytes: the fraction trimmed of trailing zeros, Z for UTC, the local offset the
    /// runtime's K format gives, four-digit years and every offset sign.
    /// </summary>
    [Theory]
    [MemberData(nameof(TextRoutes))]
    public void DefaultOptions_WriteIsoDatesAndShortestNumbers(string route)
    {
        WithCulture("pl-PL", () =>
        {
            var dates = new DataTable();
            dates.Columns.Add("When", typeof(DateTime)).DateTimeMode = DataSetDateTime.Unspecified;
            dates.Columns.Add("Utc", typeof(DateTime)).DateTimeMode = DataSetDateTime.Utc;
            dates.Columns.Add("Local", typeof(DateTime)).DateTimeMode = DataSetDateTime.Local;
            dates.Columns.Add("Number", typeof(double));
            dates.Rows.Add(new DateTime(2024, 1, 2, 3, 4, 5).AddTicks(1200000),
                new DateTime(1, 1, 1, 0, 0, 0, DateTimeKind.Utc).AddTicks(1),
                new DateTime(2026, 10, 25, 2, 30, 0, DateTimeKind.Local).AddTicks(1000000), 0.1 + 0.2);
            dates.Rows.Add(new DateTime(2024, 1, 2).AddTicks(10),
                new DateTime(9999, 12, 31, 23, 59, 59, DateTimeKind.Utc).AddTicks(9999999),
                new DateTime(2026, 3, 29, 1, 59, 59, DateTimeKind.Local), 1e20);
            var offsets = new DataTable();
            offsets.Columns.Add("Offset", typeof(DateTimeOffset));
            offsets.Columns.Add("Binary", typeof(byte[]));
            offsets.Rows.Add(new DateTimeOffset(2024, 1, 2, 3, 4, 5, TimeSpan.FromHours(-7)), new byte[] { 1, 2, 3 });
            offsets.Rows.Add(new DateTimeOffset(9999, 12, 31, 23, 59, 59, TimeSpan.FromHours(14)).AddTicks(9999999), new byte[0]);
            offsets.Rows.Add(new DateTimeOffset(5, 6, 7, 8, 9, 10, TimeSpan.FromMinutes(-839)).AddTicks(100000), new byte[] { 255 });
            offsets.Rows.Add(new DateTimeOffset(2024, 1, 2, 3, 4, 5, TimeSpan.Zero).AddTicks(1230), new byte[] { 0 });

            string newLine = Environment.NewLine;
            // The runtime's own round-trip pattern is the reference for the local offset, daylight saving included.
            string Local(int row) => ((DateTime)dates.Rows[row]["Local"]).ToString("yyyy-MM-dd'T'HH:mm:ss.FFFFFFFK", CultureInfo.InvariantCulture);
            Assert.Equal("When,Utc,Local,Number" + newLine +
                "2024-01-02T03:04:05.12,0001-01-01T00:00:00.0000001Z," + Local(0) + ",0.30000000000000004" + newLine +
                "2024-01-02T00:00:00.000001,9999-12-31T23:59:59.9999999Z," + Local(1) + ",1E+20" + newLine, WriteTable(route, dates));            Assert.Equal("Offset,Binary" + newLine +
                "2024-01-02T03:04:05-07:00,AQID" + newLine +
                "9999-12-31T23:59:59.9999999+14:00," + newLine +
                "0005-06-07T08:09:10.01-13:59,/w==" + newLine +
                "2024-01-02T03:04:05.000123+00:00,AA==" + newLine, WriteTable(route, offsets));
        });
    }

    private static string WriteTable(string route, DataTable table)
    {
        if (route == "TextWriter")
        {
            using var writer = new StringWriter();
            CsvDocument.WriteDataTable(writer, table);
            return writer.ToString();
        }

        using var stream = new MemoryStream();
        CsvDocument.WriteDataTable(stream, table, new CsvSaveOptions { Encoding = new UTF8Encoding(false) });
        return Encoding.UTF8.GetString(stream.ToArray());
    }
    [Theory]
    [MemberData(nameof(WriteRoutes))]
    public void NullValueMarker_KeepsNullDistinctFromEmptyText(string route)
    {
        var saveOptions = new CsvSaveOptions { NullValue = @"\N" };
        var readerOptions = new CsvReaderOptions { NullValue = @"\N" };
        DataTable table = CreateTable(includeNullRow: true);
        table.Rows[0]["Text"] = string.Empty;

        string csv = Write(route, table, saveOptions);

        foreach (List<object[]> rows in ReadBoth(csv, readerOptions))
        {
            Assert.Equal(string.Empty, rows[0][Columns.Length - 1]);
            Assert.All(rows[1], value => Assert.Equal(DBNull.Value, value));
        }
    }

    [Theory]
    [InlineData('\'', null, "'it''s'|'say \"hi\"'|back\\slash|'both\\''s'|'a,b'|'line\r\nbreak'|plain")]
    [InlineData('"', '\\', "it's|\"say \\\"hi\\\"\"|back\\slash|both\\'s|\"a,b\"|\"line\r\nbreak\"|plain")]
    [InlineData('\'', '\\', "'it\\'s'|'say \"hi\"'|back\\slash|'both\\\\\\'s'|'a,b'|'line\r\nbreak'|plain")]
    public void QuoteAndEscapeCharacters_RoundTripTextHoldingEveryQuote(char quote, char? escape, string expectedRecords)
    {
        string[] values = { "it's", "say \"hi\"", @"back\slash", @"both\'s", "a,b", "line\r\nbreak", "plain" };
        var document = CsvDocument.Create().WithHeader("Value");
        foreach (string value in values)
        {
            document.AddRow(value);
        }

        var saveOptions = new CsvSaveOptions { QuoteCharacter = quote, EscapeCharacter = escape, NewLine = "\n" };
        string csv = document.ToString(saveOptions);
        using var rowWriter = new StringWriter();
        using (var writer = new CsvRowWriter(rowWriter, saveOptions))
        {
            foreach (string value in values)
            {
                writer.WriteRow(new[] { "Value" }, new object?[] { value });
            }
        }

        Assert.Equal("Value\n" + expectedRecords.Replace('|', '\n') + "\n", csv);
        Assert.Equal(csv, rowWriter.ToString());

        var readerOptions = new CsvReaderOptions { QuoteCharacter = quote, EscapeCharacter = escape };
        foreach (List<object[]> rows in ReadBoth(csv, readerOptions))
        {
            Assert.Equal(values, rows.Select(row => (string)row[0]));
        }

        var loaded = CsvDocument.Parse(csv, new CsvLoadOptions { QuoteCharacter = quote, EscapeCharacter = escape });
        Assert.Equal(values, loaded.AsEnumerable().Select(row => row[0]?.ToString()));
    }

    [Fact]
    public void EscapeCharacter_ReadsMySqlStyleBackslashEscapes()
    {
        const string csv = "Name,Note\n\"O\\\"Brien\",\"a \\\\ b\"\n";

        using var reader = CsvDocument.OpenReader(new StringReader(csv), new CsvReaderOptions { EscapeCharacter = '\\' });

        Assert.True(reader.Read());
        Assert.Equal("O\"Brien", reader.GetString(0));
        Assert.Equal(@"a \ b", reader.GetString(1));
    }

    [Fact]
    public void EscapeCharacter_KeepsEscapesBeforeOtherCharactersAsText()
    {
        const string csv = "Path\n\"C:\\temp\\new\"\n";

        using var reader = CsvDocument.OpenReader(new StringReader(csv), new CsvReaderOptions { EscapeCharacter = '\\' });

        Assert.True(reader.Read());
        Assert.Equal(@"C:\temp\new", reader.GetString(0));
    }

    [Fact]
    public void CustomQuoting_StrictModeRejectsQuoteInsideUnquotedField()
    {
        var options = new CsvLoadOptions { EscapeCharacter = '\\', QuoteParsingMode = CsvQuoteParsingMode.Strict };

        Assert.Throws<CsvParseException>(() => CsvDocument.Parse("Name,Value\nab\"cd\",42\n", options));
    }

    [Fact]
    public void CustomQuoting_UnclosedCommentStopsAtTheContinuationLimit()
    {
        var options = new CsvLoadOptions { QuoteCharacter = '\'', SkipCommentRows = true, MaxCommentContinuationLines = 1 };

        var document = CsvDocument.Parse("#,'unfinished\nName,Value\nAlpha,1\n", options);

        Assert.Equal(new[] { "Name", "Value" }, document.Header);
        Assert.Equal("Alpha", document.AsEnumerable().Single()[0]);
    }

    [Fact]
    public async Task CustomQuoting_SaveAsyncWritesALastRecordWithoutLineBreak()
    {
        var document = CsvDocument.Create().WithHeader("Value").AddRow("it's");
        var options = new CsvSaveOptions { IncludeHeader = false, QuoteCharacter = '\'', NewLine = string.Empty };
        using var synchronous = new MemoryStream();
        using var asynchronous = new MemoryStream();

        document.Save(synchronous, options);
        await document.SaveAsync(asynchronous, options);

        Assert.Equal("'it''s'", Encoding.UTF8.GetString(synchronous.ToArray()));
        Assert.Equal(synchronous.ToArray(), asynchronous.ToArray());
    }

    [Fact]
    public void CustomQuoting_RejectedDialectLeavesTheExistingFile()
    {
        string path = Path.Combine(Path.GetTempPath(), "officeimo-dialect-" + Guid.NewGuid().ToString("N") + ".csv");
        File.WriteAllText(path, "keep");
        try
        {
            var document = CsvDocument.Create().WithHeader("Value").AddRow("x");

            Assert.Throws<ArgumentException>(() => document.Save(path, new CsvSaveOptions { QuoteCharacter = ',' }));
            Assert.Throws<ArgumentException>(() => CsvRowWriter.CreateFile(path, new CsvSaveOptions { QuoteCharacter = ',' }));
            Assert.Equal("keep", File.ReadAllText(path));
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void IsoDates_ReadAsGregorianUnderACultureWithAnotherCalendar()
    {
        var culture = new CultureInfo("th-TH");
        culture.DateTimeFormat.Calendar = new ThaiBuddhistCalendar();
        var schema = new CsvSchemaBuilder().Column("When").AsDateTime().Column("Offset").AsType(typeof(DateTimeOffset)).Done().Build();

        using var reader = CsvDocument.OpenReader(
            new StringReader("When,Offset\n2026-01-02T03:04:05.5,2026-01-02T03:04:05+01:00\n"),
            new CsvReaderOptions { Culture = culture, Schema = schema });

        Assert.True(reader.Read());
        Assert.Equal(new DateTime(2026, 1, 2, 3, 4, 5, 500), reader.GetDateTime(0));
        Assert.Equal(new DateTimeOffset(2026, 1, 2, 3, 4, 5, TimeSpan.FromHours(1)), reader.GetValue(1));
    }

    private static DataTable CreateTable(bool includeNullRow)
    {
        var table = new DataTable();
        foreach (var column in Columns)
        {
            DataColumn added = table.Columns.Add(column.Name, column.Type);
            if (column.Value is DateTime dateTime)
            {
                // Keep the kind of each value through the table.
                added.DateTimeMode = dateTime.Kind == DateTimeKind.Utc ? DataSetDateTime.Utc : DataSetDateTime.Unspecified;
            }
        }

        table.Rows.Add(Columns.Select(column => column.Value).ToArray());
        if (includeNullRow)
        {
            table.Rows.Add(Columns.Select(_ => (object)DBNull.Value).ToArray());
        }

        return table;
    }

    private static string Write(string route, DataTable table, CsvSaveOptions options)
    {
        using var writer = new StringWriter(CultureInfo.InvariantCulture);
        string[] names = table.Columns.Cast<DataColumn>().Select(column => column.ColumnName).ToArray();
        switch (route)
        {
            case "DataReader":
                using (var source = table.CreateDataReader())
                {
                    CsvDocument.WriteDataReader(writer, source, options);
                }

                break;
            case "DataTable":
                CsvDocument.WriteDataTable(writer, table, options);
                break;
            case "Rows":
                using (var rowWriter = new CsvRowWriter(writer, options, leaveOpen: true))
                {
                    foreach (DataRow row in table.Rows)
                    {
                        rowWriter.WriteRow(names, row.ItemArray.Select(value => value is DBNull ? null : value).ToArray());
                    }
                }

                break;
            case "Document":
                var document = CsvDocument.Create().WithHeader(names);
                foreach (DataRow row in table.Rows)
                {
                    document.AddRow(row.ItemArray.Select(value => value is DBNull ? null : value).ToArray());
                }

                return document.ToString(options);
            default:
                throw new ArgumentOutOfRangeException(nameof(route));
        }

        return writer.ToString();
    }

    /// <summary>Reads the text through the text reader and through the UTF-8 stream reader.</summary>
    private static IEnumerable<List<object[]>> ReadBoth(string csv, CsvReaderOptions options)
    {
        var schema = new CsvSchemaBuilder();
        CsvColumnBuilder? column = null;
        foreach (var definition in Columns)
        {
            column = (column?.Column(definition.Name) ?? schema.Column(definition.Name)).AsType(definition.Type);
        }

        if (csv.StartsWith("Value", StringComparison.Ordinal))
        {
            schema = new CsvSchemaBuilder();
            column = schema.Column("Value").AsString();
        }

        options.Schema = column!.Done().Build();
        using (var reader = CsvDocument.OpenReader(new StringReader(csv), options))
        {
            yield return Load(reader);
        }

        using var stream = new MemoryStream(Encoding.UTF8.GetBytes(csv));
        using (var reader = CsvDocument.OpenReader(stream, options))
        {
            yield return Load(reader);
        }

        static List<object[]> Load(IDataReader reader)
        {
            var rows = new List<object[]>();
            while (reader.Read())
            {
                var values = new object[reader.FieldCount];
                reader.GetValues(values);
                rows.Add(values);
            }

            return rows;
        }
    }

    private static void WithCulture(string name, Action action)
    {
        CultureInfo original = CultureInfo.CurrentCulture;
        CultureInfo.CurrentCulture = new CultureInfo(name);
        try
        {
            action();
        }
        finally
        {
            CultureInfo.CurrentCulture = original;
        }
    }
}
