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

    [Fact]
    public void DefaultOptions_WriteIsoDatesAndShortestNumbers()
    {
        WithCulture("pl-PL", () =>
        {
            var table = new DataTable();
            table.Columns.Add("When", typeof(DateTime));
            table.Columns.Add("Offset", typeof(DateTimeOffset));
            table.Columns.Add("Number", typeof(double));
            table.Columns.Add("Binary", typeof(byte[]));
            table.Rows.Add(
                new DateTime(2024, 1, 2, 3, 4, 5).AddTicks(1200000),
                new DateTimeOffset(2024, 1, 2, 3, 4, 5, TimeSpan.FromHours(-7)),
                0.1 + 0.2,
                new byte[] { 1, 2, 3 });

            using var writer = new StringWriter();
            CsvDocument.WriteDataTable(writer, table);

            Assert.Equal("When,Offset,Number,Binary" + Environment.NewLine +
                "2024-01-02T03:04:05.12,2024-01-02T03:04:05-07:00,0.30000000000000004,AQID" + Environment.NewLine, writer.ToString());
        });
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
