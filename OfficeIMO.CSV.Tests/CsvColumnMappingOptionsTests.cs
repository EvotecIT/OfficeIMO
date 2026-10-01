using System;
using System.Data;
using System.Data.Common;
using System.Globalization;
using System.Linq;
using System.Threading.Tasks;
using OfficeIMO.CSV;
using OfficeIMO.Data;
using Xunit;

namespace OfficeIMO.CSV.Tests;

public class CsvColumnMappingOptionsTests {
    public sealed class Item {
        public int Id { get; set; }
        public decimal Price { get; set; }
        public DateTime Date { get; set; }
        public string? Note { get; set; } = "initialized";
    }

    private const string Text = "record_id;Price;When\n7;12,50;20260927\n8;14,75;20260928\n";

    private static void Configure(RowMapper<Item> map) {
        var priceCulture = (CultureInfo)CultureInfo.InvariantCulture.Clone();
        priceCulture.NumberFormat.NumberDecimalSeparator = ",";
        string[] formats = { "yyyyMMdd" };
        map.FromColumns<int>(new[] { "Id", "record_id" }, (row, value) => { row.Id = value; return row; })
            .FromColumn<decimal>("Price", (row, value) => { row.Price = value; return row; },
                new RowMappingColumnOptions { Culture = priceCulture })
            .FromColumn<DateTime>("When", (row, value) => { row.Date = value; return row; },
                new RowMappingColumnOptions { DateTimeFormats = formats })
            .FromColumns<string?>(new[] { "Note", "Comment" }, (row, value) => { row.Note = value; return row; },
                new RowMappingColumnOptions { Optional = true });
        // Binding captures configuration; deferred and parallel readers cannot observe these edits.
        priceCulture.NumberFormat.NumberDecimalSeparator = ":";
        formats[0] = "ddMMyyyy";
    }

    [Theory]
    [InlineData(0)]
    [InlineData(1)]
    [InlineData(2)]
    [InlineData(3)]
    public void OptionalAliasesCultureAndFormatsApplyAcrossMappingPaths(int path) {
        var document = CsvDocument.Parse(Text, new CsvLoadOptions { Delimiter = ';' });
        Item[] items;
        if (path == 0) items = document.RowsAs<Item>(Configure).ToArray();
        else if (path == 1) {
            using var reader = document.CreateDataReader();
            items = reader.RowsAs<Item>(Configure).ToArray();
        } else if (path == 2) {
            using var reader = CsvDocument.OpenTextDataReader(Text, new CsvLoadOptions { Delimiter = ';' });
            items = reader.RowsAsParallel<Item>(Configure, new ParallelRowMappingOptions { MaxDegreeOfParallelism = 2, BatchSize = 1 }).ToArray();
        } else {
            var table = new DataTable();
            table.Columns.Add("record_id"); table.Columns.Add("Price"); table.Columns.Add("When");
            table.Rows.Add("7", "12,50", "20260927"); table.Rows.Add("8", "14,75", "20260928");
            using var reader = table.CreateDataReader();
            items = reader.RowsAsParallel<Item>(Configure, new ParallelRowMappingOptions { MaxDegreeOfParallelism = 2, BatchSize = 1 }).ToArray();
        }
        Assert.Equal(new[] { 7, 8 }, items.Select(row => row.Id));
        Assert.Equal(new[] { 12.50m, 14.75m }, items.Select(row => row.Price));
        Assert.Equal(new[] { new DateTime(2026, 9, 27), new DateTime(2026, 9, 28) }, items.Select(row => row.Date));
        Assert.All(items, row => Assert.Equal("initialized", row.Note));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void MissingOptionalOnlyBindingsStillEmitInitializedModels(bool parallel) {
        var document = CsvDocument.Parse("Id\n1\n2\n3\n");
        Action<RowMapper<Item>> configure = map => map.FromColumn<string>("Missing",
            (row, value) => { row.Note = value; return row; }, new RowMappingColumnOptions { Optional = true });
        using var reader = document.CreateDataReader();
        Item[] items = parallel
            ? reader.RowsAsParallel<Item>(configure, new ParallelRowMappingOptions { MaxDegreeOfParallelism = 2, BatchSize = 1 }).ToArray()
            : reader.RowsAs<Item>(configure).ToArray();
        Assert.Equal(3, items.Length);
        Assert.All(items, row => Assert.Equal("initialized", row.Note));
    }

    [Fact]
    public void OptionalPresentColumnsAssignAndInvalidValuesStillFail() {
        var document = CsvDocument.Parse("Id,Note\n1,present\n");
        Item row = Assert.Single(document.RowsAs<Item>(map => map.FromColumn<string>("Note",
            (item, value) => { item.Note = value; return item; }, new RowMappingColumnOptions { Optional = true })));
        Assert.Equal("present", row.Note);
        var invalid = CsvDocument.Parse("Id\ninvalid\n");
        Assert.Throws<DataMappingException>(() => invalid.RowsAs<Item>(map => map.FromColumn<int>("Id",
            (item, value) => { item.Id = value; return item; }, new RowMappingColumnOptions { Optional = true })).ToArray());
        Assert.Throws<DataMappingException>(() => document.RowsAs<Item>(map => map.FromColumn<string>("Absent",
            (item, value) => item)).ToArray());
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void PerColumnConvertersHandleDeclineNullAndRedactedErrors(bool parallel) {
        var document = CsvDocument.Parse("Id,Note\n7,clear\n8,keep\n",
            CsvProfiles.CreateLoadOptions(CsvProfile.Strict));
        Action<RowMapper<Item>> configure = map => {
            map.FromColumn<int>("Id", (row, value) => { row.Id = value; return row; }, new RowMappingColumnOptions {
                TypeConverter = (raw, type, culture) => raw.Equals("7") ? (true, 70) : (false, null)
            });
            map.FromColumn<string?>("Note", (row, value) => { row.Note = value; return row; }, new RowMappingColumnOptions {
                TypeConverter = (raw, type, culture) => raw.Equals("clear") ? (true, null) : (false, null)
            });
        };
        using var reader = document.CreateDataReader();
        Item[] items = parallel ? reader.RowsAsParallel<Item>(configure).ToArray() : reader.RowsAs<Item>(configure).ToArray();
        Assert.Equal(new[] { 70, 8 }, items.Select(row => row.Id));
        Assert.Null(items[0].Note);
        Assert.Equal("keep", items[1].Note);
        var error = Assert.Throws<DataMappingException>(() => document.RowsAs<Item>(map => map.FromColumn<int>("Id",
            (row, value) => row, new RowMappingColumnOptions {
                TypeConverter = (_, _, _) => throw new InvalidOperationException("secret-source")
            })).ToArray());
        Assert.DoesNotContain("secret-source", error.Message);
    }

#if NET8_0_OR_GREATER
    [Fact]
    public async Task AsyncMappingUsesSamePerColumnControls() {
        using var reader = CsvDocument.OpenTextDataReader(Text, new CsvLoadOptions { Delimiter = ';' });
        var rows = new System.Collections.Generic.List<Item>();
        await foreach (Item item in reader.RowsAsAsync<Item>(Configure)) rows.Add(item);
        Assert.Equal(new[] { 12.50m, 14.75m }, rows.Select(row => row.Price));
        Assert.All(rows, row => Assert.Equal("initialized", row.Note));
    }
#endif
}
