using OfficeIMO.Data;
using OfficeIMO.Excel;
using System.Globalization;
using Xunit;

namespace OfficeIMO.Tests;

public partial class Excel {
    public sealed class ColumnOptionSerialRow {
        public double Calendar { get; set; }
        public double Duration { get; set; }
        public string? Note { get; set; } = "initialized";
    }

    [Theory]
    [InlineData("xlsx", 1900, false)]
    [InlineData("xls", 1900, false)]
    [InlineData("xlsb", 1900, false)]
    [InlineData("xlsx", 1904, false)]
    [InlineData("xls", 1904, false)]
    [InlineData("xlsb", 1904, false)]
    [InlineData("xlsx", 1900, true)]
    [InlineData("xls", 1900, true)]
    [InlineData("xlsb", 1900, true)]
    [InlineData("xlsx", 1904, true)]
    [InlineData("xls", 1904, true)]
    [InlineData("xlsb", 1904, true)]
    public void Reader_ColumnOptions_RetainDateConverterInputAndNumericSerialFallback(string extension, int system, bool parallel) {
        string path = Path.Combine(AppContext.BaseDirectory, "Documents", "SpreadsheetDateSerialCorpus", $"serials-{system}.{extension}");
        using var reader = ExcelDocument.OpenDataReader(path, new ExcelReadOptions { SheetName = "Data" });
        Action<RowMapper<ColumnOptionSerialRow>> configure = map => {
            var options = new RowMappingColumnOptions {
                Culture = CultureInfo.InvariantCulture,
                TypeConverter = (raw, target, culture) => {
                    Assert.IsType<DateTime>(raw);
                    Assert.Equal(typeof(double), target);
                    return (false, null);
                }
            };
            map.FromColumns<double>(new[] { "Calendar", "Field1" }, (row, value) => { row.Calendar = value; return row; }, options)
                .FromColumn<double>("Field2", (row, value) => { row.Duration = value; return row; }, options)
                .FromColumn<string>("Missing", (row, value) => { row.Note = value; return row; }, new RowMappingColumnOptions { Optional = true });
        };
        ColumnOptionSerialRow[] rows = parallel
            ? reader.RowsAsParallel<ColumnOptionSerialRow>(configure, new ParallelRowMappingOptions { MaxDegreeOfParallelism = 2, BatchSize = 2 }).ToArray()
            : reader.RowsAs<ColumnOptionSerialRow>(configure).ToArray();
        double[] expected = { -0.5, 0, 1, 1.5, 59, 59.5, 60, 60.5, 61, 61.5, 1462, 45000, 1.5 };
        Assert.Equal(expected, rows.Select(row => row.Calendar));
        Assert.Equal(expected, rows.Select(row => row.Duration));
        Assert.All(rows, row => Assert.Equal("initialized", row.Note));
    }
}
