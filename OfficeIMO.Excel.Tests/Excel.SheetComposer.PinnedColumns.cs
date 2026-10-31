using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests;

public partial class Excel {
    private sealed class PinnedColumnRow {
        public string Zeta { get; set; } = "z";
        public string Domain { get; set; } = "example.com";
        public string Alpha { get; set; } = "a";
        public int Score { get; set; } = 5;
    }

    [Fact]
    public void TableFrom_KeepsPinnedColumnsFirst() {
        using ExcelDocument document = ExcelDocument.Create(Path.Combine(_directoryWithFiles, "TableFromPinnedColumns.xlsx"));
        var composer = document.AddSheet("Data").Report();
        composer.AddTable(new[] { new PinnedColumnRow() }, title: null, configure: options => {
            options.PinnedFirst = new[] { nameof(PinnedColumnRow.Domain), nameof(PinnedColumnRow.Score) };
        });

        ExcelSheet sheet = document.Sheets[0];
        var headers = Enumerable.Range(1, 4).Select(column => sheet.TryGetCellText(1, column, out string text) ? text : string.Empty).ToArray();
        Assert.Equal("Domain", headers[0]);
        Assert.Equal("Score", headers[1]);
    }
}
