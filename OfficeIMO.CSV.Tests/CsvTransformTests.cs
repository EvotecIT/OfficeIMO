using System.Linq;
using OfficeIMO.CSV;
using Xunit;

namespace OfficeIMO.CSV.Tests;

public class CsvTransformTests
{
    [Fact]
    public void AddColumn_Filter_Sort_Works()
    {
        var doc = new CsvDocument()
            .WithHeader("Name", "Age", "City");

        doc.AddRow("Przemek", 36, "Mikołów")
           .AddRow("Dominika", 30, "Mikołów")
           .AddRow("John", 50, "Warsaw");

        doc.AddColumn("Adult", row => row.AsInt32("Age") >= 18)
           .Filter(r => r.AsString("City") == "Mikołów")
           .SortBy("Age", descending: true);

        var rows = doc.AsEnumerable().ToList();
        Assert.Equal(2, rows.Count);
        Assert.True(rows[0].AsBoolean("Adult"));
        Assert.Equal("Przemek", rows[0].AsString("Name"));
        Assert.True(rows[0].AsInt32("Age") > rows[1].AsInt32("Age"));
    }
}
