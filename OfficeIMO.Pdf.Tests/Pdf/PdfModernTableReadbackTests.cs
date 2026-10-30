using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfModernTableReadbackTests {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void ModernNumericColumnsRetainTheirHeadersAndAdjacentBooleanValues(bool extraColumn) {
        string[][] rows = extraColumn ? new[] {
            new[] { "Region", "Revenue", "Active", "Delta" },
            new[] { "North branch", "1250", "True", "5" },
            new[] { "South", "980", "False", "2" },
            new[] { "West", "1430", "True", "3" }
        } : new[] {
            new[] { "Region", "Revenue", "Active" },
            new[] { "North branch", "1250", "True" },
            new[] { "South", "980", "False" },
            new[] { "West", "1430", "True" }
        };
        using var document = PdfDocument.Create();
        document.Content.H1("Quarterly results").Text("Revenue improved in the current quarter.").Table(rows);
        using var stream = new MemoryStream(document.ToBytes());
        using var loaded = PdfDocument.Load(stream);
        PdfLogicalTable table = Assert.Single(Assert.Single(loaded.Read().Pages).Tables);
        Assert.Equal(rows[0].Length, table.Columns.Count);
        Assert.Equal(rows.Length, table.Rows.Count);
        for (int row = 0; row < rows.Length; row++) Assert.Equal(rows[row], table.Rows[row]);
    }
}
