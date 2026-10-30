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
    [Fact]
    public void RecoveredFinalRowRetainsTextBeyondTheEarlierRowsOccupiedWidth() {
        string[][] rows = {
            new[] { "Region", "Revenue", "Active" },
            new[] { "North branch", "1250", "True" },
            new[] { "South", "980", "False" },
            new[] { "West", "1430", "False pending approval" }
        };
        using var document = PdfDocument.Create();
        document.Content.H1("Quarterly results").Text("Revenue improved in the current quarter.").Table(rows);
        using var stream = new MemoryStream(document.ToBytes());
        using var loaded = PdfDocument.Load(stream);
        PdfLogicalTable table = Assert.Single(Assert.Single(loaded.Read().Pages).Tables);
        Assert.Equal(rows.Length, table.Rows.Count);
        for (int row = 0; row < rows.Length; row++) Assert.Equal(rows[row], table.Rows[row]);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void TwoRowNumericTableIsRecoveredBesideAnUnrelatedTable(bool compactFirst) {
        string[][] compact = { new[] { "Phase 2026", "Phase 2027" }, new[] { "10", "20" } };
        string[][] regular = {
            new[] { "Region", "Revenue", "Active" },
            new[] { "North", "1250", "True" },
            new[] { "South", "980", "False" }
        };
        using var document = PdfDocument.Create();
        document.Content.Table(compactFirst ? compact : regular).Spacer(24).Table(compactFirst ? regular : compact);
        using var stream = new MemoryStream(document.ToBytes());
        using var loaded = PdfDocument.Load(stream);
        var tables = Assert.Single(loaded.Read().Pages).Tables;
        Assert.Equal(2, tables.Count);
        PdfLogicalTable recovered = Assert.Single(tables, table => table.Rows[0].SequenceEqual(compact[0]));
        Assert.Equal(compact[1], recovered.Rows[1]);
        PdfLogicalTable unrelated = Assert.Single(tables, table => table.Rows[0].SequenceEqual(regular[0]));
        Assert.Equal(regular.Length, unrelated.Rows.Count);
    }

}
