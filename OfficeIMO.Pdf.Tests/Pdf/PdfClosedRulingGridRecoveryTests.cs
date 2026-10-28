using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfClosedRulingGridRecoveryTests {
    [Theory]
    [InlineData(0)]
    [InlineData(1)]
    [InlineData(2)]
    [InlineData(3)]
    public void ClosedRulesRecoverWrappedWorksheetRowsWithoutAPreexistingTextTable(int pageIndex) {
        using var pdf = Open("word-business-grids-rotated.pdf");
        PdfLogicalPage page = pdf.Read().Pages[pageIndex];
        PdfLogicalTable table = Assert.Single(page.Tables,
            table => table.Rows.Count > 0 && table.Rows[0][0] == "Worksheet");

        Assert.Equal(2, table.Columns.Count);
        Assert.Equal(new[] { "Worksheet", "Action Plan", "Workstream Plan", "Overlay Drafts", "Validation Plan", "Review Tracker" },
            table.Rows.Select(row => row[0]));
        Assert.Equal("Sequence recommendations, review owner/workstream, confirm acceptance criteria, and set the next planning step.", table.Rows[1][1]);
        Assert.Equal("Capture client response, target date, evidence notes, and validation status using stable RecommendationId/FindingId/RuleId values.", table.Rows[5][1]);
        Assert.Equal(new[] { "Worksheet", "Delivery Worksheet" }, page.Tables.Select(item => item.Rows[0][0]));
    }

    [Fact]
    public void ClosedRulesKeepEmptyResponseCellsAndTheirDistinctLabels() {
        using var pdf = Open("word-business-summary.pdf");
        PdfLogicalTable table = Assert.Single(pdf.Read().Pages[1].Tables,
            table => table.Rows.Count > 0 && table.Rows[0][0] == "Delivery Worksheet");

        Assert.Equal(new[] { "Delivery Worksheet", "Response" }, table.Rows[0]);
        Assert.Equal(6, table.Rows.Count);
        Assert.Equal(new[] { "Auditor notes", "" }, table.Rows[1]);
        Assert.Equal(new[] { "Client response", "" }, table.Rows[2]);
        Assert.Equal(new[] { "Decision", "Open / Accepted / Deferred / Risk accepted / Not applicable" }, table.Rows[3]);
        Assert.Equal(new[] { "Owner and target date", "" }, table.Rows[4]);
        Assert.Equal(new[] { "Evidence and validation notes", "" }, table.Rows[5]);
    }

    [Theory]
    [InlineData(0, true)]
    [InlineData(1, false)]
    [InlineData(2, true)]
    [InlineData(3, false)]
    public void NearbyFragmentAxesRecoverCompleteCellsOnlyWhenTheirIntervalsJoin(int pageIndex, bool complete) {
        using var pdf = Open("word-business-grids-fragmented.pdf");
        IReadOnlyList<PdfLogicalTable> tables = pdf.Read().Pages[pageIndex].Tables;
        bool recovered = tables.Any(table => table.Rows.Count == 6 && table.Columns.Count == 2 &&
            table.Rows[0][0] == "Worksheet" && table.Rows[5][0] == "Review Tracker");

        Assert.Equal(complete, recovered);
    }

    private static PdfDocument Open(string name) => PdfDocument.Load(Path.Combine(AppContext.BaseDirectory,
        "Pdf", "Fixtures", "Interoperability", "RuledVectorTables", name));
}
