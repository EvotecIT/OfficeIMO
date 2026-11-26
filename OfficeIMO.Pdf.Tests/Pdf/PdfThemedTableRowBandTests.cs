using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfThemedTableRowBandTests {
    [Fact]
    public void RevenueTableKeepsItsLastRegionAndTotalInTheirColumns() {
        using var pdf = Open("word-themed-report.pdf");
        var table = Assert.Single(pdf.Read().Pages[0].Tables,
            table => table.Rows.Count > 0 && table.Rows[0][0] == "Region");

        Assert.Equal(7, table.Rows.Count);
        Assert.Equal(new[] { "Middle East & Africa", "$612,000", "$598,400", "-2.2%", "S. Haddad" }, table.Rows[5]);
        Assert.Equal(new[] { "Total", "$11,722,750", "$12,602,400", "+7.5%", "" }, table.Rows[6]);
    }

    [Fact]
    public void WrappedRiskCellsStayWithTheirOwnRowAndDoNotDuplicateTheHeader() {
        using var pdf = Open("word-themed-report.pdf");
        var table = Assert.Single(pdf.Read().Pages[0].Tables,
            table => table.Rows.Count > 0 && table.Rows[0][0] == "ID");

        Assert.Equal(5, table.Rows.Count);
        Assert.Equal(new[] { "ID", "Risk", "Likelihood", "Impact", "Mitigation" }, table.Rows[0]);
        Assert.Equal("Qualify second source by Q3; hold 6 weeks of safety stock", table.Rows[1][4]);
        Assert.Equal("Hedge 60% of forecast exposure on a rolling basis", table.Rows[2][4]);
        Assert.Equal("Freeze scope; weekly steering review", table.Rows[3][4]);
        Assert.Equal("Book 70% of peak capacity under contract by June", table.Rows[4][4]);
    }

    [Fact]
    public void NativeWrappedTableRetainsEverySignalAndItsCompleteEvidenceCell() {
        using var pdf = Open("native-wrapped-report.pdf");
        var table = Assert.Single(pdf.Read().Pages[0].Tables,
            table => table.Rows.Count > 0 && table.Rows[0][0] == "Signal");

        Assert.Equal(new[] { "Signal", "DMARC", "TLS", "DNS", "PDF" }, table.Rows.Select(row => row[0]));
        Assert.Equal("Certificate chain and protocol posture are ready for automated PSWriteOffice reports.", table.Rows[2][1]);
        Assert.Equal("Delegation and stale-record checks are summarized without overflowing table cells.", table.Rows[3][1]);
        Assert.Equal("Expand", table.Rows[4][2]);
    }

    [Fact]
    public void WrappedCandidateDoesNotDiscardThePrecedingFormIdentityRow() {
        using var pdf = Open("form-layout.pdf");
        PdfDocumentReadResult read = pdf.Read();
        // The form is tagged: each TD holds its label and value, and spanning cells keep their grid position.
        Assert.Collection(read.Pages[0].Tables,
            applicant => Assert.Equal(new[] { "1. Surname KOWALSKI", "", "2. Given names ANNA MARIA", "" }, applicant.Rows[0]),
            equipment => Assert.Equal("8. Start date 02/03/2026", equipment.Rows[0][0]));
        Assert.Contains(read.Headings, static heading => heading.Text == "Section B — Equipment");
    }

    private static PdfDocument Open(string name) => PdfDocument.Load(Path.Combine(AppContext.BaseDirectory,
        "Pdf", "Fixtures", "Interoperability", "TableRowBands", name));
}
