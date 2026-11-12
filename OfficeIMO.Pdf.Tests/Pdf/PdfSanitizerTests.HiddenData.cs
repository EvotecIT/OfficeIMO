using OfficeIMO;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public partial class PdfSanitizerTests {
    [Fact]
    public void HiddenDataReportsPdfCategoriesAndDefaultRemovalKeepsOptInKinds() {
        using PdfDocument document = PdfDocument.Load(BuildBeforeSharingPdf());

        OfficeHiddenDataReport inspection = document.InspectHiddenData();

        Assert.Equal("PDF", inspection.Format);
        Assert.All(inspection.Findings, static finding => Assert.Equal(OfficeHiddenDataOutcome.Found, finding.Outcome));
        Assert.Contains(inspection.GetFindings(OfficeHiddenDataKind.DocumentProperties),
            static finding => finding.Location == "Info/Author" && finding.Value == "PRIVATE-AUTHOR");
        Assert.Contains(inspection.GetFindings(OfficeHiddenDataKind.DocumentProperties), static finding => finding.Location == "Catalog/Metadata");
        Assert.Equal("PRIVATE-COMMENT", Assert.Single(inspection.GetFindings(OfficeHiddenDataKind.Comments)).Value);
        Assert.Contains(inspection.GetFindings(OfficeHiddenDataKind.ActiveContent),
            static finding => finding.Description.Contains("JavaScript", StringComparison.Ordinal));
        Assert.Equal("payload.txt", Assert.Single(inspection.GetFindings(OfficeHiddenDataKind.Attachments)).Value);
        Assert.Equal("PRIVATE-BOOKMARK", Assert.Single(inspection.GetFindings(OfficeHiddenDataKind.Bookmarks)).Value);
        Assert.Equal("PRIVATE-LAYER", Assert.Single(inspection.GetFindings(OfficeHiddenDataKind.HiddenContent)).Value);

        OfficeHiddenDataResult removal = document.RemoveHiddenData();

        foreach (OfficeHiddenDataKind removed in new[] { OfficeHiddenDataKind.DocumentProperties, OfficeHiddenDataKind.Comments, OfficeHiddenDataKind.ActiveContent }) {
            Assert.NotEmpty(removal.Report.GetFindings(removed));
            Assert.All(removal.Report.GetFindings(removed), static finding => Assert.Equal(OfficeHiddenDataOutcome.Removed, finding.Outcome));
        }
        foreach (OfficeHiddenDataKind kept in new[] { OfficeHiddenDataKind.Attachments, OfficeHiddenDataKind.Bookmarks, OfficeHiddenDataKind.HiddenContent }) {
            OfficeHiddenDataFinding finding = Assert.Single(removal.Report.GetFindings(kept));
            Assert.Equal(OfficeHiddenDataOutcome.Retained, finding.Outcome);
            Assert.NotNull(finding.Note);
        }

        using PdfDocument cleaned = PdfDocument.Load(removal.ToArray());
        OfficeHiddenDataReport after = cleaned.InspectHiddenData();
        Assert.Empty(after.GetFindings(OfficeHiddenDataKind.DocumentProperties));
        Assert.Empty(after.GetFindings(OfficeHiddenDataKind.Comments));
        Assert.Empty(after.GetFindings(OfficeHiddenDataKind.ActiveContent));
        Assert.Single(after.GetFindings(OfficeHiddenDataKind.Attachments));
        Assert.Single(after.GetFindings(OfficeHiddenDataKind.Bookmarks));
        Assert.Single(after.GetFindings(OfficeHiddenDataKind.HiddenContent));
        Assert.Contains("VISIBLE-PAGE-CONTENT", cleaned.Read().Text, StringComparison.Ordinal);
        Assert.Equal(inspection.Findings.Count, document.InspectHiddenData().Findings.Count);
    }
}
