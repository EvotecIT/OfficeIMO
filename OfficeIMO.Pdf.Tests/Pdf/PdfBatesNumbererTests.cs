using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public class PdfBatesNumbererTests {
    [Fact]
    public void Apply_NumbersSelectedPagesContinuouslyAcrossDocuments() {
        byte[] first = PdfProductionWorkflowTestSupport.CreatePdf("First one", "First two");
        byte[] second = PdfProductionWorkflowTestSupport.CreatePdf("Second one", "Second two");
        var secondInput = new PdfBatesDocument(second, "second.pdf") {
            TargetPages = PdfPageSelector.Parse("2")
        };

        PdfBatesBatchResult result = PdfBatesNumberer.Apply(
            new[] { new PdfBatesDocument(first, "first.pdf"), secondInput },
            new PdfBatesNumberingOptions {
                StartNumber = 42,
                Prefix = "CASE-",
                Suffix = "-EV",
                MinimumDigits = 4,
                Position = PdfBatesPosition.TopCenter
            });

        Assert.Equal(new long[] { 42, 43, 44 }, result.Assignments.Select(static assignment => assignment.Number));
        Assert.Equal(new[] { "CASE-0042-EV", "CASE-0043-EV", "CASE-0044-EV" }, result.Assignments.Select(static assignment => assignment.Text));
        Assert.Equal(45, result.NextNumber);
        Assert.All(result.Documents, static document => Assert.True(document.Preservation.IsPreserved));

        string[] firstPages = PdfProductionWorkflowTestSupport.ReadPageTexts(result.Documents[0].ToBytes());
        string[] secondPages = PdfProductionWorkflowTestSupport.ReadPageTexts(result.Documents[1].ToBytes());
        Assert.Contains("CASE-0042-EV", firstPages[0], StringComparison.Ordinal);
        Assert.Contains("CASE-0043-EV", firstPages[1], StringComparison.Ordinal);
        Assert.DoesNotContain("CASE-", secondPages[0], StringComparison.Ordinal);
        Assert.Contains("CASE-0044-EV", secondPages[1], StringComparison.Ordinal);
    }

    [Fact]
    public void Apply_RejectsDuplicatePageSelectionsBeforeAssigningNumbers() {
        var input = new PdfBatesDocument(PdfProductionWorkflowTestSupport.CreatePdf("Duplicate selection")) {
            TargetPages = PdfPageSelector.Parse("1,1")
        };

        Assert.Throws<ArgumentException>(() => PdfBatesNumberer.Apply(new[] { input }));
    }
}
