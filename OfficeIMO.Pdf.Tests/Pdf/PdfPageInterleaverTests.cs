using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public class PdfPageInterleaverTests {
    [Fact]
    public void Interleave_AlternatesPagesAndAppendsRemainderWithProvenance() {
        byte[] first = PdfProductionWorkflowTestSupport.CreatePdf("A one", "A two", "A three");
        byte[] second = PdfProductionWorkflowTestSupport.CreatePdf("B one", "B two");

        PdfInterleaveResult result = PdfPageInterleaver.Interleave(
            new[] { new PdfInterleaveSource(first, "A"), new PdfInterleaveSource(second, "B") });

        Assert.Equal(5, result.Pages.Count);
        Assert.Equal(new[] { "A", "B", "A", "B", "A" }, result.Pages.Select(static page => page.SourceName));
        Assert.Equal(new[] { 1, 1, 2, 2, 3 }, result.Pages.Select(static page => page.SourcePageNumber));
        Assert.Equal(
            new[] { "Aone", "Bone", "Atwo", "Btwo", "Athree" },
            PdfProductionWorkflowTestSupport.ReadPageTexts(result.ToBytes()));
        Assert.Equal(5, PdfInspector.Inspect(result.ToBytes()).PageCount);
    }

    [Fact]
    public void Interleave_HonorsReverseSelectionAndRejectsUnevenInputsWhenRequested() {
        byte[] first = PdfProductionWorkflowTestSupport.CreatePdf("A one", "A two");
        byte[] second = PdfProductionWorkflowTestSupport.CreatePdf("B one", "B two");
        var reversed = new PdfInterleaveSource(second, "B") { Reverse = true };

        PdfInterleaveResult result = PdfPageInterleaver.Interleave(
            new[] { new PdfInterleaveSource(first, "A"), reversed },
            new PdfInterleaveOptions { RemainderMode = PdfInterleaveRemainderMode.Reject });

        Assert.Equal(
            new[] { "Aone", "Btwo", "Atwo", "Bone" },
            PdfProductionWorkflowTestSupport.ReadPageTexts(result.ToBytes()));
        Assert.Throws<InvalidOperationException>(() => PdfPageInterleaver.Interleave(
            new[] {
                new PdfInterleaveSource(PdfProductionWorkflowTestSupport.CreatePdf("one")),
                new PdfInterleaveSource(first)
            },
            new PdfInterleaveOptions { RemainderMode = PdfInterleaveRemainderMode.Reject }));
    }

    [Fact]
    public void Interleave_ReportsOnlySelectedPagesAsImported() {
        var selected = new PdfInterleaveSource(PdfProductionWorkflowTestSupport.CreatePdf("A one", "A two", "A three")) {
            Pages = PdfPageSelector.Parse("2")
        };

        PdfInterleaveResult result = PdfPageInterleaver.Interleave(
            new[] { selected, new PdfInterleaveSource(PdfProductionWorkflowTestSupport.CreatePdf("B one")) });

        Assert.Equal(1, result.MergeReport.Sources[0].PageCount);
        Assert.Equal(1, result.MergeReport.Sources[1].PageCount);
        Assert.Equal(new[] { "Atwo", "Bone" }, PdfProductionWorkflowTestSupport.ReadPageTexts(result.ToBytes()));
    }
}
