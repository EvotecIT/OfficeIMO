using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfRotatedTableGeometryTests {
    [Theory]
    [InlineData("latin-0", true)]
    [InlineData("latin-90", false)]
    [InlineData("latin-180", true)]
    [InlineData("latin-270", false)]
    [InlineData("rtl-90", false)]
    [InlineData("rtl-270", false)]
    public void ContinuationComparisonUsesTheActualColumnProgressionAxis(string id, bool compatibleAfterVerticalCrop) {
        byte[] source = File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "MultilingualLayout", id + "-native.pdf"));
        PdfLogicalPage original = Read(source);
        PdfLogicalPage same = Read(source);
        PdfLogicalPage cropped = Read(source, cropTop: 20D);
        Assert.True(PdfLogicalTableContinuations.HasCompatibleColumns(
            Assert.Single(original.Tables), original, Assert.Single(same.Tables), same, 4D));
        Assert.Equal(compatibleAfterVerticalCrop, PdfLogicalTableContinuations.HasCompatibleColumns(
            Assert.Single(original.Tables), original, Assert.Single(cropped.Tables), cropped, 4D));
    }

    [Theory]
    [InlineData("latin-90")]
    [InlineData("latin-180")]
    [InlineData("latin-270")]
    [InlineData("rtl-90")]
    [InlineData("rtl-180")]
    [InlineData("rtl-270")]
    public void ColumnRectanglesSurviveLogicalAndLegacyProjection(string id) {
        byte[] source = File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "MultilingualLayout", id + "-native.pdf"));
        PdfLogicalPage page = Read(source);
        PdfLogicalPage cropped = Read(source, cropTop: 20D);
        PdfLogicalTable table = Assert.Single(page.Tables);
        PdfLogicalTable shifted = Assert.Single(cropped.Tables);
        PdfUnderstandingTableCandidate candidate = Assert.Single(page.Analysis.TableCandidates);
        PdfReadPage rawPage = PdfReadDocument.Open(source).Pages[0];
        StructuredTable legacy = candidate.ToStructuredTable(rawPage);
        Assert.True(legacy.YTop > legacy.YBottom);
        Assert.True(PdfLogicalTableContinuations.HasCompatibleColumns(table, page, PdfLogicalTable.From(1, legacy), page, 0.001D));
        PdfLogicalTableDiagnostics diagnostics = PdfLogicalTableAnalysis.Extract(table).Diagnostics;
        Assert.True(diagnostics.Height > 0D);
        Assert.Equal(1D, diagnostics.ColumnGeometryConfidence, 6);
        for (int index = 0; index < table.Columns.Count; index++) {
            PdfLogicalVisualBounds bounds = Assert.IsType<PdfLogicalVisualBounds>(table.Columns[index].VisualBounds);
            Assert.Same(bounds, candidate.Columns[index].VisualBounds);
            Assert.Same(bounds, legacy.Columns[index].VisualBounds);
            Assert.Equal(bounds.Left, table.Columns[index].From, 6);
            Assert.Equal(bounds.Right, table.Columns[index].To, 6);
            Assert.Equal(bounds.Left, shifted.Columns[index].VisualBounds!.Left, 6);
            Assert.Equal(bounds.Top - 20D, shifted.Columns[index].VisualBounds!.Top, 6);
            Assert.Equal(bounds.Bottom - 20D, shifted.Columns[index].VisualBounds!.Bottom, 6);
        }
        PdfLogicalVisualBounds first = table.Columns[0].VisualBounds!, second = table.Columns[1].VisualBounds!;
        Assert.True(first.Right <= second.Left || second.Right <= first.Left || first.Bottom <= second.Top || second.Bottom <= first.Top);
    }

    private static PdfLogicalPage Read(byte[] bytes, double cropTop = 0D) {
        PdfReadDocument document = PdfReadDocument.Open(bytes);
        PdfReadPage page = document.Pages[0];
        if (cropTop > 0D) {
            (double width, double height) = page.GetPageSize();
            // Vary only the input page box; keep the independent producer's content intact.
            var crop = new PdfArray();
            foreach (double value in new[] { 0D, 0D, width, height - cropTop }) crop.Items.Add(new PdfNumber(value));
            page.PageDictionary.Items["CropBox"] = crop;
        }
        var layout = new PdfTextLayoutOptions();
        PdfUnderstandingPageResult analysis = new PdfUnderstandingPipeline(layout, PdfUnderstandingPipelineOptions.Structured())
            .RunPages(document, new[] { 1 })[0];
        return PdfLogicalPage.From(document, page, 1, layout, analysis: analysis);
    }
}
