using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Pdf;
using OfficeIMO.Drawing;
using OfficeIMO.Pdf.Word;
using Xunit;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfRuledTableFigureTests {
    [Fact]
    public void KeynoteCellGridRetainsItsTextInsteadOfDuplicatingTheTableAsAPicture() {
        string path = Path.Combine(RepositoryTestPaths.Find(), "OfficeIMO.TestAssets", "Documents", "IWorkCorpus",
            "native-exports", "keynote-created-tables-v15.4.pdf");
        string body = ImportEditableBody(path);
        foreach (string label in new[] { "Product", "Units", "Revenue", "Active", "1200,5", "-25,75", "PRAWDA", "FAŁSZ" })
            Assert.Contains(label, body, StringComparison.Ordinal);
    }

    [Fact]
    public void RuledWordRowsRemainEditableEvenWhenTextInferenceDoesNotRecoverTheGrid() {
        string path = Path.Combine(AppContext.BaseDirectory, "Pdf", "Fixtures", "Interoperability",
            "RuledVectorTables", "word-business-summary.pdf");
        string body = ImportEditableBody(path);
        Assert.Contains("Worksheet", body, StringComparison.Ordinal);
        Assert.Contains("Purpose", body, StringComparison.Ordinal);
        Assert.Contains("Review Tracker", body, StringComparison.Ordinal);
    }

    [Fact]
    public void OutlinedStackedBarSegmentsRemainAChartInsteadOfBecomingATableGrid() {
        OfficeSeries[] series = Enumerable.Range(1, 3)
            .Select(index => new OfficeSeries("Series " + index, new[] { (double)index })
                .WithSeriesStyle(new OfficeChartPointStyle(outlineColor: OfficeColor.Black, outlineWidth: OfficeLength.Points(1D), showOutline: true)))
            .ToArray();
        OfficeDrawing drawing = OfficeChartDrawingRenderer.Render(new OfficeChartSnapshot(
            "Stacked results", "Results", OfficeChartType.BarStacked100,
            new OfficeChartData(new[] { "Quarter" }, series), OfficeLength.Points(400D), OfficeLength.Points(220D)));
        using var source = PdfDocument.Create(new PdfCreateOptions { PageWidth = OfficeLength.Points(500D), PageHeight = OfficeLength.Points(320D),
            MarginLeft = OfficeLength.Points(36D), MarginRight = OfficeLength.Points(36D), MarginTop = OfficeLength.Points(36D), MarginBottom = OfficeLength.Points(36D) }).Drawing(drawing, PdfAlign.Left);
        using var pdf = PdfDocument.Load(source.ToBytes());
        Assert.NotEmpty(pdf.Read().Pages[0].VectorFigures);
    }

    private static string ImportEditableBody(string path) {
        using var pdf = PdfDocument.Load(path);
        using var output = new MemoryStream();
        var report = pdf.SaveAsWord(output);
        Assert.DoesNotContain(report.Diagnostics, diagnostic => diagnostic.Code == "pdf-vector-figure-rasterized");
        using WordprocessingDocument document = WordprocessingDocument.Open(new MemoryStream(output.ToArray()), false);
        return string.Join(" ", document.MainDocumentPart!.Document!.Body!.Descendants<W.Paragraph>()
            .Select(paragraph => string.Concat(paragraph.Descendants<W.Text>().Select(text => text.Text))));
    }
}
