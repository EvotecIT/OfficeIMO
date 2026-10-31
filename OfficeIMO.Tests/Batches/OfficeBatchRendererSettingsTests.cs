using OfficeIMO.Pdf;
using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;

namespace OfficeIMO.Tests;

public sealed class OfficeBatchRendererSettingsTests {
    [Fact]
    public async Task SavedWordPdfRendererRetainsPhysicalGeometryMetadataAndPageNumberSettings() {
        string root = Path.Combine(Path.GetTempPath(), "officeimo-batch-renderer-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(root);
        try {
            string input = Path.Combine(root, "source.docx"), output = Path.Combine(root, "pdf"), path = Path.Combine(root, "job.json");
            using (var document = WordDocument.Create()) { document.AddParagraph("Saved renderer settings"); document.Save(input); }
            var renderer = new WordPdfExportOptions {
                PageSize = new PdfPageSize(OfficeLength.Points(500), OfficeLength.Points(700)),
                Margins = new PdfPageMargins(OfficeLength.Points(30), OfficeLength.Points(40), OfficeLength.Points(50), OfficeLength.Points(60)),
                Title = "Saved renderer", IncludePageNumbers = true, PageNumberFormat = "Page {PAGE}",
                ResourcePolicy = PdfResourcePolicy.CreatePortableDeterministic()
            };
            new OfficeBatchOptions { InputPaths = [input], OutputDirectory = output, ConversionOptions = new() { Word = renderer } }.Save(path);
            var restored = OfficeBatchOptions.Load(path);
            var actual = Assert.IsType<WordPdfExportOptions>(restored.ConversionOptions.Word);
            Assert.Equal(500, actual.PageSize!.Value.Width.ToPoints());
            Assert.Equal(700, actual.PageSize.Value.Height.ToPoints());
            Assert.Equal(50, actual.Margins!.Value.Right.ToPoints());
            Assert.Equal(60, actual.Margins.Value.Bottom.ToPoints());
            Assert.True(actual.IncludePageNumbers);
            Assert.Equal("Page {PAGE}", actual.PageNumberFormat);
            Assert.False(actual.ResourcePolicy.AllowSystemFontEmbedding);
            Assert.False(actual.ResourcePolicy.AllowRemoteResourceResolution);
            var report = await OfficeConverter.ConvertAsync(restored);
            Assert.Equal(1, report.Summary.Completed);
            string pdfPath = Assert.Single(report.Items).OutputPath!;
            using var pdf = PdfDocument.Load(pdfPath);
            Assert.Contains("Saved renderer settings", pdf.Read().Text);
            var page = Assert.Single(pdf.Inspect().Pages);
            Assert.Equal(500, page.Width);
            Assert.Equal(700, page.Height);
        } finally { Directory.Delete(root, true); }
    }
}
