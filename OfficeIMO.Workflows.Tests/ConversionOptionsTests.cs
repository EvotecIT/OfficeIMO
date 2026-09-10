using System.Text;
using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Pdf;
using OfficeIMO.Word.Pdf;

namespace OfficeIMO.Workflows.Tests;

public sealed class ConversionOptionsTests {
    [Fact]
    public void FluentSettingsRemainIndependentOfCallerAndBuiltRequests() {
        var settings = new OfficeWorkflowConversionOptions { PageRanges = "2", WordMode = PdfWordImportMode.VisualPages };
        var builder = OfficeWorkflow.Convert("input.pdf").To("output.docx").WithConversionOptions(settings);
        settings.PageRanges = "1";
        var first = builder.Build();
        Assert.Equal("2", first.ConversionOptions!.PageRanges);
        first.ConversionOptions.PageRanges = "3";
        Assert.Equal("2", builder.Build().ConversionOptions!.PageRanges);
    }

    [Theory]
    [InlineData("pdf-docx", true)]
    [InlineData("pdf-html", true)]
    [InlineData("html-pdf", false)]
    public void RouteCapabilitiesAndInvalidOptionsHaveOneOwner(string id, bool pages) {
        var route = OfficeWorkflowCatalog.ExecutableRoutes.Single(item => item.Id == id);
        Assert.Equal(pages, route.SupportsPageSelection);
        Assert.Throws<ArgumentException>(() => new OfficeWorkflowConversionOptions { RasterDpi = 144 }.Snapshot(route));
        Assert.Throws<ArgumentException>(() => new OfficeWorkflowConversionOptions { WordMode = (PdfWordImportMode)999 }.Snapshot(route));
        if (!pages) Assert.Throws<ArgumentException>(() => new OfficeWorkflowConversionOptions { PageRanges = "1" }.Snapshot(route));
        if (!route.SupportsPdfCompression) Assert.Throws<ArgumentException>(() => new OfficeWorkflowConversionOptions { CompressPdfOutput = true }.Snapshot(route));
    }

    [Theory]
    [InlineData(PdfWordImportMode.EditableContent)]
    [InlineData(PdfWordImportMode.VisualPages)]
    public async Task WordRouteAppliesSelectedPageAndContentModeToPublishedArtifact(PdfWordImportMode mode) {
        string root = Path.Combine(Path.GetTempPath(), "officeimo-conversion-options-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(root);
        try {
            string input = Path.Combine(root, "source.pdf"), output = Path.Combine(root, "result.docx");
            File.WriteAllBytes(input, Source().ToBytes());
            var result = await new OfficeWorkflowRunner().RunAsync(OfficeWorkflow.Convert(input).To(output)
                .WithConversionOptions(new() { PageRanges = "2", WordMode = mode, RasterDpi = mode == PdfWordImportMode.VisualPages ? 72 : null }).Build());
            Assert.True(result.Status == OfficeWorkflowStatus.Completed, result.Summary + string.Join(" ", result.Diagnostics.Select(d => d.Message)));
            using var package = WordprocessingDocument.Open(output, false);
            string text = package.MainDocumentPart!.Document!.InnerText;
            if (mode == PdfWordImportMode.VisualPages) {
                Assert.Single(package.MainDocumentPart!.ImageParts);
                Assert.DoesNotContain("Second", text);
                Assert.Contains(result.Diagnostics, d => d.Message.Contains("VisualPagesNotEditable"));
            } else {
                Assert.Contains("Second", text);
                Assert.DoesNotContain("First", text);
            }
            var preview = OfficeWorkflowRunner.PreviewDocument(File.ReadAllBytes(output), ".docx");
            Assert.Single(preview.Pages);
            Assert.NotEmpty(preview.Pages[0].Bytes!);
        } finally { Directory.Delete(root, true); }
    }

    [Fact]
    public void PreviewSamplesOnlyThreePagesAndRejectsUnsupportedFormats() {
        var preview = OfficeWorkflowRunner.PreviewDocument(Source(5).ToBytes(), "pdf");
        Assert.Equal(new[] { 1, 2, 3 }, preview.Pages.Select(page => page.PageNumber));
        Assert.All(preview.Pages, page => Assert.NotEmpty(page.Bytes!));
        Assert.Throws<NotSupportedException>(() => OfficeWorkflowRunner.PreviewDocument([1], "exe"));
        Assert.ThrowsAny<OperationCanceledException>(() => OfficeWorkflowRunner.PreviewDocument(Source().ToBytes(), "pdf", new CancellationToken(true)));
    }

    [Theory]
    [InlineData("html")]
    [InlineData(".HTM")]
    public void HtmlPreviewAcceptsEverySourceExtension(string extension) {
        var preview = OfficeWorkflowRunner.PreviewDocument(Encoding.UTF8.GetBytes("<p>Saved HTML preview</p>"), extension);
        Assert.Single(preview.Pages);
        Assert.NotEmpty(preview.Pages[0].Bytes!);
    }

    [Fact]
    public async Task PdfCompressionPreservesConvertedTextAndReportsVerification() {
        string root = Path.Combine(Path.GetTempPath(), "officeimo-conversion-compression-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(root);
        try {
            string input = Path.Combine(root, "source.html"), output = Path.Combine(root, "result.pdf");
            File.WriteAllText(input, "<p>Conversion compression contract</p>", Encoding.UTF8);
            var result = await new OfficeWorkflowRunner().RunAsync(OfficeWorkflow.Convert(input).To(output)
                .WithConversionOptions(new() { CompressPdfOutput = true }).Build());
            Assert.True(result.Status == OfficeWorkflowStatus.Completed, result.Summary);
            Assert.Contains(result.Diagnostics, d => d.Code == "PdfOutputCompression");
            Assert.Contains("Conversion compression contract", PdfDocument.Load(output).Reader.Text());
        } finally { Directory.Delete(root, true); }
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task PdfCompressionNeverEnlargesAnAlreadyWithinLimitConversion(bool tightLimit) {
        string root = Path.Combine(Path.GetTempPath(), "officeimo-compression-size-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(root);
        try {
            string input = Path.Combine(root, "source.docx"), original = Path.Combine(root, "original.pdf"), output = Path.Combine(root, "compressed.pdf");
            using (var word = OfficeIMO.Word.WordDocument.Create(input)) {
                word.AddParagraph("x");
                word.Save();
            }
            var runner = new OfficeWorkflowRunner();
            var baseline = await runner.RunAsync(OfficeWorkflow.Convert(input).To(original).WithProfile(OfficeWorkflowOutputProfile.Lightweight).Build());
            Assert.Equal(OfficeWorkflowStatus.Completed, baseline.Status);
            long originalBytes = new FileInfo(original).Length;
            var result = await runner.RunAsync(OfficeWorkflow.Convert(input).To(output)
                .WithProfile(OfficeWorkflowOutputProfile.Lightweight)
                .WithLimits(1024 * 1024, tightLimit ? originalBytes : 1024 * 1024)
                .WithConversionOptions(new() { CompressPdfOutput = true }).Build());
            Assert.True(result.Status == OfficeWorkflowStatus.Completed, result.Summary + string.Join(" ", result.Diagnostics.Select(d => d.Message)));
            Assert.InRange(new FileInfo(output).Length, 1, originalBytes);
            Assert.Contains("x", PdfDocument.Load(output).Reader.Text());
        } finally { Directory.Delete(root, true); }
    }

    private static PdfDocument Source(int count = 2) => PdfDocument.Create(compose => {
        for (int index = 0; index < count; index++) {
            string text = index == 0 ? "First page" : "Second page " + index;
            compose.Page(page => page.Size(240, 320).Content(content => content.Item(item => item.Paragraph(paragraph => paragraph.Text(text)))));
        }
    });
}
