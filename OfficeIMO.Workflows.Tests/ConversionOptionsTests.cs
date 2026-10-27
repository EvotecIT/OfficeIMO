using OfficeIMO.Pdf.Word;
using System.Text;
using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Pdf;
using OfficeIMO.Word.Pdf;

namespace OfficeIMO.Workflows.Tests;

public sealed class ConversionOptionsTests {
    [Fact]
    public void ClonedConversionSettingsRetainIndependentFormatOptions() {
        var settings = new OfficeConversionOptions {
            PdfLoad = new PdfLoadOptions { PageSelection = PdfPageSelection.Parse("2") },
            PdfWord = WordPdfImportOptions.CreateVisualPages()
        };
        OfficeConversionOptions first = settings.Clone();
        settings.PdfLoad.PageSelection = PdfPageSelection.Parse("1");
        Assert.Equal("2", first.PdfLoad!.PageSelection!.ToString());
        OfficeConversionOptions second = first.Clone();
        first.PdfLoad.PageSelection = PdfPageSelection.Parse("3");
        Assert.Equal("2", second.PdfLoad!.PageSelection!.ToString());
    }

    [Theory]
    [InlineData("pdf-docx", true)]
    [InlineData("pdf-html", true)]
    [InlineData("html-pdf", false)]
    public void RouteCapabilitiesAndInvalidOptionsHaveOneOwner(string id, bool pages) {
        var route = OfficeWorkflowCatalog.ExecutableRoutes.Single(item => item.Id == id);
        Assert.Equal(pages, route.SupportsPageSelection);
        Assert.Throws<ArgumentException>(() => OfficeIMO.Workflows.OfficeWorkflowRenderingOptions.Capture(new OfficeConversionOptions {
            PdfWord = new WordPdfImportOptions { Dpi = 144 }
        }).Snapshot(route));
        Assert.Throws<ArgumentException>(() => OfficeIMO.Workflows.OfficeWorkflowRenderingOptions.Capture(new OfficeConversionOptions {
            PdfWord = new WordPdfImportOptions { Mode = (WordPdfImportMode)999 }
        }).Snapshot(route));
        if (!pages) Assert.Throws<ArgumentException>(() => OfficeIMO.Workflows.OfficeWorkflowRenderingOptions.Capture(new OfficeConversionOptions {
            PdfLoad = "1" == null ? null : new OfficeIMO.Pdf.PdfLoadOptions { PageSelection = OfficeIMO.Pdf.PdfPageSelection.Parse("1") }
        }).Snapshot(route));
        if (!route.SupportsPdfCompression) Assert.Throws<ArgumentException>(() => OfficeIMO.Workflows.OfficeWorkflowRenderingOptions.Capture(new OfficeConversionOptions {
            CompressPdfOutput = true
        }).Snapshot(route));
    }

    [Theory]
    [InlineData(WordPdfImportMode.EditableContent)]
    [InlineData(WordPdfImportMode.VisualPages)]
    public async Task WordRouteAppliesSelectedPageAndContentModeToPublishedArtifact(WordPdfImportMode mode) {
        string root = Path.Combine(Path.GetTempPath(), "officeimo-conversion-options-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(root);
        try {
            string input = Path.Combine(root, "source.pdf"), output = Path.Combine(root, "result.docx");
            File.WriteAllBytes(input, Source().ToBytes());
            var result = await new OfficeWorkflowRunner().RunAsync(((new OfficeWorkflowRequest {
                Operation = OfficeWorkflowOperation.Convert,
                InputPath = input,
                OutputPath = output,
                ConversionOptions = (new OfficeConversionOptions() {
                    PdfLoad = new PdfLoadOptions { PageSelection = PdfPageSelection.Parse("2") },
                    PdfWord = new WordPdfImportOptions { Mode = mode, Dpi = mode == WordPdfImportMode.VisualPages ? 72 : 144 }
                }

    ).Clone()
,
                ConversionRouteId = OfficeWorkflowCatalog.Find(Path.GetExtension(input), Path.GetExtension(output), executableOnly: true)?.Id
            })));
            Assert.True(result.Status == OfficeWorkflowStatus.Completed, result.Summary + string.Join(" ", result.Diagnostics.Select(d => d.Message)));
            using var package = WordprocessingDocument.Open(output, false);
            string text = package.MainDocumentPart!.Document!.InnerText;
            if (mode == WordPdfImportMode.VisualPages) {
                Assert.Single(package.MainDocumentPart!.ImageParts);
                Assert.DoesNotContain("Second", text);
                Assert.Contains(result.Diagnostics, d => d.Code == "visual-pages-not-editable"
                    && d.Details["source"] == "OfficeIMO.Pdf.Word" && d.Details["lossKind"] == "Approximation");
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
            var result = await new OfficeWorkflowRunner().RunAsync(((new OfficeWorkflowRequest {
                Operation = OfficeWorkflowOperation.Convert,
                InputPath = input,
                OutputPath = output,
                ConversionOptions = (new OfficeConversionOptions() {
                    CompressPdfOutput = true
                }

    ).Clone()
,
                ConversionRouteId = OfficeWorkflowCatalog.Find(Path.GetExtension(input), Path.GetExtension(output), executableOnly: true)?.Id
            })));
            Assert.True(result.Status == OfficeWorkflowStatus.Completed, result.Summary);
            Assert.Contains(result.Diagnostics, d => d.Code == "pdf-output-compressed");
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
            var baseline = await runner.RunAsync(((new OfficeWorkflowRequest {
                Operation = OfficeWorkflowOperation.Convert,
                InputPath = input,
                OutputPath = original,
                OutputProfile = OfficeConversionProfile.Lightweight
,
                ConversionRouteId = OfficeWorkflowCatalog.Find(Path.GetExtension(input), Path.GetExtension(original), executableOnly: true)?.Id
            })));
            Assert.Equal(OfficeWorkflowStatus.Completed, baseline.Status);
            long originalBytes = new FileInfo(original).Length;
            var result = await runner.RunAsync(((new OfficeWorkflowRequest {
                Operation = OfficeWorkflowOperation.Convert,
                InputPath = input,
                OutputPath = output,
                OutputProfile = OfficeConversionProfile.Lightweight,
                Limits = new OfficeOperationLimits {
                    MaximumInputBytes = 1024 * 1024,
                    MaximumOutputBytes = tightLimit ? originalBytes : 1024 * 1024
                },
                ConversionOptions = (new OfficeConversionOptions() {
                    CompressPdfOutput = true
                }

    ).Clone()
,
                ConversionRouteId = OfficeWorkflowCatalog.Find(Path.GetExtension(input), Path.GetExtension(output), executableOnly: true)?.Id
            })));
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