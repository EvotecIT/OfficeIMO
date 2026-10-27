using System.Text;
using OfficeIMO.Excel;
using OfficeIMO.Pdf;
using OfficeIMO.PowerPoint;
using OfficeIMO.Word;
using OfficeIMO.Word.Markdown;

namespace OfficeIMO.Tests;

public sealed class OfficeConverterOutputSecurityTests {
    [Theory]
    [InlineData(".pdf", ".docx", false)]
    [InlineData(".pdf", ".xlsx", false)]
    [InlineData(".pdf", ".pptx", false)]
    [InlineData(".md", ".docx", false)]
    [InlineData(".html", ".docx", false)]
    [InlineData(".pdf", ".docx", true)]
    [InlineData(".pdf", ".xlsx", true)]
    [InlineData(".pdf", ".pptx", true)]
    public async Task CrossFormatOutputsHonorDestinationPasswords(string sourceExtension, string targetExtension, bool stream) {
        string directory = Path.Combine(Path.GetTempPath(), "officeimo-output-security-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        try {
            byte[] source;
            if (sourceExtension == ".pdf") {
                using var pdf = PdfDocument.Create();
                pdf.Compose(builder => builder.Content(content => content.Text("Protected reconstructed content")));
                source = pdf.ToBytes(new PdfSaveOptions { AllowLoss = true });
            } else source = Encoding.UTF8.GetBytes(sourceExtension == ".html" ? "<p>Protected imported content</p>" : "Protected imported content");
            const string password = "synthetic-destination-password";
            var options = new OfficeConversionOptions {
                SourceName = "source" + sourceExtension, TargetName = "output" + targetExtension,
                WordSave = targetExtension == ".docx" ? new WordSaveOptions { Password = password } : null,
                ExcelSave = targetExtension == ".xlsx" ? new ExcelSaveOptions { Password = password } : null,
                PowerPointSave = targetExtension == ".pptx" ? new PowerPointSaveOptions { Password = password } : null
            };
            byte[] output;
            if (stream) {
                using var input = new MemoryStream(source);
                using var destination = new MemoryStream();
                await OfficeConverter.ConvertAsync(input, OfficeDocumentFormat.Pdf, destination,
                    targetExtension == ".docx" ? OfficeDocumentFormat.Word : targetExtension == ".xlsx" ? OfficeDocumentFormat.Excel : OfficeDocumentFormat.PowerPoint, options);
                output = destination.ToArray();
            } else {
                string input = Path.Combine(directory, options.SourceName), destination = Path.Combine(directory, options.TargetName);
                File.WriteAllBytes(input, source);
                await OfficeConverter.ConvertAsync(input, destination, options);
                output = File.ReadAllBytes(destination);
            }
            Assert.Equal(new byte[] { 0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1 }, output.Take(8));
            using var encrypted = new MemoryStream(output);
            if (targetExtension == ".docx") { using var document = WordDocument.Load(encrypted, new WordLoadOptions { Password = password }); Assert.NotEmpty(document.Sections); }
            else if (targetExtension == ".xlsx") { using var document = ExcelDocument.Load(encrypted, new ExcelLoadOptions { Password = password }); Assert.NotEmpty(document.Sheets); }
            else { using var document = PowerPointPresentation.Load(encrypted, new PowerPointLoadOptions { Password = password }); Assert.NotEmpty(document.Slides); }
        } finally { Directory.Delete(directory, true); }
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task PublicationRejectionsRetainCompletedConversionFindings(bool guarded) {
        string directory = Path.Combine(Path.GetTempPath(), "officeimo-report-publication-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        try {
            string input = Path.Combine(directory, "source.docx"), output = Path.Combine(directory, "result.md");
            using (var source = WordDocument.Create(input)) {
                source.AddHeadersAndFooters(); source.Header!.Default!.AddParagraph("Running content");
                source.AddParagraph("Body content"); source.Save();
            }
            File.WriteAllText(output, "original destination");
            var error = await Assert.ThrowsAsync<OfficeConversionException>(() => OfficeConverter.ConvertAsync(input, output,
                new OfficeConversionOptions { PublicationGuard = guarded ? new RejectPublication() : null }));
            var report = Assert.IsType<OfficeConversionReport>(error.Report);
            Assert.Contains(report.Diagnostics, diagnostic => diagnostic.Code == OfficeDiagnosticCodes.WordToMarkdownWarning && diagnostic.LossKind == OfficeConversionLossKind.Omission);
            Assert.Contains(report.Diagnostics, diagnostic => diagnostic.Code == OfficeDiagnosticCodes.ConversionOutputFailed);
            Assert.Equal("original destination", File.ReadAllText(output));
        } finally { Directory.Delete(directory, true); }
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task AnalysisAndPublicationRejectionPreserveExistingSidecars(bool analysis) {
        string directory = Path.Combine(Path.GetTempPath(), "officeimo-sidecar-publication-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        try {
            string input = Path.Combine(directory, "source.docx"), output = Path.Combine(directory, "result.md");
            string assets = Path.Combine(directory, "assets"); Directory.CreateDirectory(assets);
            string sidecar = Path.Combine(assets, "01-chart.svg");
            using (var source = WordDocument.Create(input)) {
                var chart = source.AddChart("Chart", width: 400, height: 240);
                chart.AddCategories(new List<string> { "Q1", "Q2" });
                chart.AddBar("Series", new List<int> { 10, 20 }, OfficeColor.CornflowerBlue);
                source.Save();
            }
            File.WriteAllText(sidecar, "original sidecar"); File.WriteAllText(output, "original markdown");
            var options = new OfficeConversionOptions {
                FileConflictPolicy = OfficePublicationPolicy.Replace,
                PublicationGuard = new DenyMainOutput(output),
                WordMarkdown = new WordMarkdownExportOptions {
                    VisualFallbackMode = WordMarkdownVisualFallbackMode.SvgFile,
                    VisualFallbackDirectory = assets, VisualFallbackPathPrefix = "assets"
                }
            };
            if (analysis) {
                OfficeConversionReport report = await OfficeConverter.AnalyzeAsync(input, output, options);
                Assert.True(report.HasLoss);
            } else {
                var error = await Assert.ThrowsAsync<OfficeConversionException>(() => OfficeConverter.ConvertAsync(input, output, options));
                Assert.True(error.Report.HasLoss);
            }
            Assert.Equal("original sidecar", File.ReadAllText(sidecar));
            Assert.Equal("original markdown", File.ReadAllText(output));
        } finally { Directory.Delete(directory, true); }
    }

    private sealed class DenyMainOutput(string output) : IOfficePublicationGuard {
        public Task<bool> CanPublishAsync(string absoluteDestination, bool isDirectory, CancellationToken cancellationToken) => Task.FromResult(absoluteDestination != output);
    }

    private sealed class RejectPublication : IOfficePublicationGuard {
        public Task<bool> CanPublishAsync(string absoluteDestination, bool isDirectory, CancellationToken cancellationToken) => Task.FromResult(false);
    }
}
