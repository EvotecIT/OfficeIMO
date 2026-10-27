using OfficeIMO.Markdown;
using OfficeIMO.Markdown.Pdf;
using OfficeIMO.Pdf;
using OfficeIMO.Word;

namespace OfficeIMO.Tests;

public sealed class OfficeConverterConversionTests {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void PdfModelConversionCompletesRenderingBeforeReturningItsReport(bool encrypt) {
        using var source = MarkdownDocument.Parse("Completed PDF conversion\n\n![Blocked image](https://example.invalid/image.png)");
        const string password = "synthetic-model-conversion-password";
        var options = new MarkdownPdfExportOptions {
            PdfOptions = encrypt ? new PdfCreateOptions().SetEncryption(password) : null,
            ResourcePolicy = new PdfResourcePolicy { AllowRemoteResourceResolution = false } };
        using PdfDocument result = source.ToPdf(options, out OfficeConversionReport report);
        Assert.True(report.HasLoss);
        Assert.True(report.ItemsWritten > 0);
        // Conversion reports a completed rendering while retaining the target authoring model.
        result.Compose(builder => builder.Content(content => content.Text("Post-conversion authoring")));
        byte[] bytes = result.ToBytes(new PdfSaveOptions { AllowLoss = true });
        Assert.NotEmpty(bytes);
        using var reopened = PdfDocument.Load(new MemoryStream(bytes), new PdfLoadOptions { Password = encrypt ? password : null });
        Assert.Contains("Completed PDF conversion", reopened.ExtractText());
        Assert.Contains("Post-conversion authoring", reopened.ExtractText());
        options.FailOnLoss = true;
        Assert.Throws<OfficeConversionException>(() => source.ToPdf(options));
    }

    [Fact]
    public void StreamConversionRestoresSourcePositionAndLeavesStreamsOpen() {
        using var word = WordDocument.Create();
        word.AddParagraph("Captured source text");
        using var source = new MemoryStream(word.ToBytes());
        source.Position = 3;
        using var destination = new MemoryStream();
        OfficeConversionReport report = OfficeConverter.Convert(source, OfficeDocumentFormat.Word, destination, OfficeDocumentFormat.Pdf);
        Assert.Equal(3, source.Position);
        Assert.True(source.CanRead);
        Assert.True(destination.CanWrite);
        using var pdf = PdfDocument.Load(new MemoryStream(destination.ToArray()));
        Assert.Contains("Captured source text", pdf.ExtractText());
        Assert.Equal(1, report.ItemsWritten);
    }

    [Fact]
    public void CapturedAdapterLossPolicyRejectsBeforeDestinationWrites() {
        using var source = new MemoryStream(Encoding.UTF8.GetBytes("<div>Explicit omission</div>"));
        using var destination = new MemoryStream(new byte[] { 1, 2, 3 }, writable: true);
        OfficeConversionException error = Assert.Throws<OfficeConversionException>(() => OfficeConverter.Convert(source,
            OfficeDocumentFormat.Markdown, destination, OfficeDocumentFormat.Html,
            new OfficeConversionOptions { MarkdownHtml = new MarkdownHtmlExportOptions { RawHtmlHandling = MarkdownRawHtmlHandling.Strip, FailOnLoss = true } }));
        var report = Assert.IsType<OfficeConversionReport>(error.Report);
        Assert.True(report.HasLoss);
        Assert.Contains(report.Diagnostics, finding => finding.Code == OfficeDiagnosticCodes.ContentOmitted && finding.LossKind == OfficeConversionLossKind.Omission);
        Assert.Equal(new byte[] { 1, 2, 3 }, destination.ToArray());
    }

    [Fact]
    public void FileConversionKeepsExistingOutputAndPreservesNativeSourceEncryption() {
        string directory = Directory.CreateTempSubdirectory("officeimo-converter-test-").FullName;
        try {
            string source = Path.Combine(directory, "source.docx"), destination = Path.Combine(directory, "result.docm");
            using (var word = WordDocument.Create()) {
                word.AddParagraph("Encrypted native source");
                word.Save(source, new WordSaveOptions { Password = "synthetic-password" });
            }
            OfficeConverter.Convert(source, destination, new OfficeConversionOptions { SourcePassword = "synthetic-password" });
            byte[] original = File.ReadAllBytes(destination);
            using (var reopened = WordDocument.Load(destination, new WordLoadOptions { Password = "synthetic-password" }))
                Assert.Contains(reopened.Paragraphs, paragraph => paragraph.Text == "Encrypted native source");
            var conflict = Assert.Throws<OfficeConversionException>(() => OfficeConverter.Convert(source, destination,
                new OfficeConversionOptions { SourcePassword = "synthetic-password" }));
            Assert.IsType<IOException>(conflict.InnerException);
            Assert.Contains(Assert.IsType<OfficeConversionReport>(conflict.Report).Diagnostics,
                finding => finding.Code == OfficeDiagnosticCodes.ConversionOutputFailed);
            Assert.Equal(original, File.ReadAllBytes(destination));
        } finally { Directory.Delete(directory, recursive: true); }
    }

    [Fact]
    public void FrontDoorSuppliesTheRawHtmlOwnerForMarkdownImport() {
        using var source = new MemoryStream(Encoding.UTF8.GetBytes("<div><strong>Rendered HTML</strong></div>"));
        using var output = new MemoryStream();
        OfficeConversionReport report = OfficeConverter.Convert(source, OfficeDocumentFormat.Markdown, output, OfficeDocumentFormat.Word);
        using var word = WordDocument.Load(new MemoryStream(output.ToArray()));
        Assert.Contains(word.Paragraphs, paragraph => paragraph.Text.Contains("Rendered HTML", StringComparison.Ordinal));
        Assert.DoesNotContain(report.Diagnostics, finding => finding.Code == OfficeDiagnosticCodes.HtmlBlockAsText);
    }
    [Fact]
    public void FrontDoorPreservesRawHtmlCellParagraphsAndInlineFormatting() {
        const string markdown = "<table><tr><th>Notes</th></tr><tr><td><custom-card><p>Alpha</p><p><strong>Beta</strong></p></custom-card></td></tr></table>";
        using var source = new MemoryStream(Encoding.UTF8.GetBytes(markdown));
        using var output = new MemoryStream();
        OfficeConversionReport report = OfficeConverter.Convert(source, OfficeDocumentFormat.Markdown, output, OfficeDocumentFormat.Word);
        using var word = WordDocument.Load(new MemoryStream(output.ToArray()));
        var paragraphs = word.Tables[0].Rows[1].Cells[0].Paragraphs.Where(paragraph => !string.IsNullOrWhiteSpace(paragraph.Text)).ToArray();
        Assert.Equal(new[] { "Alpha", "Beta" }, paragraphs.Select(paragraph => paragraph.Text));
        Assert.Contains(paragraphs.SelectMany(paragraph => paragraph.GetRuns()), run => run.Bold && run.Text == "Beta");
        Assert.DoesNotContain(report.Diagnostics, finding => finding.Code == OfficeDiagnosticCodes.HtmlBlockAsText);
    }

    [Fact]
    public void SourceReplacementDuringPublicationCannotReplaceExistingOutput() {
        string root = Directory.CreateTempSubdirectory("officeimo-source-publication-test-").FullName;
        try {
            string source = Path.Combine(root, "source.txt"), output = Path.Combine(root, "result.pdf");
            File.WriteAllText(source, "Captured original text");
            byte[] original = [1, 2, 3]; File.WriteAllBytes(output, original);
            var guard = new ReplaceSourceBeforePublication(source);
            var rejection = Assert.Throws<OfficeConversionException>(() => OfficeConverter.Convert(source, output,
                new OfficeConversionOptions { FileConflictPolicy = OfficePublicationPolicy.Replace, PublicationGuard = guard }));
            Assert.IsAssignableFrom<IOException>(rejection.InnerException);
            Assert.True(guard.Visited);
            Assert.Equal(original, File.ReadAllBytes(output));
            Assert.Empty(Directory.GetFiles(root, "*.tmp"));
        } finally { Directory.Delete(root, true); }
    }

    private sealed class ReplaceSourceBeforePublication(string source) : IOfficePublicationGuard {
        public bool Visited { get; private set; }
        public Task<bool> CanPublishAsync(string path, bool isDirectory, CancellationToken cancellationToken) {
            Visited = true;
            // Windows may refuse replacement while its source handle is open. Either
            // that refusal or the final source-identity check must preserve output.
            File.Move(source, source + ".original");
            File.WriteAllText(source, "Replacement text");
            return Task.FromResult(true);
        }
    }

}
