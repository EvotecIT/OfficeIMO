using System.Text;
using System.Text.Json;
using OfficeIMO.AI;
using OfficeIMO.Drawing;
using OfficeIMO.Excel;
using OfficeIMO.Pdf;
using OfficeIMO.Reader;
using Xunit;

namespace OfficeIMO.AI.Tests;

public sealed class SourceAndArtifactTests {
    [Theory]
    [InlineData(16383, true)]
    [InlineData(20000, false)]
    public async Task UnicodeExportChecksExcelLimitsBeforeWritingArtifacts(int emojiCount, bool exportable) {
        string raw = string.Concat(Enumerable.Repeat("😀", emojiCount));
        var document = OfficeAiDocument.FromReadResult(Encoding.UTF8.GetBytes(raw), new OfficeDocumentReadResult {
            Blocks = new[] { new OfficeDocumentBlock { Text = raw } }
        });
        var result = await new OfficeAiEngine(new LiteralExecutor(raw)).RunAsync(document, new OfficeAiRequest {
            Operation = OfficeAiOperation.ExtractFields, Instruction = "Extract the literal value.",
            Fields = new[] { new OfficeAiFieldDefinition("value") },
            Limits = new() { MaxRequestCharacters = 1_000_000, MaxResponseCharacters = 1_000_000 }
        });
        Assert.Equal(OfficeAiResultStatus.Completed, result.Status);
        string output = Path.Combine(Path.GetTempPath(), "officeimo-ai-unicode-" + Guid.NewGuid().ToString("N"));
        try {
            if (exportable) {
                await ArtifactWriter.SaveAsync(output, document, result);
                using var workbook = ExcelDocument.Load(Path.Combine(output, "extraction.xlsx"));
                Assert.Equal(raw, workbook.Sheets[0].CellAt(2, 3).GetValue<string>());
            } else {
                await Assert.ThrowsAsync<NotSupportedException>(() => ArtifactWriter.SaveAsync(output, document, result));
                Assert.False(Directory.Exists(output));
            }
        } finally { if (Directory.Exists(output)) Directory.Delete(output, recursive: true); }
    }

    [Fact]
    public async Task ContentDetectedNativePdfDoesNotReportSourceOmissions() {
        byte[] source = PdfDocument.Create(builder => builder.Content(content => content.Text("Native total 42"))).ToBytes();
        var document = await DocumentInputs.ReadAsync(source, "attachment", false, Array.Empty<int>(), new(), CancellationToken.None);
        Assert.Contains(document.Evidence, item => item.Text.Contains("Native total 42"));
        Assert.False(document.HasSourceDiagnostics);
    }

    [Theory]
    [InlineData("attachment", true)]
    [InlineData("attachment.bin", true)]
    [InlineData("attachment.bin", false)]
    public async Task DetectedPdfRetainsNativeTextAndRendersPagesUnderAnyLogicalName(string name, bool images) {
        byte[] page = PdfDocument.Create(builder => builder.Content(content => content.Text("SCANNED"))).ToBytes();
        byte[] scan = PdfDocument.Load(page).ExportImages(OfficeImageExportFormat.Png).Single().Bytes;
        byte[] source = PdfDocument.Create(builder => builder.Content(content => content.Text("NATIVE")
            .PageBreak().Image(scan, 400, 566))).ToBytes();
        var document = await DocumentInputs.ReadAsync(source, name, images, Array.Empty<int>(), new(), CancellationToken.None);
        Assert.Contains(document.Evidence, item => item.Text.Contains("NATIVE"));
        Assert.Equal(new[] { 1, 2 }, document.Pages);
        if (images) {
            Assert.Equal(new[] { 1, 2 }, document.Images.Select(image => image.Page));
            var expected = PdfDocument.Load(source).ExportImages(OfficeImageExportFormat.Png, new PdfImageExportOptions { TargetDpi = 120 });
            Assert.Equal(expected[1].Bytes, document.Images[1].CopyBytes());
        } else Assert.Empty(document.Images);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task PasswordProtectedPdfCannotBecomeTextOrVisionEvidence(bool images) {
        byte[] bytes = PdfDocument.Create(builder => builder.Content(content => content.Text("Protected source")),
            new PdfOptions().SetEncryption("open", "owner")).ToBytes();
        await Assert.ThrowsAnyAsync<Exception>(() => DocumentInputs.ReadAsync(bytes, "protected.pdf", images,
            Array.Empty<int>(), new OfficeAiLimits(), CancellationToken.None));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void DeniedCopyAndPrintCannotBeBypassedWithVision(bool images) {
        byte[] bytes = PdfDocument.Create(builder => builder.Content(content => content.Text("Restricted source")),
            new PdfOptions().SetEncryption(new PdfStandardEncryptionOptions("open") {
                OwnerPassword = "owner", AllowedPermissions = PdfStandardPermissions.None
            })).ToBytes();
        var options = new PdfLoadOptions { Password = "open", PermissionPolicy = PdfPermissionPolicy.Enforce };
        if (images) Assert.Throws<PdfPermissionDeniedException>(() => PdfDocument.Load(bytes, options).ExportImages(OfficeImageExportFormat.Png));
        else Assert.Throws<PdfPermissionDeniedException>(() => PdfReadDocument.Open(bytes, options).ExtractText());
    }

    [Fact]
    public async Task MalformedPdfCannotBePromotedToEvidence() {
        await Assert.ThrowsAnyAsync<Exception>(() => DocumentInputs.ReadAsync(Encoding.ASCII.GetBytes("%PDF-1.7 broken"),
            "broken.pdf", true, Array.Empty<int>(), new OfficeAiLimits(), CancellationToken.None));
    }

    [Fact]
    public async Task RenderedImagesKeepCallerSelectedPageIdentity() {
        byte[] bytes = PdfDocument.Create(builder => builder.Content(content => content.Text("FIRST").PageBreak().Text("SECOND"))).ToBytes();
        var document = await DocumentInputs.ReadAsync(bytes, "pages.pdf", true, new[] { 2, 1 }, new OfficeAiLimits(), CancellationToken.None);
        var direct = PdfDocument.Load(bytes).ExportImages(OfficeImageExportFormat.Png,
            new PdfImageExportOptions { TargetDpi = 120 });
        Assert.Equal(new[] { 2, 1 }, document.Images.Select(image => image.Page));
        Assert.Equal(direct[1].Bytes, document.Images[0].CopyBytes());
        Assert.Equal(direct[0].Bytes, document.Images[1].CopyBytes());
    }

    [Fact]
    public async Task ExtractedFormulaLikeTextRemainsLiteralInExcelAndEscapedInCsv() {
        const string raw = "=2+3";
        var document = OfficeAiDocument.FromReadResult(Encoding.UTF8.GetBytes(raw), new OfficeDocumentReadResult {
            Blocks = new[] { new OfficeDocumentBlock { Text = raw } }
        });
        var result = await new OfficeAiEngine(new LiteralExecutor()).RunAsync(document, new OfficeAiRequest {
            Operation = OfficeAiOperation.ExtractFields, Instruction = "Extract the literal value.", Fields = new[] { new OfficeAiFieldDefinition("value") }
        });
        Assert.Equal(OfficeAiResultStatus.Completed, result.Status);
        string output = Path.Combine(Path.GetTempPath(), "officeimo-ai-artifact-" + Guid.NewGuid().ToString("N"));
        try {
            await ArtifactWriter.SaveAsync(output, document, result);
            using var workbook = ExcelDocument.Load(Path.Combine(output, "extraction.xlsx"));
            Assert.Equal(raw, workbook.Sheets[0].CellAt(2, 3).GetValue<string>());
            Assert.Contains("'=2+3", File.ReadAllText(Path.Combine(output, "Fields.csv")));
            using var report = JsonDocument.Parse(File.ReadAllText(Path.Combine(output, "report.json")));
            Assert.Contains(document.SourceHash, report.RootElement.GetRawText());
            await Assert.ThrowsAsync<IOException>(() => ArtifactWriter.SaveAsync(output, document, result));
        } finally { if (Directory.Exists(output)) Directory.Delete(output, recursive: true); }
    }

    private sealed class LiteralExecutor(string raw = "=2+3") : IOfficeAiExecutor {
        public OfficeAiExecutionProfile Profile { get; } = new() { Id = "artifact", Provider = "fixture", Model = "fixture", IsLocal = true, MaxRequestCharacters = 1_000_000 };
        public Task<OfficeAiExecutionResponse> ExecuteAsync(OfficeAiExecutionRequest request, CancellationToken cancellationToken = default) =>
            Task.FromResult(new OfficeAiExecutionResponse(JsonSerializer.Serialize(new {
                claims = Array.Empty<object>(), fields = new { field1 = new { status = "present",
                    rawValue = raw, evidence = new[] { new { id = "e1", quote = raw } } } }, blocks = Array.Empty<object>(), tables = Array.Empty<object>()
            })));
    }
}
