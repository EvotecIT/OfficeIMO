using System.Text.Json;
using OfficeIMO.Ocr;
using OfficeIMO.Pdf;
using OfficeIMO.Tool.Agent;
using OfficeIMO.Tool.Commands.Pdf;
using OfficeIMO.Tool.Mcp;
using ModelContextProtocol.Protocol;
using Xunit;

namespace OfficeIMO.Tool.Tests;

public sealed class PdfAutomationCommandTests {
    [Fact]
    public async Task ExtractionPreservesRelativeOrderAndRepeatsWithoutReplacingSource() {
        using var scope = new DirectoryScope();
        string source = scope.File("source.pdf"), destination = scope.File("extracted.pdf");
        CreatePdf(source, 3);
        byte[] original = await File.ReadAllBytesAsync(source);
        var copied = await Run(["extract", source, "--pages", "last,1,last", "--output", destination]);
        Assert.Equal(0, copied.Exit);
        var pages = PdfReadDocument.Open(await File.ReadAllBytesAsync(destination)).Pages;
        Assert.Equal(["Page 3", "Page 1", "Page 3"], pages.Select(page => page.ExtractText().Trim()));
        Assert.Equal(original, await File.ReadAllBytesAsync(source));
        var replaceSource = await Run(["extract", source, "--pages", "1", "--output", source, "--force"]);
        Assert.Equal((int)OfficeImoToolExitCode.Usage, replaceSource.Exit);
        Assert.Equal(original, await File.ReadAllBytesAsync(source));
        var existing = await Run(["extract", source, "--pages", "1", "--output", destination]);
        Assert.Equal((int)OfficeImoToolExitCode.Usage, existing.Exit);
    }

    [Fact]
    public async Task SplittingPublishesBoundedPartsAndRejectsAnOutputFolderContainingSource() {
        using var scope = new DirectoryScope();
        string source = scope.File("source.pdf"), destination = scope.File("parts");
        CreatePdf(source, 5);
        var result = await Run(["split", source, "--output", destination, "--pages-per-document", "2"]);
        Assert.Equal(0, result.Exit);
        Assert.Equal([2, 2, 1], Directory.GetFiles(destination, "*.pdf").OrderBy(path => path).Select(path => PdfDocument.Load(path).Inspect().PageCount));
        using var json = JsonDocument.Parse(result.Output);
        Assert.Equal(3, json.RootElement.GetProperty("artifactCount").GetInt32());
        var unsafeFolder = await Run(["split", source, "--output", scope.Root, "--force"]);
        Assert.Equal((int)OfficeImoToolExitCode.Usage, unsafeFolder.Exit);
        Assert.True(File.Exists(source));
        var limited = await Run(["split", source, "--output", scope.File("too-many"), "--maximum-pages", "2"]);
        Assert.NotEqual(0, limited.Exit);
        Assert.False(Directory.Exists(scope.File("too-many")));
    }

    [Fact]
    public async Task DecryptionRequiresOwnerAuthorizationAndNeverReportsThePassword() {
        using var scope = new DirectoryScope();
        string source = scope.File("protected.pdf"), destination = scope.File("clear.pdf");
        byte[] clear = PdfDocument.Create(d => d.Page(p => p.Content(c => c.Item(i => i.AddParagraph(t => t.AddText("Retained content")))))).ToBytes();
        byte[] protectedBytes = PdfDocument.Load(clear).Security.Encrypt(new("reader-secret") { OwnerPassword = "owner-secret" }).Pdf;
        await File.WriteAllBytesAsync(source, protectedBytes);
        string variable = "OFFICEIMO_PDF_TEST_" + Guid.NewGuid().ToString("N");
        try {
            Environment.SetEnvironmentVariable(variable, "reader-secret");
            var refused = await Run(["decrypt", source, "--output", destination, "--password-env", variable]);
            Assert.NotEqual(0, refused.Exit); Assert.False(File.Exists(destination));
            Environment.SetEnvironmentVariable(variable, "owner-secret");
            var result = await Run(["decrypt", source, "--output", destination, "--password-env", variable]);
            Assert.Equal(0, result.Exit);
            var reopened = PdfDocument.Load(destination);
            Assert.False(reopened.Inspect().Security.HasEncryption);
            Assert.Contains("Retained content", PdfReadDocument.Open(await File.ReadAllBytesAsync(destination)).Pages[0].ExtractText());
            Assert.DoesNotContain("owner-secret", result.Output + result.Error);
            Assert.DoesNotContain("reader-secret", refused.Output + refused.Error);
            Assert.Equal(protectedBytes, await File.ReadAllBytesAsync(source));
        } finally { Environment.SetEnvironmentVariable(variable, null); }
    }

    [Fact]
    public async Task RasterFlatteningRequiresAcknowledgementAndProducesSelectedImagePages() {
        using var scope = new DirectoryScope();
        string source = scope.File("source.pdf"), destination = scope.File("flattened.pdf");
        CreatePdf(source, 2);
        var refused = await Run(["flatten", source, "--output", destination]);
        Assert.Equal((int)OfficeImoToolExitCode.Usage, refused.Exit); Assert.False(File.Exists(destination));
        var result = await Run(["flatten", source, "--output", destination, "--pages", "last", "--dpi", "72", "--acknowledge-raster-output"]);
        Assert.Equal(0, result.Exit);
        var reopened = PdfDocument.Load(destination);
        Assert.Equal(1, reopened.Inspect().PageCount);
        Assert.True(string.IsNullOrWhiteSpace(PdfReadDocument.Open(await File.ReadAllBytesAsync(destination)).Pages[0].ExtractText()));
        Assert.Single(reopened.Images.Placements());
        var limited = await Run(["flatten", source, "--output", scope.File("limited.pdf"), "--maximum-pixels-per-page", "1", "--acknowledge-raster-output"]);
        Assert.NotEqual(0, limited.Exit); Assert.False(File.Exists(scope.File("limited.pdf")));
    }

    [Fact]
    public async Task SearchableOcrPublishesTextThroughConfiguredProviderWithoutReturningRecognizedText() {
        using var scope = new DirectoryScope();
        string source = scope.File("scan.pdf"), destination = scope.File("searchable.pdf");
        PdfDocument.Create(d => { d.Page(p => p.Size(global::OfficeIMO.OfficeLength.Points(200), global::OfficeIMO.OfficeLength.Points(100)).Margin(global::OfficeIMO.OfficeLength.Points(10))); d.Page(p => p.Size(global::OfficeIMO.OfficeLength.Points(200), global::OfficeIMO.OfficeLength.Points(100)).Margin(global::OfficeIMO.OfficeLength.Points(10))); }).Save(source);
        int calls = 0;
        var catalog = new OcrEngineCatalog().Register(new FixtureProvider(request => { calls++; Assert.Equal(2, request.PageNumber); return Recognized(); }));
        var result = await Run(["ocr", source, "--output", destination, "--pages", "last", "--ocr-provider", "fixture", "--dpi", "72"], catalog);
        Assert.Equal(0, result.Exit); Assert.Equal(1, calls);
        var pages = PdfReadDocument.Open(await File.ReadAllBytesAsync(destination)).Pages;
        Assert.Equal(2, pages.Count); Assert.Contains("Recognized secret", pages[1].ExtractText());
        Assert.DoesNotContain("Recognized secret", result.Output + result.Error);
        using var json = JsonDocument.Parse(result.Output);
        Assert.Equal(1, json.RootElement.GetProperty("addedWordCount").GetInt32());
    }

    [Fact]
    public async Task AgentOcrRechecksTheRootBeforeStagingWhenAnOutputParentIsRedirected() {
        if (OperatingSystem.IsWindows()) return; // Windows symbolic-link creation requires separate host privilege.
        using var scope = new DirectoryScope();
        string allowed = scope.File("allowed"), outside = scope.File("outside");
        Directory.CreateDirectory(allowed); Directory.CreateDirectory(outside);
        string source = Path.Combine(allowed, "scan.pdf"), parent = Path.Combine(allowed, "output");
        PdfDocument.Create(d => d.Page(p => p.Size(global::OfficeIMO.OfficeLength.Points(200), global::OfficeIMO.OfficeLength.Points(100)).Margin(global::OfficeIMO.OfficeLength.Points(10)))).Save(source);
        var catalog = new OcrEngineCatalog().Register(new FixtureProvider(_ => {
            Directory.CreateSymbolicLink(parent, outside); return Recognized();
        }));
        var service = new OfficeImoAgentService(new AgentPathPolicy([allowed]), pdfOcrCatalog: catalog);
        var result = await service.SearchablePdfAsync(source, Path.Combine(parent, "copy.pdf"), "fixture", new() { Dpi = 72 });
        Assert.False(result.Succeeded); Assert.Empty(Directory.GetFileSystemEntries(outside));
        Assert.DoesNotContain("Recognized secret", AgentJson.Serialize(result));
    }

    [Fact]
    public async Task AgentReportsRemainBoundedAndOutsideOutputsAreRejectedBeforeMutation() {
        using var scope = new DirectoryScope();
        string allowed = scope.File("allowed"); Directory.CreateDirectory(allowed);
        string source = Path.Combine(allowed, "source.pdf"); CreatePdf(source, 8);
        var service = new OfficeImoAgentService(new AgentPathPolicy([allowed]));
        var result = await service.SplitPdfAsync(source, Path.Combine(allowed, "parts"), 1, new(), 512);
        Assert.True(result.Succeeded); Assert.Equal(8, result.ArtifactCount); Assert.True(result.Truncated);
        Assert.True(AgentJson.Serialize(result).Length <= 512);
        await Assert.ThrowsAsync<UnauthorizedAccessException>(() => service.PdfAsync(source, scope.File("outside.pdf"), "extract", new() { Pages = "1" }));
        Assert.False(File.Exists(scope.File("outside.pdf")));
        await Assert.ThrowsAsync<AgentUsageException>(() => service.PdfAsync(source, source, "extract", new() { Pages = "1", Overwrite = true }));
        using var cancellation = new CancellationTokenSource(); cancellation.Cancel();
        var cancelled = await service.PdfAsync(source, Path.Combine(allowed, "cancelled.pdf"), "extract", new() { Pages = "1" }, cancellationToken: cancellation.Token);
        Assert.Equal("Cancelled", cancelled.Status); Assert.False(File.Exists(Path.Combine(allowed, "cancelled.pdf")));
    }

    [Fact]
    public void ProviderDiscoveryWithLongEscapedIdsHonorsTheMinimumReportBudget() {
        using var scope = new DirectoryScope();
        var catalog = new OcrEngineCatalog();
        for (int i = 0; i < 30; i++) catalog.Register(new DiscoveryProvider(i.ToString("D2") + new string('é', 254)));
        var service = new OfficeImoAgentService(new AgentPathPolicy([scope.Root]), pdfOcrCatalog: catalog);
        var result = service.PdfOcrProviders(512);
        Assert.Equal(30, result.ProviderCount);
        Assert.Empty(result.ProviderIds);
        Assert.True(result.Truncated);
        Assert.True(AgentJson.Serialize(result).Length <= 512);
        Assert.Throws<AgentUsageException>(() => service.PdfOcrProviders(511));
    }

    [Fact]
    public async Task ProviderFactoryAccessErrorsDoNotExposePrivateMessagesThroughMcp() {
        using var scope = new DirectoryScope();
        string source = scope.File("source.pdf"), destination = scope.File("copy.pdf"); CreatePdf(source, 1);
        var catalog = new OcrEngineCatalog().Register(new RefusingProvider());
        var tools = new OfficeImoMcpTools(new OfficeImoAgentService(new AgentPathPolicy([scope.Root]), pdfOcrCatalog: catalog));
        var result = await tools.SearchablePdfAsync(source, destination, "refusing");
        Assert.True(result.IsError); Assert.False(File.Exists(destination));
        Assert.DoesNotContain("private-provider-message", string.Join(" ", result.Content.OfType<TextContentBlock>().Select(item => item.Text)));
    }

    [Fact]
    public async Task PdfAConversionListsChangesAndARefusalExitsAsUnsupportedWithoutWriting() {
        using var scope = new DirectoryScope();
        string source = scope.File("source.pdf"), destination = scope.File("archive.pdf"), refused = scope.File("refused.pdf");
        await File.WriteAllBytesAsync(source, PdfWithAttachment());

        var converted = await Run(["convert-to-pdfa", source, "--output", destination, "--profile", "2b", "--attachments", "remove"]);
        var refusal = await Run(["convert-to-pdfa", source, "--output", refused, "--attachments", "refuse"]);

        Assert.Equal(0, converted.Exit);
        Assert.Equal(2, PdfReadDocument.Open(await File.ReadAllBytesAsync(destination)).XmpMetadata?.PdfAPart);
        using (var json = JsonDocument.Parse(converted.Output)) {
            JsonElement report = json.RootElement.GetProperty("pdfA");
            Assert.Equal("PDF/A-2b", report.GetProperty("profile").GetString());
            Assert.Contains(report.GetProperty("changes").EnumerateArray(), change => change.GetProperty("code").GetString() == "EmbeddedFileRemoved" &&
                change.GetProperty("objectNumber").GetInt32() == 5);
        }
        Assert.Contains("change EmbeddedFileRemoved", converted.Error, StringComparison.Ordinal);
        Assert.Equal((int)OfficeImoToolExitCode.UnsupportedInput, refusal.Exit);
        Assert.Contains("violation EmbeddedFileNotAllowed", refusal.Error, StringComparison.Ordinal);
        Assert.False(File.Exists(refused));
    }

    [Fact]
    public async Task PdfAConversionThroughMcpReadsFontsOnlyWithinAllowedRoots() {
        using var scope = new DirectoryScope();
        string allowed = Directory.CreateDirectory(scope.File("allowed")).FullName;
        string source = Path.Combine(allowed, "source.pdf"), destination = Path.Combine(allowed, "archive.pdf"), outsideFont = scope.File("font.ttf");
        await File.WriteAllBytesAsync(source, PdfWithAttachment());
        await File.WriteAllBytesAsync(outsideFont, [0, 1, 0, 0]);
        var tools = new OfficeImoMcpTools(new OfficeImoAgentService(new AgentPathPolicy([allowed])));

        var outside = await tools.ConvertToPdfAAsync(source, destination, fonts: ["Archive Sans=" + outsideFont]);
        var converted = await tools.ConvertToPdfAAsync(source, destination, profile: "3b");

        Assert.True(outside.IsError);
        Assert.False(converted.IsError);
        JsonElement report = converted.StructuredContent!.Value.GetProperty("pdfA");
        Assert.Equal("PDF/A-3b", report.GetProperty("profile").GetString());
        Assert.Equal(0, report.GetProperty("violationCount").GetInt32());
        Assert.Single(PdfDocument.Load(destination).Attachments.Extract());
    }

    // A vector-only page with one embedded file, so conversion does not depend on installed fonts.
    internal static byte[] PdfWithAttachment() {
        string[] objects = [
            "<< /Type /Catalog /Pages 2 0 R /Names << /EmbeddedFiles << /Names [(data.xml) 5 0 R] >> >> >>",
            "<< /Type /Pages /Kids [3 0 R] /Count 1 >>",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 200] /Resources << >> /Contents 4 0 R >>",
            "<< /Length 21 >>\nstream\n0 g 20 20 50 50 re f\n\nendstream",
            "<< /Type /Filespec /F (data.xml) /EF << /F 6 0 R >> >>",
            "<< /Type /EmbeddedFile /Length 7 >>\nstream\n<data/>\nendstream"
        ];
        var pdf = new System.Text.StringBuilder("%PDF-1.4\n");
        var offsets = new List<int>();
        for (int index = 0; index < objects.Length; index++) {
            offsets.Add(pdf.Length);
            pdf.Append(index + 1).Append(" 0 obj\n").Append(objects[index]).Append("\nendobj\n");
        }
        int xref = pdf.Length;
        pdf.Append("xref\n0 ").Append(objects.Length + 1).Append("\n0000000000 65535 f \n");
        foreach (int offset in offsets) pdf.Append(offset.ToString("0000000000", System.Globalization.CultureInfo.InvariantCulture)).Append(" 00000 n \n");
        pdf.Append("trailer\n<< /Size ").Append(objects.Length + 1).Append(" /Root 1 0 R >>\nstartxref\n").Append(xref).Append("\n%%EOF\n");
        return System.Text.Encoding.ASCII.GetBytes(pdf.ToString());
    }

    private static async Task<(int Exit, string Output, string Error)> Run(string[] args, OcrEngineCatalog? catalog = null) {
        using var output = new StringWriter(); using var error = new StringWriter();
        int exit = await PdfCommand.RunAsync(args, output, error, ocrCatalog: catalog);
        return (exit, output.ToString(), error.ToString());
    }

    private static OcrResult Recognized() => new() {
        Provider = "fixture", Spans = [new OcrTextSpan { Text = "Recognized secret", Level = OcrTextSpanLevel.Word, Confidence = 1,
            CoordinateUnit = OcrCoordinateUnit.Points, Region = new() { X = 10, Y = 20, Width = 100, Height = 12 } }]
    };
    private sealed class FixtureProvider(Func<OcrRequest, OcrResult> recognize) : IOcrEngineProvider {
        public string Id => "fixture"; public string DisplayName => "Fixture";
        public OcrEngineCapabilities Capabilities => new() { SupportsWordSpans = true };
        public IOcrEngine Create(IReadOnlyDictionary<string, string> options) => new DelegateOcrEngine(Id, (request, _) => Task.FromResult(recognize(request)), Capabilities);
    }
    private sealed class RefusingProvider : IOcrEngineProvider {
        public string Id => "refusing"; public string DisplayName => "Refusing fixture";
        public OcrEngineCapabilities Capabilities => new();
        public IOcrEngine Create(IReadOnlyDictionary<string, string> options) => throw new UnauthorizedAccessException("private-provider-message");
    }
    private sealed class DiscoveryProvider(string id) : IOcrEngineProvider {
        public string Id => id;
        public string DisplayName => "Discovery fixture";
        public OcrEngineCapabilities Capabilities => new();
        public IOcrEngine Create(IReadOnlyDictionary<string, string> options) => throw new InvalidOperationException("Discovery does not execute a provider.");
    }
    internal static void CreatePdf(string path, int count) => PdfDocument.Create(d => {
        for (int index = 1; index <= count; index++) { int page = index;
            d.Page(p => p.Size(global::OfficeIMO.OfficeLength.Points(200 + page), global::OfficeIMO.OfficeLength.Points(100)).Margin(global::OfficeIMO.OfficeLength.Points(10)).Content(c => c.Item(i => i.AddParagraph(t => t.AddText("Page " + page)))));
        }
    }).Save(path);
    private sealed class DirectoryScope : IDisposable {
        internal string Root { get; } = Path.Combine(Path.GetTempPath(), "officeimo-pdf-automation-" + Guid.NewGuid().ToString("N"));
        internal DirectoryScope() => Directory.CreateDirectory(Root);
        internal string File(string name) => Path.Combine(Root, name);
        public void Dispose() => Directory.Delete(Root, recursive: true);
    }
}
