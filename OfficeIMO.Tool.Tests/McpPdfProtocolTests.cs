using ModelContextProtocol.Client;
using ModelContextProtocol.Protocol;
using OfficeIMO.Pdf;
using OfficeIMO.Tool.Agent;
using Xunit;

namespace OfficeIMO.Tool.Tests;

public sealed class McpPdfProtocolTests {
    [Fact]
    public async Task StdioPdfToolsCreateVerifiedCopiesAndEnforceRootsAndAcknowledgement() {
        string root = Path.Combine(Path.GetTempPath(), "officeimo-mcp-pdf-" + Guid.NewGuid().ToString("N"));
        string allowed = Path.Combine(root, "allowed"); Directory.CreateDirectory(allowed);
        string source = Path.Combine(allowed, "source.pdf"); PdfAutomationCommandTests.CreatePdf(source, 3);
        string assembly = typeof(OfficeImoToolApp).Assembly.Location;
        var transport = new StdioClientTransport(new() {
            Name = "officeimo-pdf-test", Command = "dotnet", Arguments = [assembly, "mcp", "serve", "--stdio"], WorkingDirectory = allowed,
            EnvironmentVariables = new Dictionary<string, string?> { [AgentPathPolicy.AllowedRootsEnvironmentVariable] = allowed }
        });
        using var timeout = new CancellationTokenSource(TimeSpan.FromSeconds(45));
        try {
            await using McpClient client = await McpClient.CreateAsync(transport, cancellationToken: timeout.Token);
            string extracted = Path.Combine(allowed, "extracted.pdf");
            var copy = await client.CallToolAsync("officeimo_pdf", new Dictionary<string, object?> {
                ["path"] = source, ["outputPath"] = extracted, ["operation"] = "extract", ["pages"] = "last,1", ["maxOutputCharacters"] = 512
            }, cancellationToken: timeout.Token);
            Assert.False(copy.IsError, Text(copy)); Assert.True(copy.StructuredContent!.Value.GetProperty("succeeded").GetBoolean());
            Assert.True(copy.StructuredContent.Value.GetRawText().Length <= 512);
            Assert.Equal(["Page 3", "Page 1"], PdfReadDocument.Open(File.ReadAllBytes(extracted)).Pages.Select(page => page.ExtractText().Trim()));
            var outside = await client.CallToolAsync("officeimo_pdf", new Dictionary<string, object?> {
                ["path"] = source, ["outputPath"] = Path.Combine(root, "outside.pdf"), ["operation"] = "extract", ["pages"] = "1"
            }, cancellationToken: timeout.Token);
            Assert.True(outside.IsError); Assert.False(File.Exists(Path.Combine(root, "outside.pdf")));
            var unacknowledged = await client.CallToolAsync("officeimo_pdf", new Dictionary<string, object?> {
                ["path"] = source, ["outputPath"] = Path.Combine(allowed, "raster.pdf"), ["operation"] = "flatten"
            }, cancellationToken: timeout.Token);
            Assert.True(unacknowledged.IsError); Assert.False(File.Exists(Path.Combine(allowed, "raster.pdf")));
            var raster = await client.CallToolAsync("officeimo_pdf", new Dictionary<string, object?> {
                ["path"] = source, ["outputPath"] = Path.Combine(allowed, "raster.pdf"), ["operation"] = "flatten",
                ["acknowledgeRasterOutput"] = true, ["pages"] = "last", ["dpi"] = 72
            }, cancellationToken: timeout.Token);
            Assert.False(raster.IsError, Text(raster)); Assert.True(string.IsNullOrWhiteSpace(PdfReadDocument.Open(File.ReadAllBytes(Path.Combine(allowed, "raster.pdf"))).Pages[0].ExtractText()));
            var split = await client.CallToolAsync("officeimo_pdf_split", new Dictionary<string, object?> {
                ["path"] = source, ["outputDirectory"] = Path.Combine(allowed, "parts"), ["pagesPerDocument"] = 2
            }, cancellationToken: timeout.Token);
            Assert.False(split.IsError, Text(split)); Assert.Equal(2, split.StructuredContent!.Value.GetProperty("artifactCount").GetInt32());
            var exported = await client.CallToolAsync("officeimo_pdf_export_pages", new Dictionary<string, object?> {
                ["path"] = source, ["outputDirectory"] = Path.Combine(allowed, "images"), ["pages"] = "1", ["dpi"] = 72
            }, cancellationToken: timeout.Token);
            Assert.False(exported.IsError, Text(exported)); Assert.Single(Directory.GetFiles(Path.Combine(allowed, "images"), "*.png"));
            var limitedExport = await client.CallToolAsync("officeimo_pdf_export_pages", new Dictionary<string, object?> {
                ["path"] = source, ["outputDirectory"] = Path.Combine(allowed, "limited-images"), ["maximumPages"] = 1
            }, cancellationToken: timeout.Token);
            Assert.True(limitedExport.IsError); Assert.False(Directory.Exists(Path.Combine(allowed, "limited-images")));
            var assembled = await client.CallToolAsync("officeimo_pdf_assemble", new Dictionary<string, object?> {
                ["paths"] = new[] { extracted, source }, ["outputPath"] = Path.Combine(allowed, "assembled.pdf")
            }, cancellationToken: timeout.Token);
            Assert.False(assembled.IsError, Text(assembled)); Assert.Equal(5, PdfDocument.Load(Path.Combine(allowed, "assembled.pdf")).Inspect().PageCount);
            var plan = await client.CallToolAsync("officeimo_pdf_print_plan", new Dictionary<string, object?> {
                ["path"] = source, ["pages"] = "last,1", ["pagesPerSheet"] = 2
            }, cancellationToken: timeout.Token);
            Assert.False(plan.IsError, Text(plan)); Assert.Equal(1, plan.StructuredContent!.Value.GetProperty("sheetCount").GetInt32());
            var providers = await client.CallToolAsync("officeimo_pdf_ocr_providers", cancellationToken: timeout.Token);
            Assert.False(providers.IsError, Text(providers)); Assert.Equal(0, providers.StructuredContent!.Value.GetProperty("providerCount").GetInt32());
        } finally { Directory.Delete(root, recursive: true); }
    }

    private static string Text(CallToolResult result) => string.Join(" | ", result.Content.OfType<TextContentBlock>().Select(item => item.Text));
}
