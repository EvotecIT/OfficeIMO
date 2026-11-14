using ModelContextProtocol.Client;
using ModelContextProtocol.Protocol;
using OfficeIMO.Pdf;
using OfficeIMO.Tool.Agent;
using Xunit;

namespace OfficeIMO.Tool.Tests;

public sealed class PdfRedlineCommandTests {
    private const string Sentence = "Northern offices grew faster than expected after the spring campaign.";

    [Fact]
    public async Task CliWritesTheAnnotatedPdfChangeListAndReportWithoutEchoingText() {
        string root = Path.Combine(Path.GetTempPath(), "officeimo-pdf-redline-cli-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(root);
        try {
            string first = Write(root, "v1.pdf", Sentence), second = Write(root, "v2.pdf", Sentence.Replace("faster", "slower"));
            string output = Path.Combine(root, "redline.pdf"), changes = Path.Combine(root, "changes.json"), report = Path.Combine(root, "report.html");
            using var stdout = new MemoryStream(); var stderr = new StringWriter();

            int code = await OfficeImoToolApp.RunAsync(["pdf", "redline", first, second, "--output", output, "--changes", changes, "--report", report],
                Stream.Null, stdout, stderr);

            Assert.True(code == 0, stderr.ToString());
            string json = System.Text.Encoding.UTF8.GetString(stdout.ToArray());
            Assert.Contains("1 change(s): 0 inserted, 0 deleted, 1 replaced, 0 moved; 1 annotation(s) added.", json, StringComparison.Ordinal);
            Assert.DoesNotContain("faster", json, StringComparison.Ordinal);
            Assert.Single(PdfDocument.Load(File.ReadAllBytes(output)).Inspect().Annotations, static annotation => annotation.Subtype == "Highlight");
            Assert.Contains("\"oldText\": \"faster\"", File.ReadAllText(changes), StringComparison.Ordinal);
            Assert.Contains("<del>faster</del>", File.ReadAllText(report), StringComparison.Ordinal);

            code = await OfficeImoToolApp.RunAsync(["pdf", "redline", first, second, "--output", output], Stream.Null, new MemoryStream(), stderr);
            Assert.Equal(2, code);
            Assert.Contains("Output already exists", stderr.ToString(), StringComparison.Ordinal);
        } finally { Directory.Delete(root, true); }
    }

    [Fact]
    public async Task CliRefusesAPdfWithoutATextLayerAndWritesNothing() {
        string root = Path.Combine(Path.GetTempPath(), "officeimo-pdf-redline-blank-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(root);
        try {
            string first = Write(root, "v1.pdf", Sentence), blank = Path.Combine(root, "scan.pdf");
            PdfDocument.Create().Save(blank);
            string output = Path.Combine(root, "redline.pdf");
            var stderr = new StringWriter();

            int code = await OfficeImoToolApp.RunAsync(["pdf", "redline", first, blank, "--output", output], Stream.Null, new MemoryStream(), stderr);

            Assert.Equal(2, code);
            Assert.Contains("The new PDF has no extractable text", stderr.ToString(), StringComparison.Ordinal);
            Assert.False(File.Exists(output));
        } finally { Directory.Delete(root, true); }
    }

    [Fact]
    public async Task StdioMcpRedlinesRegisteredSourcesAndReturnsCountsOnly() {
        string root = Path.Combine(Path.GetTempPath(), "officeimo-pdf-redline-mcp-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(root);
        try {
            string first = Write(root, "v1.pdf", Sentence), second = Write(root, "v2.pdf", Sentence.Replace("the spring", "the long spring"));
            var transport = new StdioClientTransport(new() {
                Name = "officeimo-redline-test", Command = "dotnet", Arguments = [typeof(OfficeImoToolApp).Assembly.Location, "mcp", "serve", "--stdio"],
                WorkingDirectory = root,
                EnvironmentVariables = new Dictionary<string, string?> { [AgentPathPolicy.AllowedRootsEnvironmentVariable] = root }
            });
            using var timeout = new CancellationTokenSource(TimeSpan.FromSeconds(60));
            await using var client = await McpClient.CreateAsync(transport, cancellationToken: timeout.Token);
            string Id(CallToolResult inspected) => inspected.StructuredContent!.Value.GetProperty("sourceId").GetString()!;
            string oldId = Id(await client.CallToolAsync("officeimo_inspect", new Dictionary<string, object?> { ["path"] = first }, cancellationToken: timeout.Token));
            string newId = Id(await client.CallToolAsync("officeimo_inspect", new Dictionary<string, object?> { ["path"] = second }, cancellationToken: timeout.Token));
            string output = Path.Combine(root, "redline.pdf"), changes = Path.Combine(root, "changes.json");

            CallToolResult result = await client.CallToolAsync("officeimo_pdf_redline", new Dictionary<string, object?> {
                ["sourceId"] = oldId, ["comparisonSourceId"] = newId, ["outputPath"] = output, ["changesPath"] = changes
            }, cancellationToken: timeout.Token);

            Assert.False(result.IsError, Text(result));
            Assert.Contains("1 inserted", Text(result), StringComparison.Ordinal);
            Assert.DoesNotContain("long", result.StructuredContent!.Value.GetRawText(), StringComparison.Ordinal);
            Assert.Equal(2, result.StructuredContent!.Value.GetProperty("artifactCount").GetInt32());
            Assert.Single(PdfDocument.Load(File.ReadAllBytes(output)).Inspect().Annotations, static annotation => annotation.Subtype == "Highlight");
            Assert.Contains("\"newText\": \"long\"", File.ReadAllText(changes), StringComparison.Ordinal);

            CallToolResult intoSource = await client.CallToolAsync("officeimo_pdf_redline", new Dictionary<string, object?> {
                ["sourceId"] = oldId, ["comparisonSourceId"] = newId, ["outputPath"] = second, ["overwrite"] = true
            }, cancellationToken: timeout.Token);
            Assert.True(intoSource.IsError);
            Assert.Contains("separate from every source", Text(intoSource), StringComparison.Ordinal);
        } finally { Directory.Delete(root, true); }
    }

    private static string Write(string root, string name, string paragraph) {
        string path = Path.Combine(root, name);
        PdfDocument.Create().Paragraph(builder => builder.AddText("Quarterly report for all regions.")).Paragraph(builder => builder.AddText(paragraph)).Save(path);
        return path;
    }

    private static string Text(CallToolResult result) => string.Join("\n", result.Content.OfType<TextContentBlock>().Select(block => block.Text));
}
