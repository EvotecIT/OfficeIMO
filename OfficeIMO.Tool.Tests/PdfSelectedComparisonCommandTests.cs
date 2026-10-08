using ModelContextProtocol.Client;
using ModelContextProtocol.Protocol;
using OfficeIMO.Pdf;
using OfficeIMO.Tool.Agent;
using Xunit;

namespace OfficeIMO.Tool.Tests;

public sealed class PdfSelectedComparisonCommandTests {
    [Fact]
    public async Task CliAndStdioMcpCompareRangesThroughSharedWorkflowAndRefuseStaleSourcesAndAliases() {
        string root = Path.Combine(Path.GetTempPath(), "officeimo-pdf-compare-tool-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(root);
        try {
            string first = Path.Combine(root, "first.pdf"), second = Path.Combine(root, "second.pdf");
            PdfDocument.Create(document => { for (int i = 0; i < 4; i++) document.Page(page => page.Size(150, 180)); }).Save(first);
            File.Copy(first, second); byte[] before = File.ReadAllBytes(first);
            string cli = Path.Combine(root, "cli.html");
            using var stdout = new MemoryStream(); var stderr = new StringWriter();
            int code = await OfficeImoToolApp.RunAsync(["workflow", "compare", first, second, "--output", cli,
                "--expected-pages", "4,2", "--actual-pages", "3,2,1"], Stream.Null, stdout, stderr);
            Assert.Equal(0, code); Assert.Contains("Expected 4 / actual 3", File.ReadAllText(cli));
            var transport = new StdioClientTransport(new() {
                Name = "officeimo-compare-test", Command = "dotnet", Arguments = [typeof(OfficeImoToolApp).Assembly.Location, "mcp", "serve", "--stdio"],
                WorkingDirectory = root,
                EnvironmentVariables = new Dictionary<string, string?> { [AgentPathPolicy.AllowedRootsEnvironmentVariable] = root }
            });
            using var timeout = new CancellationTokenSource(TimeSpan.FromSeconds(60));
            await using var client = await McpClient.CreateAsync(transport, cancellationToken: timeout.Token);
            var inspectedExpected = await client.CallToolAsync("officeimo_inspect", new Dictionary<string, object?> { ["path"] = first }, cancellationToken: timeout.Token);
            var inspectedActual = await client.CallToolAsync("officeimo_inspect", new Dictionary<string, object?> { ["path"] = second }, cancellationToken: timeout.Token);
            string expectedId = inspectedExpected.StructuredContent!.Value.GetProperty("sourceId").GetString()!;
            string actualId = inspectedActual.StructuredContent!.Value.GetProperty("sourceId").GetString()!;
            string mcp = Path.Combine(root, "mcp.html");
            Dictionary<string, object?> arguments = new() {
                ["sourceId"] = expectedId, ["comparisonSourceId"] = actualId, ["outputPath"] = mcp,
                ["expectedPages"] = "4,2", ["actualPages"] = "3,2,1", ["maxOutputCharacters"] = 1000
            };
            var result = await client.CallToolAsync("officeimo_pdf_compare", arguments, cancellationToken: timeout.Token);
            Assert.False(result.IsError, Text(result)); Assert.True(File.Exists(mcp));
            Assert.True(result.StructuredContent!.Value.GetRawText().Length <= 1000);
            Assert.DoesNotContain("data:image", Text(result));
            Assert.Contains("Unmatched actual page 1", File.ReadAllText(mcp));
            string alias = Path.Combine(root, "source-alias.html"); File.CreateSymbolicLink(alias, second);
            arguments["outputPath"] = alias; arguments["overwrite"] = true;
            Assert.True((await client.CallToolAsync("officeimo_pdf_compare", arguments, cancellationToken: timeout.Token)).IsError);
            Assert.Equal(before, File.ReadAllBytes(first)); Assert.Equal(before, File.ReadAllBytes(second));
            PdfDocument.Create(document => document.Page(page => page.Size(160, 190))).Save(second);
            string stale = Path.Combine(root, "stale.html"); arguments["outputPath"] = stale;
            Assert.True((await client.CallToolAsync("officeimo_pdf_compare", arguments, cancellationToken: timeout.Token)).IsError);
            Assert.False(File.Exists(stale));
            var refreshedActual = await client.CallToolAsync("officeimo_inspect", new Dictionary<string, object?> { ["path"] = second }, cancellationToken: timeout.Token);
            arguments["comparisonSourceId"] = refreshedActual.StructuredContent!.Value.GetProperty("sourceId").GetString()!;
            PdfDocument.Create(document => document.Page(page => page.Size(170, 200))).Save(first);
            Assert.True((await client.CallToolAsync("officeimo_pdf_compare", arguments, cancellationToken: timeout.Token)).IsError);
            Assert.False(File.Exists(stale));
        } finally { Directory.Delete(root, true); }
    }
    private static string Text(CallToolResult result) => string.Join("\n", result.Content.OfType<TextContentBlock>().Select(block => block.Text));
}
