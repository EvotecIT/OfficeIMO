using System.Text;
using System.Text.Json;
using ModelContextProtocol.Client;
using ModelContextProtocol.Protocol;
using OfficeIMO.Drawing;
using OfficeIMO.Tool.Agent;
using OfficeIMO.Provenance;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tool.Tests;

public sealed class HiddenDataCommandTests {
    [Theory]
    [InlineData("remove", "--anonymous-author", "--help")]
    [InlineData("inspect", "--format", "--help")]
    [InlineData("remove", "--anonymous-author", "-h")]
    public async Task HelpBeforeAnOptionValuePrintsUsageWithoutWritingOutput(string operation, string option, string help) {
        string directory = Path.Combine(Path.GetTempPath(), "officeimo-hidden-data-help-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        try {
            string input = Path.Combine(directory, "input.docx");
            using (WordDocument document = WordDocument.Create(input)) {
                document.Properties.Creator = "Private Author";
                document.AddParagraph("Body");
                document.Save();
            }
            byte[] original = File.ReadAllBytes(input);
            using var output = new MemoryStream();
            using var error = new StringWriter();

            int exitCode = await OfficeImoToolApp.RunAsync(["hidden-data", operation, input, option, help], Stream.Null, output, error);

            Assert.Equal((int)OfficeImoToolExitCode.Success, exitCode);
            Assert.Contains("officeimo hidden-data", Encoding.UTF8.GetString(output.ToArray()), StringComparison.Ordinal);
            Assert.Equal(string.Empty, error.ToString());
            Assert.Equal(original, File.ReadAllBytes(input));
            Assert.Equal(new[] { input }, Directory.GetFiles(directory));
        } finally {
            Directory.Delete(directory, recursive: true);
        }
    }

    [Fact]
    public async Task CliAndMcpWriteACleanCopyAndKeepProvenanceUnlessItIsSelected() {
        string directory = Path.Combine(Path.GetTempPath(), "officeimo-tool-hidden-data-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        try {
            string input = Path.Combine(directory, "shared.docx");
            string image = Path.Combine(directory, "generated.jpg");
            byte[] jpeg = OfficeJpegCodec.Encode(new OfficeRasterImage(1, 1, OfficeColor.White));
            OfficeImageMetadata metadata = OfficeImageMetadata.Read(jpeg);
            metadata.XmpProfile = Encoding.UTF8.GetBytes(
                    "<x:xmpmeta xmlns:x=\"adobe:ns:meta/\"><rdf:RDF xmlns:rdf=\"http://www.w3.org/1999/02/22-rdf-syntax-ns#\">" +
                    "<rdf:Description rdf:about=\"\" xmlns:Iptc4xmpExt=\"http://iptc.org/std/Iptc4xmpExt/2008-02-29/\">" +
                    "<Iptc4xmpExt:DigitalSourceType rdf:resource=\"http://cv.iptc.org/newscodes/digitalsourcetype/trainedAlgorithmicMedia\"/>" +
                    "</rdf:Description></rdf:RDF></x:xmpmeta>");
            File.WriteAllBytes(image, OfficeImageMetadata.Apply(jpeg, metadata));
            using (WordDocument document = WordDocument.Create(input)) {
                document.Properties.Creator = "Private Author";
                document.AddParagraph().AddImage(image);
                document.Save();
            }
            byte[] original = File.ReadAllBytes(input);

            (int exitCode, string output) kept = await RunAsync("hidden-data", "remove", input);
            Assert.Equal(0, kept.exitCode);
            using (JsonDocument json = JsonDocument.Parse(kept.output)) {
                JsonElement[] findings = json.RootElement.GetProperty("findings").EnumerateArray().ToArray();
                Assert.Contains(findings, finding => finding.GetProperty("kind").GetString() == "DocumentProperties" && finding.GetProperty("outcome").GetString() == "Removed");
                JsonElement provenance = Assert.Single(findings, finding => finding.GetProperty("kind").GetString() == "Provenance");
                Assert.Equal("Retained", provenance.GetProperty("outcome").GetString());
                Assert.Equal("TrainedAlgorithmicMedia", provenance.GetProperty("digitalSourceKind").GetString());
            }
            string cleaned = Path.Combine(directory, "shared.cleaned.docx");
            Assert.True(OfficeProvenanceInspector.InspectFile(cleaned).HasGenerativeAiDeclaration);
            Assert.Equal(original, File.ReadAllBytes(input));

            Assert.Equal(6, (await RunAsync("hidden-data", "remove", input)).ExitCode);
            Assert.Equal(0, (await RunAsync("hidden-data", "remove", input, "--force", "--kinds", "default,provenance")).ExitCode);
            Assert.Empty(OfficeProvenanceInspector.InspectFile(cleaned).Evidence);
            Assert.Equal(2, (await RunAsync("hidden-data", "remove", input, "--kinds", "everything")).ExitCode);

            var transport = new StdioClientTransport(new() {
                Name = "officeimo-hidden-data-test", Command = "dotnet", Arguments = [typeof(OfficeImoToolApp).Assembly.Location, "mcp", "serve", "--stdio"],
                WorkingDirectory = directory,
                EnvironmentVariables = new Dictionary<string, string?> { [AgentPathPolicy.AllowedRootsEnvironmentVariable] = directory }
            });
            using var timeout = new CancellationTokenSource(TimeSpan.FromSeconds(60));
            await using McpClient client = await McpClient.CreateAsync(transport, cancellationToken: timeout.Token);
            string mcpOutput = Path.Combine(directory, "mcp.docx");
            var arguments = new Dictionary<string, object?> { ["path"] = input, ["outputPath"] = mcpOutput, ["kinds"] = new[] { "default", "provenance" } };
            CallToolResult removed = await client.CallToolAsync("officeimo_remove_hidden_data", arguments, cancellationToken: timeout.Token);
            Assert.False(removed.IsError);
            JsonElement finding = removed.StructuredContent!.Value.GetProperty("findings").EnumerateArray()
                .Single(item => item.GetProperty("kind").GetString() == "Provenance");
            Assert.Equal("Removed", finding.GetProperty("outcome").GetString());
            Assert.Empty(OfficeProvenanceInspector.InspectFile(mcpOutput).Evidence);
            Assert.True((await client.CallToolAsync("officeimo_remove_hidden_data", arguments, cancellationToken: timeout.Token)).IsError);
        } finally {
            Directory.Delete(directory, recursive: true);
        }
    }

    private static async Task<(int ExitCode, string Output)> RunAsync(params string[] args) {
        await using var output = new MemoryStream();
        using var error = new StringWriter();
        int exitCode = await OfficeImoToolApp.RunAsync(args, Stream.Null, output, error);
        return (exitCode, Encoding.UTF8.GetString(output.ToArray()) + error);
    }
}
