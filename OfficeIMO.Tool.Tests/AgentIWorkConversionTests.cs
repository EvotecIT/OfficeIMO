using System.IO.Compression;
using System.Security.Cryptography;
using ModelContextProtocol.Client;
using ModelContextProtocol.Protocol;
using OfficeIMO.IWork;
using OfficeIMO.Pdf;
using OfficeIMO.Tool.Agent;
using OfficeIMO.Tool.Commands.Agent;
using OfficeIMO.Tool.Mcp;
using Xunit;

namespace OfficeIMO.Tool.Tests;

public sealed partial class AgentIWorkConversionTests {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task Agent_cli_exit_code_matches_the_structured_iwork_workflow_result(bool overwrite) {
        using var files = new Files();
        string output = files.Output("docx");
        File.WriteAllText(output, "keep existing");
        using var stdout = new StringWriter();
        using var stderr = new StringWriter();
        var arguments = new List<string> { "convert", files.Input, "--output", output, "--format", "docx" };
        if (overwrite) arguments.Add("--overwrite");
        int exit = await AgentCommand.RunAsync(arguments.ToArray(), stdout, stderr);
        using var response = System.Text.Json.JsonDocument.Parse(stdout.ToString());
        Assert.Equal(overwrite, response.RootElement.GetProperty("workflow").GetProperty("succeeded").GetBoolean());
        Assert.Equal((int)(overwrite ? OfficeImoToolExitCode.Success : OfficeImoToolExitCode.OperationFailed), exit);
        Assert.Equal(overwrite ? output : string.Empty, response.RootElement.GetProperty("outputPath").GetString());
        if (!overwrite) Assert.Equal("keep existing", File.ReadAllText(output));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task Stdio_iwork_conversion_accepts_typed_policy_and_returns_structured_stage_evidence(bool directoryPackage) {
        using var files = new Files(directoryPackage);
        var transport = new StdioClientTransport(new StdioClientTransportOptions {
            Name = "officeimo-iwork-source-test",
            Command = Environment.GetEnvironmentVariable("DOTNET_HOST_PATH") ?? "dotnet",
            Arguments = [typeof(OfficeImoToolApp).Assembly.Location, "mcp", "serve", "--stdio"],
            WorkingDirectory = files.Root,
            EnvironmentVariables = new Dictionary<string, string?> {
                [AgentPathPolicy.AllowedRootsEnvironmentVariable] = files.Root
            }
        });
        using var timeout = new CancellationTokenSource(TimeSpan.FromSeconds(30));
        await using McpClient client = await McpClient.CreateAsync(transport, cancellationToken: timeout.Token);
        var arguments = new Dictionary<string, object?> {
            ["path"] = files.Input, ["outputPath"] = files.Output("pdf"), ["format"] = "pdf",
            ["iworkOptions"] = new Dictionary<string, object?> { ["mode"] = 1 }
        };
        CallToolResult result = await client.CallToolAsync("officeimo_convert", arguments, cancellationToken: timeout.Token);
        Assert.False(result.IsError, string.Join("; ", result.Content.OfType<TextContentBlock>().Select(item => item.Text)));
        var data = result.StructuredContent!.Value;
        if (!directoryPackage)
            Assert.Equal(Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(files.Input))), data.GetProperty("sourceSha256").GetString(), ignoreCase: true);
        else {
            Assert.Equal(64, data.GetProperty("sourceSha256").GetString()!.Length);
            Assert.Contains("DirectoryPackage", data.GetRawText());
        }
        Assert.Equal("iwork-editable,pdf", data.GetProperty("workflow").GetProperty("conversionEvidence").GetProperty("facts").GetProperty("conversionStages").GetString());
        arguments["overwrite"] = true;
        arguments["iworkOptions"] = new Dictionary<string, object?> { ["allowPartialEditableReconstruction"] = true };
        string outputHash = Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(files.Output("pdf"))));
        CallToolResult rejected = await client.CallToolAsync("officeimo_convert", arguments, cancellationToken: timeout.Token);
        Assert.True(rejected.IsError);
        Assert.NotNull(rejected.StructuredContent);
        Assert.Equal(string.Empty, rejected.StructuredContent.Value.GetProperty("outputPath").GetString());
        Assert.Equal(outputHash, Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(files.Output("pdf")))));
    }

    [Theory]
    [InlineData("docx")]
    [InlineData("pdf")]
    public async Task Mcp_iwork_conversion_uses_the_shared_workflow_and_retains_source_hash(string format) {
        using var files = new Files();
        var service = new OfficeImoAgentService(new AgentPathPolicy([files.Root]));
        AgentConvertResult result = await service.ConvertAsync(files.Input, files.Output(format), format);
        Assert.True(result.Workflow!.Succeeded, result.Workflow.Summary);
        Assert.Equal(Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(files.Input))), result.SourceSha256, ignoreCase: true);
        Assert.Equal(Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(result.OutputPath))), result.Sha256, ignoreCase: true);
        Assert.Equal("EditableReconstruction", result.Workflow.ConversionEvidence!.Facts["projectionKind"]);
        Assert.Contains(result.Workflow.Diagnostics, item => item.Code == "OutputReopened");
        Assert.Contains("sourceSha256", AgentJson.Serialize(result));
        var tools = new OfficeImoMcpTools(service);
        var response = await tools.ConvertAsync(files.Input, files.Output("pdf"), "pdf", overwrite: true,
            iworkOptions: new IWorkConversionOptions { Mode = IWorkConversionMode.EditableOnly });
        Assert.False(response.IsError);
        Assert.NotNull(response.StructuredContent);
        Assert.Contains("hello pages", PdfReadDocument.Open(File.ReadAllBytes(files.Output("pdf"))).ExtractText());
    }

    [Fact]
    public async Task Mcp_output_limit_is_a_structured_failure_and_protects_existing_output() {
        using var files = new Files();
        string output = files.Output("pdf");
        File.WriteAllText(output, "keep existing");
        var tools = new OfficeImoMcpTools(new OfficeImoAgentService(new AgentPathPolicy([files.Root])));
        var response = await tools.ConvertAsync(files.Input, output, "pdf", overwrite: true, maximumOutputBytes: 32);
        Assert.True(response.IsError);
        Assert.NotNull(response.StructuredContent);
        Assert.Equal(string.Empty, response.StructuredContent.Value.GetProperty("outputPath").GetString());
        Assert.Equal(0, response.StructuredContent.Value.GetProperty("lengthBytes").GetInt64());
        Assert.Equal(string.Empty, response.StructuredContent.Value.GetProperty("sha256").GetString());
        Assert.Equal("keep existing", File.ReadAllText(output));
    }

    [Fact]
    public async Task Mcp_preserves_conflicts_and_rejects_output_outside_authorized_roots() {
        using var files = new Files();
        string output = files.Output("docx");
        File.WriteAllText(output, "keep existing");
        var service = new OfficeImoAgentService(new AgentPathPolicy([files.Root]));
        AgentConvertResult result = await service.ConvertAsync(files.Input, output, "docx");
        Assert.False(result.Workflow!.Succeeded);
        Assert.Empty(result.OutputPath);
        Assert.Equal(0, result.LengthBytes);
        Assert.Empty(result.Sha256);
        Assert.Equal("keep existing", File.ReadAllText(output));
        await Assert.ThrowsAsync<UnauthorizedAccessException>(() => service.ConvertAsync(files.Input,
            Path.Combine(Path.GetDirectoryName(files.Root)!, "outside.docx"), "docx"));
    }

    private sealed class Files : IDisposable {
        public string Root { get; } = Path.Combine(Path.GetTempPath(), "officeimo-agent-iwork-" + Guid.NewGuid().ToString("N"));
        public string Input { get; }
        public string Output(string format) => Path.Combine(Root, "output." + format);
        public Files(bool directoryPackage = false, string kind = "pages", bool noteFreeKeynote = false, bool completeNumbers = false) {
            Directory.CreateDirectory(Root);
            Input = Path.Combine(Root, "source." + kind);
            string fixture = completeNumbers
                ? Path.Combine(AppContext.BaseDirectory, "Documents", "IWorkCorpus", "synthetic", "numbers-unpadded-scalar.numbers")
                : Path.Combine(AppContext.BaseDirectory, "Documents", "IWorkCorpus", "nim-iwork", "simple." + kind);
            if (noteFreeKeynote) {
                IWorkKeynoteDocument source = IWorkKeynoteDocument.Create();
                source.AddSlide().AddText("hello keynote", 40, 40, 600, 180, fontSizePoints: 24);
                source.AddSlide().AddText("Second slide", 40, 40, 600, 180, fontSizePoints: 24);
                if (directoryPackage) {
                    using var bytes = new MemoryStream(source.SaveBytes(), writable: false);
                    using var archive = new ZipArchive(bytes, ZipArchiveMode.Read);
                    archive.ExtractToDirectory(Input);
                } else source.Save(Input);
            } else if (directoryPackage) ZipFile.ExtractToDirectory(fixture, Input);
            else File.Copy(fixture, Input);
        }
        public void Dispose() => Directory.Delete(Root, true);
    }
}
