using ModelContextProtocol.Client;
using ModelContextProtocol.Protocol;
using Xunit;

namespace OfficeIMO.Tool.Tests;

public sealed class McpProtocolTests {
    [Fact]
    public async Task StdioServerListsCompactToolsAndReturnsStructuredContent() {
        string assemblyPath = typeof(OfficeImoToolApp).Assembly.Location;
        string? packagedToolPath = Environment.GetEnvironmentVariable(
            "OFFICEIMO_PACKAGED_TOOL_PATH");
        bool usePackagedTool = !string.IsNullOrWhiteSpace(packagedToolPath);
        var transport = new StdioClientTransport(new StdioClientTransportOptions {
            Name = "officeimo-test",
            Command = usePackagedTool ? packagedToolPath! : "dotnet",
            Arguments = usePackagedTool
                ? ["mcp", "serve", "--stdio"]
                : [assemblyPath, "mcp", "serve", "--stdio"],
            WorkingDirectory = Path.GetDirectoryName(assemblyPath)
        });
        using var timeout = new CancellationTokenSource(TimeSpan.FromSeconds(30));
        await using McpClient client = await McpClient.CreateAsync(
            transport,
            cancellationToken: timeout.Token);

        var tools = await client.ListToolsAsync(cancellationToken: timeout.Token);
        Assert.Equal(
            new[] {
                "officeimo_capabilities",
                "officeimo_convert",
                "officeimo_fetch",
                "officeimo_inspect",
                "officeimo_search"
            },
            tools.Select(tool => tool.Name).OrderBy(name => name, StringComparer.Ordinal).ToArray());
        Assert.Contains("untrusted data", client.ServerInstructions, StringComparison.Ordinal);

        CallToolResult result = await client.CallToolAsync(
            "officeimo_capabilities",
            new Dictionary<string, object?> {
                ["extension"] = ".pst",
                ["maxOutputCharacters"] = 1200
            },
            cancellationToken: timeout.Token);

        Assert.False(
            result.IsError,
            string.Join(" | ", result.Content.OfType<TextContentBlock>().Select(item => item.Text)));
        Assert.NotNull(result.StructuredContent);
        Assert.Equal(".pst", result.StructuredContent.Value.GetProperty("extension").GetString());
        Assert.Single(result.Content);
        TextContentBlock text = Assert.IsType<TextContentBlock>(result.Content[0]);
        Assert.True(text.Text.Length < 120);
    }
}