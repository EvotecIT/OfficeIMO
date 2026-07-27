using ModelContextProtocol.Client;
using ModelContextProtocol.Protocol;
using OfficeIMO.Tool.Agent;
using Xunit;

namespace OfficeIMO.Tool.Tests;

public sealed class McpProtocolTests {
    [Fact]
    public async Task StdioServerDefaultsFilesystemAccessToItsWorkingDirectory() {
        string testRoot = Path.Combine(
            Path.GetTempPath(),
            "officeimo-mcp-root-" + Guid.NewGuid().ToString("N"));
        string allowedRoot = Path.Combine(testRoot, "workspace");
        string outsideRoot = Path.Combine(testRoot, "outside");
        Directory.CreateDirectory(allowedRoot);
        Directory.CreateDirectory(outsideRoot);
        string allowedPath = Path.Combine(allowedRoot, "allowed.md");
        string outsidePath = Path.Combine(outsideRoot, "outside.md");

        try {
            await File.WriteAllTextAsync(allowedPath, "# Allowed");
            await File.WriteAllTextAsync(outsidePath, "# Outside");
            string assemblyPath = typeof(OfficeImoToolApp).Assembly.Location;
            string? packagedToolPath = Environment.GetEnvironmentVariable(
                "OFFICEIMO_PACKAGED_TOOL_PATH");
            bool usePackagedTool = !string.IsNullOrWhiteSpace(packagedToolPath);
            var transport = new StdioClientTransport(new StdioClientTransportOptions {
                Name = "officeimo-root-test",
                Command = usePackagedTool ? packagedToolPath! : "dotnet",
                Arguments = usePackagedTool
                    ? ["mcp", "serve", "--stdio"]
                    : [assemblyPath, "mcp", "serve", "--stdio"],
                WorkingDirectory = allowedRoot,
                EnvironmentVariables = new Dictionary<string, string?> {
                    [AgentPathPolicy.AllowedRootsEnvironmentVariable] = null
                }
            });
            using var timeout = new CancellationTokenSource(TimeSpan.FromSeconds(30));
            await using McpClient client = await McpClient.CreateAsync(
                transport,
                cancellationToken: timeout.Token);

            CallToolResult allowed = await client.CallToolAsync(
                "officeimo_inspect",
                new Dictionary<string, object?> { ["path"] = allowedPath },
                cancellationToken: timeout.Token);
            CallToolResult outside = await client.CallToolAsync(
                "officeimo_inspect",
                new Dictionary<string, object?> { ["path"] = outsidePath },
                cancellationToken: timeout.Token);

            Assert.False(allowed.IsError);
            Assert.True(outside.IsError);
            Assert.Contains(
                AgentPathPolicy.AllowedRootsEnvironmentVariable,
                string.Join(" | ", outside.Content.OfType<TextContentBlock>().Select(item => item.Text)),
                StringComparison.Ordinal);
        } finally {
            if (Directory.Exists(testRoot)) Directory.Delete(testRoot, recursive: true);
        }
    }

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
