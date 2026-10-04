using ModelContextProtocol.Client;
using ModelContextProtocol.Protocol;
using OfficeIMO.Tool.Agent;
using Xunit;

namespace OfficeIMO.Tool.Tests;

public sealed class McpProtocolTests {
    [Fact]
    public async Task StdioServerEnforcesConfiguredFilesystemRoot() {
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
            string mailbox = Path.Combine(allowedRoot, "mail");
            Directory.CreateDirectory(mailbox);
            await File.WriteAllTextAsync(Path.Combine(mailbox, "message.eml"), "Subject: unrelated\r\n\r\nsemantic body needle");
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
                    [AgentPathPolicy.AllowedRootsEnvironmentVariable] = allowedRoot
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
            var searched = await client.CallToolAsync("officeimo_search_email", new Dictionary<string, object?> {
                ["path"] = mailbox, ["query"] = "body needle", ["fields"] = "TextBody", ["maxItemsScanned"] = 1
            }, cancellationToken: timeout.Token);
            Assert.False(searched.IsError);
            var page = searched.StructuredContent!.Value;
            Assert.Equal(1, page.GetProperty("itemsScanned").GetInt32());
            Assert.True(page.GetProperty("isComplete").GetBoolean());
            Assert.Equal("TextBody", page.GetProperty("results")[0].GetProperty("matchedFields").GetString());
            var fetched = await client.CallToolAsync("officeimo_fetch", new Dictionary<string, object?> {
                ["sourceId"] = page.GetProperty("sourceId").GetString(), ["id"] = page.GetProperty("results")[0].GetProperty("id").GetString()
            }, cancellationToken: timeout.Token);
            Assert.False(fetched.IsError);
            Assert.Contains("semantic body needle", fetched.StructuredContent!.Value.GetProperty("content").GetString());
            var denied = await client.CallToolAsync("officeimo_search_email", new Dictionary<string, object?> {
                ["path"] = outsideRoot, ["query"] = "body needle"
            }, cancellationToken: timeout.Token);
            Assert.True(denied.IsError);
            var inspectedEmail = await client.CallToolAsync("officeimo_inspect_email", new Dictionary<string, object?> {
                ["path"] = Path.Combine(mailbox, "message.eml"), ["maxOutputCharacters"] = 512
            }, cancellationToken: timeout.Token);
            Assert.False(inspectedEmail.IsError);
            Assert.Equal("EmailDocument", inspectedEmail.StructuredContent!.Value.GetProperty("kind").GetString());
            Assert.Equal("Unverified", inspectedEmail.StructuredContent.Value.GetProperty("signatureStatus").GetString());
            var inspectedOutside = await client.CallToolAsync("officeimo_inspect_email", new Dictionary<string, object?> {
                ["path"] = outsidePath
            }, cancellationToken: timeout.Token);
            Assert.True(inspectedOutside.IsError);
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
            WorkingDirectory = Path.GetDirectoryName(assemblyPath),
            EnvironmentVariables = new Dictionary<string, string?> {
                [AgentPathPolicy.AllowedRootsEnvironmentVariable] = Path.GetDirectoryName(assemblyPath)
            }
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
                "officeimo_inspect_email",
                "officeimo_search",
                "officeimo_search_email"
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
