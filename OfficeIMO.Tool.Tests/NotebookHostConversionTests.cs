using OfficeIMO.Tool.Agent;
using OfficeIMO.Tool.Mcp;
using Xunit;

namespace OfficeIMO.Tool.Tests;

public sealed class NotebookHostConversionTests {
    [Theory]
    [InlineData(".md", false, true)]
    [InlineData(".md", true, false)]
    [InlineData(".markdown", false, false)]
    public async Task CliNotebookMarkdownIncludesSavedPngUnlessReaderIsRequested(string extension,
        bool extractAssets, bool expectSavedPng) {
        using var files = new NotebookFiles();
        string destination = files.Output(extension);
        string[] options = extractAssets ? ["--assets", Path.Combine(files.Root, "assets")] : [];
        using var output = new MemoryStream();
        using var error = new StringWriter();

        int status = await OfficeImoToolApp.RunAsync(["convert", files.Source, destination, .. options],
            Stream.Null, output, error);

        Assert.True(status == (int)OfficeImoToolExitCode.Success, error.ToString());
        string markdown = await File.ReadAllTextAsync(destination);
        Assert.Contains("nbconvert latex test", markdown, StringComparison.Ordinal);
        Assert.Equal(expectSavedPng, markdown.Contains("data:image/png;base64,", StringComparison.Ordinal));
    }

    [Fact]
    public async Task McpNotebookMarkdownPublicationPreservesSavedPngAndHonorsSharedOutputLimit() {
        using var files = new NotebookFiles();
        string destination = files.Output(".md");
        var tools = new OfficeImoMcpTools(new OfficeImoAgentService(new AgentPathPolicy([files.Root])));

        var converted = await tools.ConvertAsync(files.Source, destination);

        Assert.False(converted.IsError);
        Assert.NotNull(converted.StructuredContent);
        Assert.Equal("markdown", converted.StructuredContent.Value.GetProperty("format").GetString());
        Assert.StartsWith("officeimo:", converted.StructuredContent.Value.GetProperty("sourceId").GetString());
        string markdown = await File.ReadAllTextAsync(destination);
        Assert.Contains("data:image/png;base64,", markdown, StringComparison.Ordinal);

        var rejected = await tools.ConvertAsync(files.Source, destination, "md", overwrite: true, maximumOutputBytes: 32);
        Assert.True(rejected.IsError);
        Assert.Equal(markdown, await File.ReadAllTextAsync(destination));
    }

    private sealed class NotebookFiles : IDisposable {
        internal NotebookFiles() {
            Root = Path.Combine(Path.GetTempPath(), "officeimo-notebook-host-" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(Root);
            Source = Path.Combine(Root, "source.ipynb");
            File.Copy(Path.Combine(AppContext.BaseDirectory, "Fixtures", "Notebook", "test4.5.ipynb"), Source);
        }

        internal string Root { get; }
        internal string Source { get; }
        internal string Output(string extension) => Path.Combine(Root, "report" + extension);

        public void Dispose() => Directory.Delete(Root, recursive: true);
    }
}
