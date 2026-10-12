using OfficeIMO.IWork;
using OfficeIMO.Pdf;
using OfficeIMO.Tool.Agent;
using OfficeIMO.Tool.Commands.Reader;
using OfficeIMO.Tool.Mcp;
using Xunit;

namespace OfficeIMO.Tool.Tests;

public sealed partial class AgentIWorkConversionTests {
    [Theory]
    [InlineData("pages", "docx")]
    [InlineData("pages", "pdf")]
    [InlineData("numbers", "xlsx")]
    [InlineData("numbers", "pdf")]
    [InlineData("key", "pptx")]
    [InlineData("key", "pdf")]
    public async Task Authorized_directory_packages_use_all_six_shared_routes_and_captured_source_identity(string kind, string format) {
        using var files = new Files(true, kind, noteFreeKeynote: kind == "key" && format == "pdf", completeNumbers: kind == "numbers");
        var registry = new AgentSourceRegistry();
        var service = new OfficeImoAgentService(new AgentPathPolicy([files.Root]), registry);
        AgentConvertResult result = await service.ConvertAsync(files.Input, files.Output(format), format);
        Assert.True(result.Workflow!.Succeeded, result.Workflow.Summary);
        string contentHash = IWorkSourceDocument.Open(files.Input).ComputePackageContentHash();
        Assert.Equal(contentHash, result.Workflow.ConversionEvidence!.Facts["packageContentSha256"]);
        var source = registry.Resolve(result.SourceId);
        Assert.True(source.IsDirectory);
        Assert.Equal(files.Input, source.Path);
        Assert.Contains(result.Workflow.Diagnostics, item => item.Code == "SourceSnapshot" && item.Details["snapshotKind"] == "DirectoryPackage");
        Assert.Contains(result.Workflow.Diagnostics, item => item.Code == "OutputReopened");
        Assert.Equal(64, result.SourceSha256!.Length);
        Assert.NotEqual(contentHash, result.SourceSha256); // The workflow hashes its transport ZIP.
        if (format == "pdf") Assert.Equal(kind == "key" ? 2 : 1, PdfReadDocument.Open(File.ReadAllBytes(result.OutputPath)).Pages.Count);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task Directory_inspect_search_fetch_share_ids_and_reject_changed_members(bool freshService) {
        using var files = new Files(true);
        string member = Path.Combine(files.Input, "extra.bin");
        File.WriteAllText(member, "baseline");
        var registry = new AgentSourceRegistry();
        var service = new OfficeImoAgentService(new AgentPathPolicy([files.Root]), registry);
        var conversion = await service.ConvertAsync(files.Input, files.Output("docx"), "docx");
        var inspected = await service.InspectAsync(files.Input);
        Assert.Equal(conversion.SourceId, inspected.SourceId);
        var search = await service.SearchAsync(files.Input, "hello");
        var hit = Assert.Single(search.Results);
        var consumer = freshService ? new OfficeImoAgentService(new AgentPathPolicy([files.Root])) : service;
        var fetched = await consumer.FetchAsync(search.SourceId, hit.Id, sourcePath: freshService ? files.Input : null);
        Assert.Contains("hello pages", AgentJson.Serialize(fetched));
        DateTime rootTime = Directory.GetLastWriteTimeUtc(files.Input);
        DateTime memberTime = File.GetLastWriteTimeUtc(member);
        File.WriteAllText(member, "mutation");
        File.SetLastWriteTimeUtc(member, memberTime);
        Directory.SetLastWriteTimeUtc(files.Input, rootTime);
        Assert.Throws<AgentUsageException>(() => registry.Resolve(conversion.SourceId));
        await Assert.ThrowsAsync<AgentUsageException>(() => consumer.FetchAsync(search.SourceId, hit.Id,
            sourcePath: freshService ? files.Input : null));
    }

    [Fact]
    public async Task Directory_conversion_rejects_package_members_outputs_and_unauthorized_roots() {
        using var files = new Files(true);
        var service = new OfficeImoAgentService(new AgentPathPolicy([files.Root]));
        string member = Path.Combine(files.Input, "Index", "Document.iwa");
        byte[] original = File.ReadAllBytes(member);
        await Assert.ThrowsAsync<ReaderToolOutputException>(() => service.ConvertAsync(files.Input,
            Path.Combine(files.Input, "output.docx"), "docx", overwrite: true));
        Assert.Equal(original, File.ReadAllBytes(member));
        using var denied = new Files(true);
        await Assert.ThrowsAsync<UnauthorizedAccessException>(() => service.ConvertAsync(denied.Input, files.Output("docx"), "docx"));
        string output = files.Output("docx"); File.WriteAllText(output, "preserve");
        var rejected = await service.ConvertAsync(files.Input, output, "docx", overwrite: true, maximumInputBytes: 32);
        Assert.False(rejected.Workflow!.Succeeded);
        Assert.Equal("preserve", File.ReadAllText(output));
        Assert.Empty(rejected.SourceId); // Capture failed before a document identity was available.
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task Linked_directory_members_are_rejected_even_when_their_targets_are_allowed(bool directoryLink) {
        using var files = new Files(true);
        using var other = new Files(true);
        string link = Path.Combine(files.Input, "linked");
        if (OperatingSystem.IsWindows()) return; // Link creation requires an opted-in Windows environment.
        if (directoryLink) Directory.CreateSymbolicLink(link, other.Input);
        else File.CreateSymbolicLink(link, Path.Combine(other.Input, "Index", "Document.iwa"));
        var tools = new OfficeImoMcpTools(new OfficeImoAgentService(new AgentPathPolicy([files.Root, other.Root])));
        var result = await tools.ConvertAsync(files.Input, files.Output("docx"), "docx");
        Assert.True(result.IsError);
        Assert.False(File.Exists(files.Output("docx")));
    }

    [Fact]
    public async Task Fetch_reauthorizes_registered_paths_before_reopening_a_replaced_directory() {
        if (OperatingSystem.IsWindows()) return;
        using var files = new Files(true);
        using var outside = new Files(true);
        var service = new OfficeImoAgentService(new AgentPathPolicy([files.Root]));
        var search = await service.SearchAsync(files.Input, "hello");
        var hit = Assert.Single(search.Results);
        Directory.Delete(files.Input, true);
        Directory.CreateSymbolicLink(files.Input, outside.Input);
        await Assert.ThrowsAsync<UnauthorizedAccessException>(() => service.FetchAsync(search.SourceId, hit.Id));
    }
}
