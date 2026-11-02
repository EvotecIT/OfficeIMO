using OfficeIMO.Reader;
using OfficeIMO.Reader.IWork;
using System.IO.Compression;
using OfficeIMO.Tool.Agent;
using OfficeIMO.Tool.Commands.Reader;
using Xunit;

namespace OfficeIMO.Tool.Tests;

public sealed partial class AgentIWorkConversionTests {
    [Theory]
    [InlineData(false, false)]
    [InlineData(false, true)]
    [InlineData(true, false)]
    [InlineData(true, true)]
    public async Task Conversion_source_ids_retain_the_admitted_package_limit(bool directoryPackage, bool failedConversion) {
        using var files = new Files(true);
        string input = directoryPackage ? files.Input : Path.Combine(files.Root, "transport.pages");
        if (!directoryPackage) ZipFile.CreateFromDirectory(files.Input, input, CompressionLevel.NoCompression, false);
        var registry = new AgentSourceRegistry();
        var service = new OfficeImoAgentService(new AgentPathPolicy([files.Root]), registry);
        var hit = Assert.Single((await service.SearchAsync(input, "hello")).Results);
        string member = Path.Combine(files.Input, "extra.bin");
        using (var padding = File.Create(member)) padding.SetLength(ReaderToolArguments.DefaultMaxInputBytes + 1);
        RefreshTransport();
        const long maximum = 96L * 1024 * 1024;
        AgentConvertResult result = await service.ConvertAsync(input, files.Output("docx"), "docx", maximumInputBytes: maximum);
        Assert.True(result.Workflow!.Succeeded, result.Workflow.Summary);
        Assert.Equal(result.SourceId, registry.Resolve(result.SourceId).SourceId);
        Assert.Equal(result.SourceId, registry.Resolve(result.SourceId, input).SourceId);
        Assert.Contains("hello pages", AgentJson.Serialize(await service.FetchAsync(result.SourceId, hit.Id)));
        Assert.Equal(result.SourceId, registry.Resolve(result.SourceId).SourceId);
        Assert.Contains("hello pages", AgentJson.Serialize(await service.FetchAsync(result.SourceId, hit.Id, sourcePath: input)));
        if (failedConversion) {
            var rejected = await service.ConvertAsync(input, files.Output("docx"), "docx", overwrite: true, maximumInputBytes: 32);
            Assert.False(rejected.Workflow!.Succeeded);
            Assert.Contains("hello pages", AgentJson.Serialize(await service.FetchAsync(result.SourceId, hit.Id)));
        } else {
            await RejectDefaultRead(() => service.InspectAsync(input));
            Assert.Contains("hello pages", AgentJson.Serialize(await service.FetchAsync(result.SourceId, hit.Id)));
            await RejectDefaultRead(() => service.SearchAsync(input, "hello"));
            Assert.Contains("hello pages", AgentJson.Serialize(await service.FetchAsync(result.SourceId, hit.Id)));
            await RejectDefaultRead(() => service.ConvertAsync(input, files.Output("md"), "markdown"));
            Assert.Contains("hello pages", AgentJson.Serialize(await service.FetchAsync(result.SourceId, hit.Id)));
        }
        using (var padding = File.OpenWrite(member)) padding.SetLength(maximum + 1);
        RefreshTransport();
        if (directoryPackage)
            await Assert.ThrowsAsync<InvalidDataException>(() => service.FetchAsync(result.SourceId, hit.Id));
        else
            await Assert.ThrowsAsync<AgentUsageException>(() => service.FetchAsync(result.SourceId, hit.Id));
        File.WriteAllText(member, "changed");
        RefreshTransport();
        Assert.Throws<AgentUsageException>(() => registry.Resolve(result.SourceId));

        void RefreshTransport() {
            if (directoryPackage) return;
            File.Delete(input);
            ZipFile.CreateFromDirectory(files.Input, input, CompressionLevel.NoCompression, false);
        }

        async Task RejectDefaultRead(Func<Task> read) {
            if (directoryPackage) await Assert.ThrowsAsync<InvalidDataException>(read);
            else await Assert.ThrowsAsync<IOException>(read);
        }
    }

    [Theory]
    [InlineData("pages", "markdown", true)]
    [InlineData("numbers", "md", false)]
    [InlineData("key", "json", true)]
    public async Task Semantic_conversion_cannot_replace_or_add_source_package_members(string kind, string format, bool existingMember) {
        using var files = new Files(true, kind);
        string output = existingMember ? Path.Combine(files.Input, "Index", "Document.iwa")
            : Path.Combine(files.Input, "new-output.md");
        byte[]? original = existingMember ? File.ReadAllBytes(output) : null;
        var service = new OfficeImoAgentService(new AgentPathPolicy([files.Root]));
        await Assert.ThrowsAsync<ReaderToolOutputException>(() =>
            service.ConvertAsync(files.Input, output, format, overwrite: true));
        if (existingMember) Assert.Equal(original, File.ReadAllBytes(output));
        else Assert.False(File.Exists(output));
        Assert.Empty(Directory.GetFiles(files.Input, "*.tmp", SearchOption.AllDirectories));
    }

    [Fact]
    public async Task Semantic_conversion_rejects_an_output_parent_alias_into_the_source_package() {
        if (OperatingSystem.IsWindows()) return;
        using var files = new Files(true);
        string alias = Path.Combine(files.Root, "output-alias");
        Directory.CreateSymbolicLink(alias, Path.Combine(files.Input, "Index"));
        string member = Path.Combine(files.Input, "Index", "Document.iwa");
        byte[] original = File.ReadAllBytes(member);
        var service = new OfficeImoAgentService(new AgentPathPolicy([files.Root]));
        await Assert.ThrowsAsync<ReaderToolOutputException>(() =>
            service.ConvertAsync(files.Input, Path.Combine(alias, "Document.iwa"), "json", overwrite: true));
        Assert.Equal(original, File.ReadAllBytes(member));
    }

    [Fact]
    public async Task Completed_directory_conversion_retains_its_source_id_when_output_publication_conflicts() {
        using var files = new Files(true);
        string output = files.Output("docx"); File.WriteAllText(output, "keep existing");
        var service = new OfficeImoAgentService(new AgentPathPolicy([files.Root]));
        var result = await service.ConvertAsync(files.Input, output, "docx");
        Assert.False(result.Workflow!.Succeeded);
        Assert.NotNull(result.Workflow.ConversionEvidence);
        Assert.Contains(result.Workflow.Diagnostics, diagnostic => diagnostic.Code == "conversion-output-reopened");
        Assert.Equal("keep existing", File.ReadAllText(output));
        Assert.Equal((await service.InspectAsync(files.Input)).SourceId, result.SourceId);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task Native_identity_binds_returned_reader_capture_even_when_members_change_before_registration(bool directoryPackage) {
        using var files = new Files(true);
        string member = Path.Combine(files.Input, "extra.bin"); File.WriteAllText(member, "baseline");
        string input = directoryPackage ? files.Input : Path.Combine(files.Root, "transport.pages");
        if (!directoryPackage) ZipFile.CreateFromDirectory(files.Input, input, CompressionLevel.NoCompression, false);
        var reader = new OfficeDocumentReaderBuilder().AddIWorkHandler().Build();
        var options = new ReaderOptions { ComputeHashes = true };
        var captured = await reader.ReadDocumentAsync(input, options);
        string hash = Assert.IsType<string>(captured.Source.SourceHash);
        DateTime timestamp = directoryPackage ? Directory.GetLastWriteTimeUtc(input) : File.GetLastWriteTimeUtc(input);
        var registry = new AgentSourceRegistry();
        string expectedId = registry.Register(input).SourceId;
        File.WriteAllText(member, "mutation");
        if (!directoryPackage) {
            File.Delete(input);
            ZipFile.CreateFromDirectory(files.Input, input, CompressionLevel.NoCompression, false);
        }
        var registration = registry.RegisterIWorkSnapshot(input, directoryPackage, hash, timestamp, CancellationToken.None,
            capturedLengthBytes: captured.Source.LengthBytes);
        Assert.Equal(expectedId, registration.SourceId);
        Assert.Contains("hello pages", captured.Markdown);
        var changed = await reader.ReadDocumentAsync(input, options);
        Assert.NotEqual(hash, changed.Source.SourceHash);
        Assert.Throws<AgentUsageException>(() => registry.ResolveIWorkSnapshot(expectedId, input, directoryPackage,
            changed.Source.SourceHash!, timestamp, CancellationToken.None));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task Semantic_json_output_and_source_id_share_the_captured_reader_hash(bool directoryPackage) {
        using var files = new Files(directoryPackage);
        var registry = new AgentSourceRegistry();
        var service = new OfficeImoAgentService(new AgentPathPolicy([files.Root]), registry);
        var result = await service.ConvertAsync(files.Input, files.Output("json"), "json");
        using var document = System.Text.Json.JsonDocument.Parse(File.ReadAllText(result.OutputPath));
        string hash = document.RootElement.GetProperty("source").GetProperty("sourceHash").GetString()!;
        var captured = registry.RegisterIWorkSnapshot(files.Input, directoryPackage, hash,
            directoryPackage ? Directory.GetLastWriteTimeUtc(files.Input) : File.GetLastWriteTimeUtc(files.Input),
            CancellationToken.None, capturedLengthBytes: directoryPackage ? null : new FileInfo(files.Input).Length);
        Assert.Equal(result.SourceId, captured.SourceId);
        Assert.Equal((await service.InspectAsync(files.Input)).SourceId, result.SourceId);
    }
}
