using System.IO.Compression;
using System.Threading;
using System.Threading.Tasks;
using OfficeIMO.Reader;
using OfficeIMO.Reader.IWork;
using OfficeIMO.Tool.Commands.Reader;
using Xunit;

namespace OfficeIMO.Tool.Tests;

public sealed class ReaderDirectoryBundleCommandTests {
    [Fact]
    public async Task Read_bundle_preserves_source_identity_and_rejects_output_inside_the_source() {
        using var temporary = new ReaderToolTemporaryDirectory();
        string bundle = ExtractBundle(temporary.Path);
        using var output = new StringWriter();
        using var error = new StringWriter();
        int exitCode = await ReaderCommand.RunAsync(new[] { "read", bundle + Path.DirectorySeparatorChar, "--format", "json" },
            Stream.Null, output, error);
        Assert.Equal((int)OfficeImoToolExitCode.Success, exitCode);
        OfficeDocumentReadResult document = OfficeDocumentReadResultJson.Deserialize(output.ToString());
        Assert.Equal(bundle, document.Source.Path);
        Assert.NotEmpty(document.Source.SourceHash!);
        Assert.NotEmpty(document.Chunks);

        string forbidden = Path.Combine(bundle, "converted.json");
        exitCode = await ReaderCommand.RunAsync(new[] { "read", bundle, "--format", "json", "--output", forbidden },
            Stream.Null, output, error);
        Assert.Equal((int)OfficeImoToolExitCode.OutputFailed, exitCode);
        Assert.False(File.Exists(forbidden));
    }

    [Fact]
    public async Task Folder_command_accepts_one_bundle_with_a_named_output_and_a_byte_budget() {
        using var temporary = new ReaderToolTemporaryDirectory();
        string bundle = ExtractBundle(temporary.Path);
        string outputDirectory = Path.Combine(temporary.Path, "output");
        using var output = new StringWriter();
        using var error = new StringWriter();
        int exitCode = await ReaderCommand.RunAsync(new[] {
            "folder", bundle, "--output", outputDirectory, "--format", "json"
        }, Stream.Null, output, error);
        Assert.Equal((int)OfficeImoToolExitCode.Success, exitCode);
        string saved = Assert.Single(Directory.EnumerateFiles(outputDirectory));
        Assert.Equal("document.pages.reader.json", Path.GetFileName(saved));
        Assert.Equal(bundle, OfficeDocumentReadResultJson.Deserialize(File.ReadAllText(saved)).Source.Path);

        string boundedOutput = Path.Combine(temporary.Path, "bounded");
        exitCode = await ReaderCommand.RunAsync(new[] {
            "folder", bundle, "--output", boundedOutput, "--format", "json", "--max-total-bytes", "1"
        }, Stream.Null, output, error);
        Assert.Equal((int)OfficeImoToolExitCode.Success, exitCode);
        Assert.Empty(Directory.EnumerateFiles(boundedOutput));
    }

    [Fact]
    public void Folder_discovery_uses_the_registered_bundle_owner_without_descending_into_resources() {
        using var temporary = new ReaderToolTemporaryDirectory();
        string bundle = ExtractBundle(temporary.Path);
        File.Copy(Fixture(), Path.Combine(bundle, "nested.pages"));
        OfficeDocumentReader reader = new OfficeDocumentReaderBuilder().AddIWorkHandler().Build();
        IReadOnlyList<string> paths = ReaderToolFileDiscovery.FindSupportedFiles(temporary.Path,
            reader, true, 10, null, CancellationToken.None);
        Assert.Equal(new[] { bundle }, paths);
    }

    private static string ExtractBundle(string root) {
        string bundle = Path.Combine(root, "document.pages");
        Directory.CreateDirectory(bundle);
        ZipFile.ExtractToDirectory(Fixture(), bundle);
        return bundle;
    }

    private static string Fixture() => Path.Combine(AppContext.BaseDirectory,
        "Documents", "IWorkCorpus", "nim-iwork", "simple.pages");
}
