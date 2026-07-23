using OfficeIMO.Reader.Markdown;
using OfficeIMO.Reader.Tool;
using System.Threading;
using System.Threading.Tasks;
using Xunit;

namespace OfficeIMO.Reader.Tests;

public sealed class ReaderToolTests {
    [Fact]
    public async Task ReadsStandardInputAsMarkdown() {
        await using var input = new MemoryStream(Encoding.UTF8.GetBytes("# Tool heading\n\nBody"));
        using var output = new StringWriter();
        using var error = new StringWriter();

        int exitCode = await ReaderToolApp.RunAsync(
            new[] { "read", "-", "--name", "input.md" },
            input,
            output,
            error);

        Assert.Equal((int)ReaderToolExitCode.Success, exitCode);
        Assert.Contains("# Tool heading", output.ToString(), StringComparison.Ordinal);
        Assert.Equal(string.Empty, error.ToString());
    }

    [Fact]
    public async Task StandardInputUsesABoundedDefaultAndSupportsAnExplicitLimit() {
        ReaderToolArguments defaults = ReaderToolArguments.Parse(new[] { "read", "-" });
        Assert.Equal(ReaderToolArguments.DefaultMaxInputBytes, defaults.MaxInputBytes);

        await using var input = new ReaderToolNonSeekableStream(Encoding.UTF8.GetBytes("0123456789abcdefg"));
        using var output = new StringWriter();
        using var error = new StringWriter();

        int exitCode = await ReaderToolApp.RunAsync(
            new[] { "read", "-", "--name", "input.txt", "--max-input-bytes", "16" },
            input,
            output,
            error);

        Assert.Equal((int)ReaderToolExitCode.ReadFailed, exitCode);
        Assert.Contains("MaxInputBytes", error.ToString(), StringComparison.Ordinal);
    }

    [Fact]
    public async Task EmitsStableV5Json() {
        await using var input = new MemoryStream(Encoding.UTF8.GetBytes("plain text"));
        using var output = new StringWriter();
        using var error = new StringWriter();

        int exitCode = await ReaderToolApp.RunAsync(
            new[] { "read", "-", "--name", "input.txt", "--format", "json" },
            input,
            output,
            error);
        OfficeDocumentReadResult document = OfficeDocumentReadResultJson.Deserialize(output.ToString());

        Assert.Equal((int)ReaderToolExitCode.Success, exitCode);
        Assert.Equal(OfficeDocumentReadResultSchema.CurrentVersion, document.SchemaVersion);
        Assert.Equal(OfficeDocumentReadResultSchema.Id, document.SchemaId);
    }

    [Fact]
    public async Task ConvertsFolderWithDeterministicRelativeOutputs() {
        using var temporary = new ReaderToolTemporaryDirectory();
        string inputRoot = Path.Combine(temporary.Path, "input");
        string nestedRoot = Path.Combine(inputRoot, "nested");
        string outputRoot = Path.Combine(temporary.Path, "output");
        Directory.CreateDirectory(nestedRoot);
        await File.WriteAllTextAsync(Path.Combine(inputRoot, "alpha.md"), "# Alpha");
        await File.WriteAllTextAsync(Path.Combine(nestedRoot, "data.csv"), "name,value\none,1");
        using var output = new StringWriter();
        using var error = new StringWriter();

        int exitCode = await ReaderToolApp.RunAsync(
            new[] {
                "folder", inputRoot,
                "--output", outputRoot,
                "--format", "json",
                "--concurrency", "2"
            },
            Stream.Null,
            output,
            error);

        Assert.Equal((int)ReaderToolExitCode.Success, exitCode);
        Assert.True(File.Exists(Path.Combine(outputRoot, "alpha.md.reader.json")));
        Assert.True(File.Exists(Path.Combine(outputRoot, "nested", "data.csv.reader.json")));
        Assert.Equal("Converted 2 document(s)." + Environment.NewLine, error.ToString());
    }

    [Fact]
    public async Task FolderUsesABoundedPerFileDefaultAndSupportsAnExplicitLimit() {
        using var temporary = new ReaderToolTemporaryDirectory();
        string inputRoot = Path.Combine(temporary.Path, "input");
        string outputRoot = Path.Combine(temporary.Path, "output");
        Directory.CreateDirectory(inputRoot);
        await File.WriteAllTextAsync(Path.Combine(inputRoot, "large.txt"), "0123456789abcdefg");
        ReaderToolArguments defaults = ReaderToolArguments.Parse(new[] {
            "folder", inputRoot, "--output", outputRoot
        });
        using var output = new StringWriter();
        using var error = new StringWriter();

        int exitCode = await ReaderToolApp.RunAsync(
            new[] {
                "folder", inputRoot,
                "--output", outputRoot,
                "--max-input-bytes", "16"
            },
            Stream.Null,
            output,
            error);

        Assert.Equal(ReaderToolArguments.DefaultMaxInputBytes, defaults.MaxInputBytes);
        Assert.Equal((int)ReaderToolExitCode.ReadFailed, exitCode);
        Assert.Contains("MaxInputBytes", error.ToString(), StringComparison.Ordinal);
        Assert.False(Directory.Exists(outputRoot));
    }

    [Fact]
    public async Task ReadRejectsAnOutputThatAliasesTheInput() {
        using var temporary = new ReaderToolTemporaryDirectory();
        string inputPath = Path.Combine(temporary.Path, "input.md");
        const string original = "# Original";
        await File.WriteAllTextAsync(inputPath, original);
        using var output = new StringWriter();
        using var error = new StringWriter();

        int exitCode = await ReaderToolApp.RunAsync(
            new[] { "read", inputPath, "--output", inputPath },
            Stream.Null,
            output,
            error);

        Assert.Equal((int)ReaderToolExitCode.OutputFailed, exitCode);
        Assert.Contains("different from the input", error.ToString(), StringComparison.Ordinal);
        Assert.Equal(original, await File.ReadAllTextAsync(inputPath));
    }

    [Fact]
    public async Task ReadRejectsASymbolicOutputThatTargetsTheInput() {
        if (OperatingSystem.IsWindows()) return;
        using var temporary = new ReaderToolTemporaryDirectory();
        string inputPath = Path.Combine(temporary.Path, "input.md");
        string outputPath = Path.Combine(temporary.Path, "output.md");
        const string original = "# Original";
        await File.WriteAllTextAsync(inputPath, original);
        File.CreateSymbolicLink(outputPath, inputPath);
        using var output = new StringWriter();
        using var error = new StringWriter();

        int exitCode = await ReaderToolApp.RunAsync(
            new[] { "read", inputPath, "--output", outputPath },
            Stream.Null,
            output,
            error);

        Assert.Equal((int)ReaderToolExitCode.OutputFailed, exitCode);
        Assert.Contains("different from the input", error.ToString(), StringComparison.Ordinal);
        Assert.Equal(original, await File.ReadAllTextAsync(inputPath));
    }

    [Fact]
    public async Task CapabilityListExcludesDependencyBackedProviders() {
        using var output = new StringWriter();
        using var error = new StringWriter();

        int exitCode = await ReaderToolApp.RunAsync(
            new[] { "capabilities" },
            Stream.Null,
            output,
            error);

        Assert.Equal((int)ReaderToolExitCode.Success, exitCode);
        Assert.Contains("officeimo.reader.epub", output.ToString(), StringComparison.Ordinal);
        Assert.DoesNotContain("ocr", output.ToString(), StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("provider", output.ToString(), StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public async Task InvalidArgumentsReturnDocumentedUsageCode() {
        using var output = new StringWriter();
        using var error = new StringWriter();

        int exitCode = await ReaderToolApp.RunAsync(
            new[] { "folder", "missing-output" },
            Stream.Null,
            output,
            error);

        Assert.Equal((int)ReaderToolExitCode.Usage, exitCode);
        Assert.Contains("requires --output", error.ToString(), StringComparison.Ordinal);
    }

    [Fact]
    public async Task MissingInputReturnsDocumentedNotFoundCode() {
        using var output = new StringWriter();
        using var error = new StringWriter();

        int exitCode = await ReaderToolApp.RunAsync(
            new[] { "read", Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N"), "missing.md") },
            Stream.Null,
            output,
            error);

        Assert.Equal((int)ReaderToolExitCode.InputNotFound, exitCode);
    }

    [Fact]
    public async Task CorruptInputReturnsDocumentedReadFailureCode() {
        using var temporary = new ReaderToolTemporaryDirectory();
        string path = Path.Combine(temporary.Path, "legacy.doc");
        await File.WriteAllBytesAsync(path, new byte[] { 1, 2, 3 });
        using var output = new StringWriter();
        using var error = new StringWriter();

        int exitCode = await ReaderToolApp.RunAsync(
            new[] { "read", path },
            Stream.Null,
            output,
            error);

        Assert.Equal((int)ReaderToolExitCode.ReadFailed, exitCode);
    }

    [Fact]
    public async Task UnsupportedInputReturnsDocumentedFormatCode() {
        await using var input = new ReaderToolUnsupportedStream();
        using var output = new StringWriter();
        using var error = new StringWriter();

        int exitCode = await ReaderToolApp.RunAsync(
            new[] { "read", "-", "--name", "input.txt" },
            input,
            output,
            error);

        Assert.Equal((int)ReaderToolExitCode.UnsupportedInput, exitCode);
    }

    [Fact]
    public async Task FolderRejectsOutputInsideInputTree() {
        using var temporary = new ReaderToolTemporaryDirectory();
        string outputPath = Path.Combine(temporary.Path, "converted");
        using var output = new StringWriter();
        using var error = new StringWriter();

        int exitCode = await ReaderToolApp.RunAsync(
            new[] { "folder", temporary.Path, "--output", outputPath },
            Stream.Null,
            output,
            error);

        Assert.Equal((int)ReaderToolExitCode.OutputFailed, exitCode);
        Assert.Contains("outside the input folder", error.ToString(), StringComparison.Ordinal);
    }

    [Fact]
    public async Task FolderRejectsLinkedOutputThatTargetsInputTree() {
        if (OperatingSystem.IsWindows()) return;
        using var temporary = new ReaderToolTemporaryDirectory();
        string inputRoot = Path.Combine(temporary.Path, "input");
        string convertedRoot = Path.Combine(inputRoot, "converted");
        string linkedOutput = Path.Combine(temporary.Path, "linked-output");
        Directory.CreateDirectory(convertedRoot);
        Directory.CreateSymbolicLink(linkedOutput, convertedRoot);
        await File.WriteAllTextAsync(Path.Combine(inputRoot, "document.md"), "# Input");
        using var output = new StringWriter();
        using var error = new StringWriter();

        int exitCode = await ReaderToolApp.RunAsync(
            new[] { "folder", inputRoot, "--output", linkedOutput },
            Stream.Null,
            output,
            error);

        Assert.Equal((int)ReaderToolExitCode.OutputFailed, exitCode);
        Assert.Contains("outside the input folder", error.ToString(), StringComparison.Ordinal);
        Assert.Empty(Directory.EnumerateFiles(convertedRoot));
    }

    [Fact]
    public async Task FolderRejectsRealOutputInsideLinkedInputTree() {
        if (OperatingSystem.IsWindows()) return;
        using var temporary = new ReaderToolTemporaryDirectory();
        string inputRoot = Path.Combine(temporary.Path, "input");
        string linkedInput = Path.Combine(temporary.Path, "linked-input");
        string outputRoot = Path.Combine(inputRoot, "converted");
        Directory.CreateDirectory(inputRoot);
        Directory.CreateSymbolicLink(linkedInput, inputRoot);
        await File.WriteAllTextAsync(Path.Combine(inputRoot, "document.md"), "# Input");
        using var output = new StringWriter();
        using var error = new StringWriter();

        int exitCode = await ReaderToolApp.RunAsync(
            new[] { "folder", linkedInput, "--output", outputRoot },
            Stream.Null,
            output,
            error);

        Assert.Equal((int)ReaderToolExitCode.OutputFailed, exitCode);
        Assert.Contains("outside the input folder", error.ToString(), StringComparison.Ordinal);
        Assert.False(Directory.Exists(outputRoot));
    }

    [Fact]
    public void FolderDiscoveryStopsAtTheConfiguredSupportedFileBound() {
        using var temporary = new ReaderToolTemporaryDirectory();
        for (int index = 99; index >= 0; index--) {
            File.WriteAllText(Path.Combine(temporary.Path, index.ToString("D3") + ".md"), "# Document");
        }
        OfficeDocumentReader reader = new OfficeDocumentReaderBuilder().AddMarkdownHandler().Build();

        IReadOnlyList<string> paths = ReaderToolFileDiscovery.FindSupportedFiles(
            temporary.Path,
            reader,
            recurse: true,
            maxFiles: 3,
            maxTotalBytes: null,
            CancellationToken.None);

        Assert.Equal(3, paths.Count);
        Assert.All(paths, path => Assert.Equal(".md", Path.GetExtension(path)));
    }

    [Fact]
    public void FolderDiscoveryStopsWhenTheNextSupportedFileExceedsTheByteBudget() {
        using var temporary = new ReaderToolTemporaryDirectory();
        File.WriteAllText(Path.Combine(temporary.Path, "oversized.md"),
            new string('x', 128));
        File.WriteAllText(Path.Combine(temporary.Path, "unvisited.md"),
            new string('y', 128));
        OfficeDocumentReader reader = new OfficeDocumentReaderBuilder()
            .AddMarkdownHandler().Build();

        IReadOnlyList<string> paths = ReaderToolFileDiscovery.FindSupportedFiles(
            temporary.Path,
            reader,
            recurse: true,
            maxFiles: 100,
            maxTotalBytes: 64,
            CancellationToken.None);

        Assert.Empty(paths);
    }

    [Fact]
    public void MacPathSafetyTreatsCasingAliasesAsTheSameInputTree() {
        if (!OperatingSystem.IsMacOS()) return;
        using var temporary = new ReaderToolTemporaryDirectory();
        string input = Path.Combine(temporary.Path, "Input");
        Directory.CreateDirectory(input);
        string casingAlias = Path.Combine(temporary.Path, "input", "converted");

        ReaderToolOutputException exception = Assert.Throws<ReaderToolOutputException>(() =>
            ReaderToolPathSafety.EnsureOutsideInput(input, casingAlias));

        Assert.Contains("outside the input folder", exception.Message, StringComparison.Ordinal);
    }
}

internal sealed class ReaderToolTemporaryDirectory : IDisposable {
    internal ReaderToolTemporaryDirectory() {
        Path = System.IO.Path.Combine(System.IO.Path.GetTempPath(), "OfficeIMO.Reader.Tool.Tests", Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(Path);
    }

    internal string Path { get; }

    public void Dispose() {
        try {
            Directory.Delete(Path, recursive: true);
        } catch (DirectoryNotFoundException) {
        } catch (IOException) {
        } catch (UnauthorizedAccessException) {
        }
    }
}

internal sealed class ReaderToolUnsupportedStream : Stream {
    public override bool CanRead => true;
    public override bool CanSeek => false;
    public override bool CanWrite => false;
    public override long Length => throw new NotSupportedException();
    public override long Position {
        get => throw new NotSupportedException();
        set => throw new NotSupportedException();
    }

    public override void Flush() { }
    public override int Read(byte[] buffer, int offset, int count) => throw new NotSupportedException("Unsupported input stream.");
    public override long Seek(long offset, SeekOrigin origin) => throw new NotSupportedException();
    public override void SetLength(long value) => throw new NotSupportedException();
    public override void Write(byte[] buffer, int offset, int count) => throw new NotSupportedException();
}

internal sealed class ReaderToolNonSeekableStream : Stream {
    private readonly MemoryStream _stream;

    internal ReaderToolNonSeekableStream(byte[] bytes) {
        _stream = new MemoryStream(bytes, writable: false);
    }

    public override bool CanRead => true;
    public override bool CanSeek => false;
    public override bool CanWrite => false;
    public override long Length => throw new NotSupportedException();
    public override long Position {
        get => throw new NotSupportedException();
        set => throw new NotSupportedException();
    }

    public override void Flush() { }
    public override int Read(byte[] buffer, int offset, int count) => _stream.Read(buffer, offset, count);
    public override long Seek(long offset, SeekOrigin origin) => throw new NotSupportedException();
    public override void SetLength(long value) => throw new NotSupportedException();
    public override void Write(byte[] buffer, int offset, int count) => throw new NotSupportedException();

    protected override void Dispose(bool disposing) {
        if (disposing) _stream.Dispose();
        base.Dispose(disposing);
    }
}
