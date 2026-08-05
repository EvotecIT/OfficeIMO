using System.Text.Json;
using OfficeIMO.Tool.Commands.Markup;
using Xunit;

namespace OfficeIMO.Tool.Tests;

public sealed class OfficeImoToolAppTests {
    [Fact]
    public async Task HelpPresentsOneNamespacedCommandSurface() {
        await using var input = new MemoryStream();
        await using var output = new MemoryStream();
        using var error = new StringWriter();

        int exitCode = await OfficeImoToolApp.RunAsync(["help"], input, output, error);
        string help = Encoding.UTF8.GetString(output.ToArray());

        Assert.Equal((int)OfficeImoToolExitCode.Success, exitCode);
        Assert.Contains("officeimo html", help, StringComparison.Ordinal);
        Assert.Contains("officeimo convert", help, StringComparison.Ordinal);
        Assert.Contains("officeimo reader", help, StringComparison.Ordinal);
        Assert.Contains("officeimo markup", help, StringComparison.Ordinal);
        Assert.Equal(string.Empty, error.ToString());
    }

    [Fact]
    public async Task UnknownAreaReturnsTheSharedUsageExitCode() {
        await using var input = new MemoryStream();
        await using var output = new MemoryStream();
        using var error = new StringWriter();

        int exitCode = await OfficeImoToolApp.RunAsync(["unknown"], input, output, error);

        Assert.Equal((int)OfficeImoToolExitCode.Usage, exitCode);
        Assert.Contains("Unknown command area", error.ToString(), StringComparison.Ordinal);
        Assert.Contains("officeimo html", error.ToString(), StringComparison.Ordinal);
    }

    [Fact]
    public async Task MarkupValidationUsesInjectedStreamsAndSourceGeneratedJson() {
        const string markup = """
---
profile: document
---
# Unified tool

Body
""";
        await using var input = new MemoryStream(Encoding.UTF8.GetBytes(markup));
        await using var output = new MemoryStream();
        using var error = new StringWriter();

        int exitCode = await OfficeImoToolApp.RunAsync(
            ["markup", "validate", "-", "--profile", "document"],
            input,
            output,
            error);

        Assert.Equal((int)OfficeImoToolExitCode.Success, exitCode);
        using JsonDocument document = JsonDocument.Parse(output.ToArray());
        Assert.False(document.RootElement.GetProperty("HasErrors").GetBoolean());
        Assert.Equal(string.Empty, error.ToString());
    }

    [Fact]
    public async Task MarkupInputLimitIsEnforcedForStandardInput() {
        await using var input = new MemoryStream(Encoding.UTF8.GetBytes("# Too long"));
        await using var output = new MemoryStream();
        using var error = new StringWriter();

        int exitCode = await OfficeImoToolApp.RunAsync(
            ["markup", "validate", "-", "--max-input-bytes", "4"],
            input,
            output,
            error);

        Assert.Equal((int)OfficeImoToolExitCode.UnsupportedInput, exitCode);
        Assert.Contains("configured byte limit", error.ToString(), StringComparison.Ordinal);
    }

    [Fact]
    public async Task MarkupInputLimitCountsUtf8BytesInsteadOfDecodedCharacters() {
        await using var input = new MemoryStream(Encoding.UTF8.GetBytes("é"));
        await using var output = new MemoryStream();
        using var error = new StringWriter();

        int exitCode = await OfficeImoToolApp.RunAsync(
            ["markup", "validate", "-", "--max-input-bytes", "1"],
            input,
            output,
            error);

        Assert.Equal((int)OfficeImoToolExitCode.UnsupportedInput, exitCode);
        Assert.Contains("configured byte limit", error.ToString(), StringComparison.Ordinal);
    }

    [Fact]
    public async Task MarkupInputLimitIsEnforcedForFileInput() {
        string inputPath = Path.Combine(
            Path.GetTempPath(),
            "OfficeIMO.Tool.Tests",
            Guid.NewGuid().ToString("N"),
            "input.markup");
        Directory.CreateDirectory(Path.GetDirectoryName(inputPath)!);
        await File.WriteAllTextAsync(inputPath, "# Too long");
        await using var input = new MemoryStream();
        await using var output = new MemoryStream();
        using var error = new StringWriter();

        try {
            int exitCode = await OfficeImoToolApp.RunAsync(
                ["markup", "validate", inputPath, "--max-input-bytes", "4"],
                input,
                output,
                error);

            Assert.Equal((int)OfficeImoToolExitCode.UnsupportedInput, exitCode);
            Assert.Contains("configured byte limit", error.ToString(), StringComparison.Ordinal);
        } finally {
            Directory.Delete(Path.GetDirectoryName(inputPath)!, recursive: true);
        }
    }

    [Fact]
    public async Task MarkupEmitClassifiesDestinationWriteFailuresAsOutputFailures() {
        string outputDirectory = Path.Combine(
            Path.GetTempPath(),
            "OfficeIMO.Tool.Tests",
            Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(outputDirectory);
        await using var input = new MemoryStream(Encoding.UTF8.GetBytes("# Heading"));
        await using var output = new MemoryStream();
        using var error = new StringWriter();

        try {
            int exitCode = await OfficeImoToolApp.RunAsync(
                ["markup", "emit", "-", "--output", outputDirectory],
                input,
                output,
                error);

            Assert.Equal((int)OfficeImoToolExitCode.OutputFailed, exitCode);
            Assert.NotEqual(string.Empty, error.ToString());
        } finally {
            Directory.Delete(outputDirectory, recursive: true);
        }
    }

    [Fact]
    public async Task MarkupMissingFileRetainsInputNotFoundClassification() {
        string missing = Path.Combine(
            Path.GetTempPath(),
            "OfficeIMO.Tool.Tests",
            Guid.NewGuid().ToString("N"),
            "missing.markup");
        await using var input = new MemoryStream();
        await using var output = new MemoryStream();
        using var error = new StringWriter();

        int exitCode = await OfficeImoToolApp.RunAsync(
            ["markup", "validate", missing],
            input,
            output,
            error);

        Assert.Equal((int)OfficeImoToolExitCode.InputNotFound, exitCode);
        Assert.NotEqual(string.Empty, error.ToString());
    }

    [Theory]
    [InlineData("html", "read")]
    [InlineData("reader", "convert")]
    [InlineData("markup", "capabilities")]
    public async Task AreaCommandsRejectCommandsOwnedByAnotherArea(string area, string command) {
        await using var input = new MemoryStream();
        await using var output = new MemoryStream();
        using var error = new StringWriter();

        int exitCode = await OfficeImoToolApp.RunAsync([area, command], input, output, error);

        Assert.NotEqual(0, exitCode);
        Assert.Contains("command", error.ToString(), StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void MarkupArgumentsRejectNonPositiveInputLimits() {
        Assert.Throws<ArgumentException>(() =>
            MarkupArguments.Parse(["validate", "-", "--max-input-bytes", "0"]));
    }

    public static TheoryData<string, string, string> MarkupExports => new() {
        {
            "docx",
            ".docx",
            """
---
profile: document
---
# Unified document

Body
"""
        },
        {
            "xlsx",
            ".xlsx",
            """
---
profile: workbook
---
@sheet {
  name: Summary
}

::range address=A1
Name,Value
One,1
"""
        },
        {
            "pptx",
            ".pptx",
            """
---
profile: presentation
---
# Unified presentation

@slide {
  layout: title-and-content
}

- One
- Two
"""
        }
    };

    [Theory]
    [MemberData(nameof(MarkupExports))]
    public async Task MarkupExportRoutesToEachOwningExporter(
        string target,
        string extension,
        string markup) {
        string root = Path.Combine(
            Path.GetTempPath(),
            "OfficeIMO.Tool.Tests",
            Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(root);
        string outputPath = Path.Combine(root, "artifact" + extension);
        await using var input = new MemoryStream(Encoding.UTF8.GetBytes(markup));
        await using var output = new MemoryStream();
        using var error = new StringWriter();

        try {
            int exitCode = await OfficeImoToolApp.RunAsync(
                ["markup", "export", "-", "--target", target, "--output", outputPath],
                input,
                output,
                error);

            Assert.Equal((int)OfficeImoToolExitCode.Success, exitCode);
            Assert.True(File.Exists(outputPath));
            Assert.True(new FileInfo(outputPath).Length > 0);
            using JsonDocument envelope = JsonDocument.Parse(output.ToArray());
            Assert.Equal(target, envelope.RootElement.GetProperty("Target").GetString());
        } finally {
            Directory.Delete(root, recursive: true);
        }
    }
}
