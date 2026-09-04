using System.Text.Json;
using OfficeIMO.Pdf;
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
        Assert.Contains("officeimo read", help, StringComparison.Ordinal);
        Assert.Contains("officeimo extract", help, StringComparison.Ordinal);
        Assert.Contains("officeimo inspect", help, StringComparison.Ordinal);
        Assert.Contains("officeimo reader", help, StringComparison.Ordinal);
        Assert.Contains("officeimo markup", help, StringComparison.Ordinal);
        Assert.Contains("officeimo tabular", help, StringComparison.Ordinal);
        Assert.Contains("officeimo workflow", help, StringComparison.Ordinal);
        Assert.Contains("officeimo pdf redact", help, StringComparison.Ordinal);
        Assert.Contains("officeimo provenance", help, StringComparison.Ordinal);
        Assert.Equal(string.Empty, error.ToString());
    }

    [Fact]
    public async Task PdfRedactionHelpDocumentsReviewSchemasAndPasswordEnvironmentVariables() {
        await using var input = new MemoryStream();
        await using var output = new MemoryStream();
        using var error = new StringWriter();

        int exitCode = await OfficeImoToolApp.RunAsync(["pdf", "--help"], input, output, error);
        string help = Encoding.UTF8.GetString(output.ToArray());

        Assert.Equal((int)OfficeImoToolExitCode.Success, exitCode);
        Assert.Contains("officeimo pdf redact plan", help, StringComparison.Ordinal);
        Assert.Contains("officeimo.pdf.redaction.decisions.v1", help, StringComparison.Ordinal);
        Assert.Contains("--password-env", help, StringComparison.Ordinal);
        Assert.DoesNotContain("--password <", help, StringComparison.Ordinal);
        Assert.Equal(string.Empty, error.ToString());
    }

    [Fact]
    public async Task PdfRedactionCliPlansAndAppliesReviewedJsonWithoutLeakingMatchedText() {
        string directory = Path.Combine(Path.GetTempPath(), "OfficeIMO.Tool.Redaction.Tests", Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        string source = Path.Combine(directory, "source.pdf");
        string recipe = Path.Combine(directory, "recipe.json");
        string plan = Path.Combine(directory, "plan.json");
        string decisions = Path.Combine(directory, "decisions.json");
        string redacted = Path.Combine(directory, "redacted.pdf");
        string evidence = Path.Combine(directory, "evidence.json");
        const string sensitive = "CliSecret-881";
        try {
            PdfDocument.Create(compose => compose.Page(page => page.Content(content => content.Item(item => item.Paragraph(paragraph => paragraph.Text(sensitive)))))).Save(source);
            await File.WriteAllTextAsync(recipe, """
                {
                  "schema": "officeimo.pdf.redaction.recipe.v1",
                  "rules": [ { "kind": "Literal", "value": "CliSecret-881" } ]
                }
                """);
            await using var input = new MemoryStream();
            await using var planOutput = new MemoryStream();
            using var planError = new StringWriter();
            int planExit = await OfficeImoToolApp.RunAsync(["pdf", "redact", "plan", source, "--recipe", recipe, "--evidence", plan], input, planOutput, planError);
            Assert.Equal((int)OfficeImoToolExitCode.Success, planExit);
            Assert.DoesNotContain(sensitive, await File.ReadAllTextAsync(plan), StringComparison.Ordinal);
            using JsonDocument planJson = JsonDocument.Parse(planOutput.ToArray());
            string sourceSha = planJson.RootElement.GetProperty("sourceSha256").GetString()!;
            string recipeSha = planJson.RootElement.GetProperty("recipeSha256").GetString()!;
            string candidateId = planJson.RootElement.GetProperty("candidates")[0].GetProperty("id").GetString()!;
            await File.WriteAllTextAsync(decisions, JsonSerializer.Serialize(new {
                schema = "officeimo.pdf.redaction.decisions.v1",
                sourceSha256 = sourceSha,
                recipeSha256 = recipeSha,
                approvedCandidateIds = new[] { candidateId },
                rejectedCandidateIds = Array.Empty<string>()
            }));
            await using var applyOutput = new MemoryStream();
            using var applyError = new StringWriter();

            int applyExit = await OfficeImoToolApp.RunAsync(["pdf", "redact", "apply", source, "--recipe", recipe, "--decisions", decisions, "--output", redacted, "--evidence", evidence], input, applyOutput, applyError);

            Assert.Equal((int)OfficeImoToolExitCode.Success, applyExit);
            Assert.DoesNotContain(sensitive, PdfDocument.Load(redacted).Read().Text, StringComparison.Ordinal);
            Assert.DoesNotContain(sensitive, await File.ReadAllTextAsync(evidence), StringComparison.Ordinal);
            Assert.Equal(string.Empty, applyError.ToString());
        } finally {
            if (Directory.Exists(directory)) Directory.Delete(directory, recursive: true);
        }
    }

    [Fact]
    public async Task PdfRedactionForceCannotOverwriteRecipeInput() {
        string directory = Path.Combine(Path.GetTempPath(), "OfficeIMO.Tool.Redaction.Tests", Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        string source = Path.Combine(directory, "source.pdf");
        string recipe = Path.Combine(directory, "recipe.json");
        const string recipeJson = """
            {
              "schema": "officeimo.pdf.redaction.recipe.v1",
              "rules": [ { "kind": "Literal", "value": "protected recipe" } ]
            }
            """;
        try {
            PdfDocument.Create(compose => compose.Page(page => page.Content(content => content.Item(item => item.Paragraph(paragraph => paragraph.Text("protected recipe")))))).Save(source);
            await File.WriteAllTextAsync(recipe, recipeJson);
            await using var input = new MemoryStream();
            await using var output = new MemoryStream();
            using var error = new StringWriter();

            int exitCode = await OfficeImoToolApp.RunAsync(["pdf", "redact", "plan", source, "--recipe", recipe, "--evidence", recipe, "--force"], input, output, error);

            Assert.Equal((int)OfficeImoToolExitCode.OperationFailed, exitCode);
            Assert.Equal(recipeJson, await File.ReadAllTextAsync(recipe));
        } finally {
            if (Directory.Exists(directory)) Directory.Delete(directory, recursive: true);
        }
    }

    [Fact]
    public async Task VersionReportsThePackedAssemblyVersion() {
        await using var input = new MemoryStream();
        await using var output = new MemoryStream();
        using var error = new StringWriter();

        int exitCode = await OfficeImoToolApp.RunAsync(["--version"], input, output, error);

        Assert.Equal((int)OfficeImoToolExitCode.Success, exitCode);
        Assert.Equal(
            "OfficeIMO.Tool " + OfficeImoToolApp.GetVersion() + Environment.NewLine,
            Encoding.UTF8.GetString(output.ToArray()));
        Assert.Equal(string.Empty, error.ToString());
    }

    [Theory]
    [InlineData("read")]
    [InlineData("extract")]
    public async Task TopLevelExtractionAliasesUseTheReaderPipeline(string command) {
        string directory = Path.Combine(
            Path.GetTempPath(),
            "OfficeIMO.Tool.Tests",
            Guid.NewGuid().ToString("N"));
        string path = Path.Combine(directory, "source.md");
        Directory.CreateDirectory(directory);
        await File.WriteAllTextAsync(path, "# Alias proof\n\nReader content");
        await using var input = new MemoryStream();
        await using var output = new MemoryStream();
        using var error = new StringWriter();

        try {
            int exitCode = await OfficeImoToolApp.RunAsync(
                [command, path, "--format", "markdown"],
                input,
                output,
                error);

            Assert.Equal((int)OfficeImoToolExitCode.Success, exitCode);
            Assert.Contains("Alias proof", Encoding.UTF8.GetString(output.ToArray()), StringComparison.Ordinal);
            Assert.Equal(string.Empty, error.ToString());
        } finally {
            Directory.Delete(directory, recursive: true);
        }
    }

    [Fact]
    public async Task TopLevelInspectAliasUsesTheAgentPipeline() {
        string directory = Path.Combine(
            Path.GetTempPath(),
            "OfficeIMO.Tool.Tests",
            Guid.NewGuid().ToString("N"));
        string path = Path.Combine(directory, "source.md");
        Directory.CreateDirectory(directory);
        await File.WriteAllTextAsync(path, "# Inspect alias\n\nAgent content");
        await using var input = new MemoryStream();
        await using var output = new MemoryStream();
        using var error = new StringWriter();

        try {
            int exitCode = await OfficeImoToolApp.RunAsync(["inspect", path], input, output, error);

            Assert.Equal((int)OfficeImoToolExitCode.Success, exitCode);
            using JsonDocument document = JsonDocument.Parse(output.ToArray());
            Assert.Equal("Markdown", document.RootElement.GetProperty("kind").GetString());
            Assert.Equal(string.Empty, error.ToString());
        } finally {
            Directory.Delete(directory, recursive: true);
        }
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
