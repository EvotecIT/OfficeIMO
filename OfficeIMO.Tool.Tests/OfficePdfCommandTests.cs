using OfficeIMO.Excel;
using OfficeIMO.PowerPoint;
using OfficeIMO.Tool.Commands.Convert;
using OfficeIMO.Word;
using System.IO.Compression;
using Xunit;

namespace OfficeIMO.Tool.Tests;

public sealed class OfficePdfCommandTests {
    [Theory]
    [InlineData(".docx")]
    [InlineData(".xlsx")]
    [InlineData(".pptx")]
    public async Task ConvertUsesTheOwningOfficePdfAdapterAndWritesAReadableArtifact(string extension) {
        string directory = Path.Combine(Path.GetTempPath(), "OfficeIMO.Tool.Tests", Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        string inputPath = Path.Combine(directory, "source" + extension);
        string outputPath = Path.Combine(directory, "result.pdf");
        CreateSource(inputPath, extension);
        await using var output = new MemoryStream();
        using var error = new StringWriter();

        try {
            int exitCode = await OfficeImoToolApp.RunAsync(
                ["convert", inputPath, outputPath],
                Stream.Null,
                output,
                error);

            Assert.Equal((int)OfficeImoToolExitCode.Success, exitCode);
            Assert.True(File.Exists(outputPath));
            Assert.NotEmpty(OfficeIMO.Pdf.PdfReadDocument.Open(File.ReadAllBytes(outputPath)).Pages);
            Assert.Contains(outputPath, Encoding.UTF8.GetString(output.ToArray()), StringComparison.Ordinal);
            Assert.DoesNotContain("Error ", error.ToString(), StringComparison.Ordinal);
        } finally {
            Directory.Delete(directory, recursive: true);
        }
    }

    [Theory]
    [InlineData(".md")]
    [InlineData(".markdown")]
    [InlineData(".json")]
    public async Task ConvertRoutesTextDestinationsThroughTheReaderPipeline(string outputExtension) {
        string directory = CreateTestDirectory();
        string inputPath = Path.Combine(directory, "source.xlsx");
        string outputPath = Path.Combine(directory, "result" + outputExtension);
        CreateSource(inputPath, ".xlsx");
        await using var output = new MemoryStream();
        using var error = new StringWriter();

        try {
            int exitCode = await OfficeImoToolApp.RunAsync(
                ["convert", inputPath, outputPath],
                Stream.Null,
                output,
                error);

            Assert.Equal((int)OfficeImoToolExitCode.Success, exitCode);
            Assert.True(File.Exists(outputPath));
            Assert.Contains("Tool Excel conversion", await File.ReadAllTextAsync(outputPath), StringComparison.Ordinal);
            Assert.Equal(string.Empty, error.ToString());
        } finally {
            Directory.Delete(directory, recursive: true);
        }
    }

    [Fact]
    public async Task ConvertRejectsUnsupportedDestinationExtensions() {
        await using var output = new MemoryStream();
        using var error = new StringWriter();

        int exitCode = await OfficeImoToolApp.RunAsync(
            ["convert", "source.xlsx", "result.csv"],
            Stream.Null,
            output,
            error);

        Assert.Equal((int)OfficeImoToolExitCode.Usage, exitCode);
        Assert.Contains(".pdf, .md, .markdown, or .json", error.ToString(), StringComparison.Ordinal);
    }

    [Fact]
    public async Task ConvertRefusesToReplaceAnExistingOutputWithoutForce() {
        string directory = Path.Combine(Path.GetTempPath(), "OfficeIMO.Tool.Tests", Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        string inputPath = Path.Combine(directory, "source.docx");
        string outputPath = Path.Combine(directory, "result.pdf");
        CreateSource(inputPath, ".docx");
        await File.WriteAllTextAsync(outputPath, "keep");
        await using var output = new MemoryStream();
        using var error = new StringWriter();

        try {
            int exitCode = await OfficeImoToolApp.RunAsync(
                ["convert", inputPath, "--output", outputPath],
                Stream.Null,
                output,
                error);

            Assert.Equal((int)OfficeImoToolExitCode.OutputFailed, exitCode);
            Assert.Equal("keep", await File.ReadAllTextAsync(outputPath));
            Assert.Contains("--force", error.ToString(), StringComparison.Ordinal);
        } finally {
            Directory.Delete(directory, recursive: true);
        }
    }

    [Fact]
    public async Task ConvertUsesTheDefaultPdfPathAndForceReplacesAnExistingOutput() {
        string directory = Path.Combine(Path.GetTempPath(), "OfficeIMO.Tool.Tests", Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        string inputPath = Path.Combine(directory, "source.docx");
        string outputPath = Path.ChangeExtension(inputPath, ".pdf");
        CreateSource(inputPath, ".docx");
        await File.WriteAllTextAsync(outputPath, "replace me");
        await using var output = new MemoryStream();
        using var error = new StringWriter();

        try {
            int exitCode = await OfficeImoToolApp.RunAsync(
                ["convert", inputPath, "--force"],
                Stream.Null,
                output,
                error);

            Assert.Equal((int)OfficeImoToolExitCode.Success, exitCode);
            Assert.NotEmpty(OfficeIMO.Pdf.PdfReadDocument.Open(File.ReadAllBytes(outputPath)).Pages);
            Assert.Contains(outputPath, Encoding.UTF8.GetString(output.ToArray()), StringComparison.Ordinal);
        } finally {
            Directory.Delete(directory, recursive: true);
        }
    }

    [Fact]
    public async Task ConvertRejectsPackageBombWithDefaultLimits() {
        string directory = CreateTestDirectory();
        string inputPath = Path.Combine(directory, "source.docx");
        string outputPath = Path.Combine(directory, "result.pdf");
        CreateCompressedDocxBomb(inputPath);
        await using var output = new MemoryStream();
        using var error = new StringWriter();

        try {
            int exitCode = await OfficeImoToolApp.RunAsync(
                ["convert", inputPath, "--output", outputPath],
                Stream.Null,
                output,
                error);

            Assert.Equal((int)OfficeImoToolExitCode.UnsupportedInput, exitCode);
            Assert.Contains("compression ratio", error.ToString(), StringComparison.OrdinalIgnoreCase);
            Assert.False(File.Exists(outputPath));
            Assert.Empty(Directory.EnumerateFiles(directory, ".officeimo-*.tmp"));
        } finally {
            Directory.Delete(directory, recursive: true);
        }
    }

    [Fact]
    public async Task ConvertAppliesConfiguredOpenXmlCharacterLimit() {
        string directory = CreateTestDirectory();
        string inputPath = Path.Combine(directory, "source.docx");
        string outputPath = Path.Combine(directory, "result.pdf");
        CreateSource(inputPath, ".docx");
        ReplaceZipEntry(
            inputPath,
            "word/document.xml",
            "<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><w:body>" +
            "<w:p><w:r><w:t>" + new string('A', 4096) + "</w:t></w:r></w:p><w:sectPr/>" +
            "</w:body></w:document>");
        await using var output = new MemoryStream();
        using var error = new StringWriter();

        try {
            int exitCode = await OfficeImoToolApp.RunAsync(
                ["convert", inputPath, "--output", outputPath, "--max-characters-in-part", "2048"],
                Stream.Null,
                output,
                error);

            Assert.Equal((int)OfficeImoToolExitCode.OperationFailed, exitCode);
            Assert.Contains("XmlException", error.ToString(), StringComparison.Ordinal);
            Assert.False(File.Exists(outputPath));
            Assert.Empty(Directory.EnumerateFiles(directory, ".officeimo-*.tmp"));
        } finally {
            Directory.Delete(directory, recursive: true);
        }
    }

    [Theory]
    [InlineData(".docx")]
    [InlineData(".xlsx")]
    [InlineData(".pptx")]
    public async Task ConvertAppliesConfiguredCharacterLimitToPackageMetadata(string extension) {
        string directory = CreateTestDirectory();
        string inputPath = Path.Combine(directory, "source" + extension);
        string outputPath = Path.Combine(directory, "result.pdf");
        CreateSource(inputPath, extension);
        ReplaceZipEntry(
            inputPath,
            "[Content_Types].xml",
            "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">" +
            "<Default Extension=\"" + new string('a', 512) + "\" ContentType=\"application/xml\" />" +
            "</Types>");
        await using var output = new MemoryStream();
        using var error = new StringWriter();

        try {
            int exitCode = await OfficeImoToolApp.RunAsync(
                ["convert", inputPath, "--output", outputPath, "--max-characters-in-part", "128"],
                Stream.Null,
                output,
                error);

            Assert.Equal((int)OfficeImoToolExitCode.UnsupportedInput, exitCode);
            Assert.Contains("content-types", error.ToString(), StringComparison.OrdinalIgnoreCase);
            Assert.False(File.Exists(outputPath));
            Assert.Empty(Directory.EnumerateFiles(directory, ".officeimo-*.tmp"));
        } finally {
            Directory.Delete(directory, recursive: true);
        }
    }

    [Fact]
    public async Task ConvertRejectsInputBeyondConfiguredLimit() {
        string directory = CreateTestDirectory();
        string inputPath = Path.Combine(directory, "source.docx");
        string outputPath = Path.Combine(directory, "result.pdf");
        CreateSource(inputPath, ".docx");
        await using var output = new MemoryStream();
        using var error = new StringWriter();

        try {
            int exitCode = await OfficeImoToolApp.RunAsync(
                ["convert", inputPath, "--output", outputPath, "--max-input-bytes", "64"],
                Stream.Null,
                output,
                error);

            Assert.Equal((int)OfficeImoToolExitCode.UnsupportedInput, exitCode);
            Assert.Contains("64", error.ToString(), StringComparison.Ordinal);
            Assert.False(File.Exists(outputPath));
            Assert.Empty(Directory.EnumerateFiles(directory, ".officeimo-*.tmp"));
        } finally {
            Directory.Delete(directory, recursive: true);
        }
    }

    [Fact]
    public async Task ConvertRejectsPdfBeyondConfiguredOutputLimit() {
        string directory = CreateTestDirectory();
        string inputPath = Path.Combine(directory, "source.docx");
        string outputPath = Path.Combine(directory, "result.pdf");
        CreateSource(inputPath, ".docx");
        await using var output = new MemoryStream();
        using var error = new StringWriter();

        try {
            int exitCode = await OfficeImoToolApp.RunAsync(
                ["convert", inputPath, "--output", outputPath, "--max-output-bytes", "64"],
                Stream.Null,
                output,
                error);

            Assert.Equal((int)OfficeImoToolExitCode.OutputFailed, exitCode);
            Assert.Contains("64", error.ToString(), StringComparison.Ordinal);
            Assert.False(File.Exists(outputPath));
            Assert.Empty(Directory.EnumerateFiles(directory, ".officeimo-*.tmp"));
        } finally {
            Directory.Delete(directory, recursive: true);
        }
    }

    [Fact]
    public async Task OutputLimitPreservesExistingDestinationWhenForceIsEnabled() {
        string directory = CreateTestDirectory();
        string inputPath = Path.Combine(directory, "source.docx");
        string outputPath = Path.Combine(directory, "result.pdf");
        CreateSource(inputPath, ".docx");
        await File.WriteAllTextAsync(outputPath, "keep");
        await using var output = new MemoryStream();
        using var error = new StringWriter();

        try {
            int exitCode = await OfficeImoToolApp.RunAsync(
                ["convert", inputPath, "--output", outputPath, "--force", "--max-output-bytes", "64"],
                Stream.Null,
                output,
                error);

            Assert.Equal((int)OfficeImoToolExitCode.OutputFailed, exitCode);
            Assert.Equal("keep", await File.ReadAllTextAsync(outputPath));
            Assert.Empty(Directory.EnumerateFiles(directory, ".officeimo-*.tmp"));
        } finally {
            Directory.Delete(directory, recursive: true);
        }
    }

    [Theory]
    [InlineData("--max-input-bytes")]
    [InlineData("--max-output-bytes")]
    [InlineData("--max-characters-in-part")]
    public async Task ConvertRejectsNonPositiveResourceLimits(string option) {
        await using var output = new MemoryStream();
        using var error = new StringWriter();

        int exitCode = await OfficeImoToolApp.RunAsync(
            ["convert", "source.docx", option, "0"],
            Stream.Null,
            output,
            error);

        Assert.Equal((int)OfficeImoToolExitCode.Usage, exitCode);
        Assert.Contains("positive integer", error.ToString(), StringComparison.Ordinal);
    }

    [Fact]
    public async Task OutputCommitFailsAtomicallyWhenDestinationAppearsWithoutForce() {
        string directory = Path.Combine(Path.GetTempPath(), "OfficeIMO.Tool.Tests", Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        string outputPath = Path.Combine(directory, "result.pdf");
        await File.WriteAllTextAsync(outputPath, "keep");

        try {
            string temporaryPath = Path.Combine(directory, ".officeimo-test.tmp");
            await File.WriteAllBytesAsync(temporaryPath, "%PDF-new"u8.ToArray());
            await Assert.ThrowsAsync<OfficePdfOutputExistsException>(() => OfficePdfCommand.CommitOutputAsync(
                temporaryPath, outputPath, force: false, CancellationToken.None));

            Assert.Equal("keep", await File.ReadAllTextAsync(outputPath));
            Assert.Empty(Directory.EnumerateFiles(directory, "*.tmp"));
        } finally {
            Directory.Delete(directory, recursive: true);
        }
    }

    [Fact]
    public async Task CancelledOutputCommitDoesNotReplaceDestinationOrLeaveStagingFiles() {
        string directory = Path.Combine(Path.GetTempPath(), "OfficeIMO.Tool.Tests", Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        string outputPath = Path.Combine(directory, "result.pdf");
        await File.WriteAllTextAsync(outputPath, "keep");
        using var cancellation = new CancellationTokenSource();
        cancellation.Cancel();

        try {
            string temporaryPath = Path.Combine(directory, ".officeimo-test.tmp");
            await File.WriteAllBytesAsync(temporaryPath, "%PDF-new"u8.ToArray());
            await Assert.ThrowsAnyAsync<OperationCanceledException>(() => OfficePdfCommand.CommitOutputAsync(
                temporaryPath, outputPath, force: true, cancellation.Token));

            Assert.Equal("keep", await File.ReadAllTextAsync(outputPath));
            Assert.Empty(Directory.EnumerateFiles(directory, "*.tmp"));
        } finally {
            Directory.Delete(directory, recursive: true);
        }
    }

    private static void CreateSource(string path, string extension) {
        switch (extension) {
            case ".docx":
                using (WordDocument document = WordDocument.Create(path)) {
                    document.AddParagraph("Tool Word conversion");
                    document.Save();
                }
                break;
            case ".xlsx":
                using (ExcelDocument workbook = ExcelDocument.Create(path)) {
                    ExcelSheet sheet = workbook.AddWorksheet("Summary");
                    sheet.CellValue(1, 1, "Tool Excel conversion");
                    workbook.Save();
                }
                break;
            case ".pptx":
                using (PowerPointPresentation presentation = PowerPointPresentation.Create(path)) {
                    presentation.AddSlide().AddTextBox("Tool PowerPoint conversion");
                    presentation.Save();
                }
                break;
            default:
                throw new ArgumentOutOfRangeException(nameof(extension));
        }
    }

    private static string CreateTestDirectory() {
        string directory = Path.Combine(Path.GetTempPath(), "OfficeIMO.Tool.Tests", Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        return directory;
    }

    private static void CreateCompressedDocxBomb(string path) {
        CreateSource(path, ".docx");
        using ZipArchive archive = ZipFile.Open(path, ZipArchiveMode.Update);
        ZipArchiveEntry documentPart = archive.GetEntry("word/document.xml")
            ?? throw new InvalidDataException("The generated DOCX has no document part.");
        documentPart.Delete();
        ZipArchiveEntry replacement = archive.CreateEntry("word/document.xml", CompressionLevel.SmallestSize);
        using var writer = new StreamWriter(replacement.Open(), new UTF8Encoding(false));
        writer.Write("<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><w:body><w:p><w:r><w:t>");
        string compressibleBlock = new string('A', 1024 * 1024);
        for (int index = 0; index < 32; index++) writer.Write(compressibleBlock);
        writer.Write("</w:t></w:r></w:p><w:sectPr/></w:body></w:document>");
    }

    private static void ReplaceZipEntry(string path, string entryName, string content) {
        using ZipArchive archive = ZipFile.Open(path, ZipArchiveMode.Update);
        ZipArchiveEntry entry = archive.GetEntry(entryName)
            ?? throw new InvalidDataException("The generated package has no " + entryName + " part.");
        entry.Delete();
        ZipArchiveEntry replacement = archive.CreateEntry(entryName, CompressionLevel.Optimal);
        using var writer = new StreamWriter(replacement.Open(), new UTF8Encoding(false));
        writer.Write(content);
    }
}
