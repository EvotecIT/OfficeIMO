using OfficeIMO.Excel;
using OfficeIMO.PowerPoint;
using OfficeIMO.Tool.Commands.Convert;
using OfficeIMO.Word;
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
                ["convert", inputPath, "--output", outputPath],
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
    public async Task OutputCommitFailsAtomicallyWhenDestinationAppearsWithoutForce() {
        string directory = Path.Combine(Path.GetTempPath(), "OfficeIMO.Tool.Tests", Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        string outputPath = Path.Combine(directory, "result.pdf");
        await File.WriteAllTextAsync(outputPath, "keep");

        try {
            await Assert.ThrowsAsync<OfficePdfOutputExistsException>(() => OfficePdfCommand.CommitOutputAsync(
                "%PDF-new"u8.ToArray(), outputPath, force: false, CancellationToken.None));

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
            await Assert.ThrowsAnyAsync<OperationCanceledException>(() => OfficePdfCommand.CommitOutputAsync(
                "%PDF-new"u8.ToArray(), outputPath, force: true, cancellation.Token));

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
}
