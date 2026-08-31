using Xunit;
using OfficeIMO.Tool.Commands.Tabular;

namespace OfficeIMO.Tool.Tests;

public sealed class TabularCommandTests {
    [Theory]
    [InlineData(".xlsx")]
    [InlineData(".xlsb")]
    [InlineData(".xls")]
    public async Task TabularCliConvertsCsvAndListsTheGeneratedSheet(string workbookExtension) {
        string directory = CreateTestDirectory();
        string csvPath = Path.Combine(directory, "input.csv");
        string workbookPath = Path.Combine(directory, "output" + workbookExtension);
        await File.WriteAllTextAsync(csvPath, "Id,Name,Amount\n1,Alpha,12.5\n2,Beta,20\n");

        try {
            (int convertExit, string convertOutput, string convertError) = await RunAsync(
                "tabular", "convert", csvPath, workbookPath);

            Assert.Equal((int)OfficeImoToolExitCode.Success, convertExit);
            Assert.True(File.Exists(workbookPath));
            Assert.Contains("Converted to", convertOutput, StringComparison.Ordinal);
            Assert.Equal(string.Empty, convertError);

            (int sheetsExit, string sheetsOutput, string sheetsError) = await RunAsync(
                "tabular", "sheets", workbookPath);

            Assert.Equal((int)OfficeImoToolExitCode.Success, sheetsExit);
            Assert.Contains("Data", sheetsOutput, StringComparison.Ordinal);
            Assert.Equal(string.Empty, sheetsError);
        } finally {
            Directory.Delete(directory, recursive: true);
        }
    }

    [Fact]
    public async Task TabularCliReportsInferredSchemaAndStreamsWorkbookBackToTsv() {
        string directory = CreateTestDirectory();
        string csvPath = Path.Combine(directory, "input.csv");
        string workbookPath = Path.Combine(directory, "output.xlsx");
        string tsvPath = Path.Combine(directory, "roundtrip.tsv");
        await File.WriteAllTextAsync(csvPath, "Id,Name,Active\n1,Alpha,true\n2,Beta,false\n");

        try {
            Assert.Equal(
                (int)OfficeImoToolExitCode.Success,
                (await RunAsync("tabular", "convert", csvPath, workbookPath)).ExitCode);

            (int schemaExit, string schemaOutput, string schemaError) = await RunAsync(
                "tabular", "schema", workbookPath, "--sheet", "Data");
            Assert.Equal((int)OfficeImoToolExitCode.Success, schemaExit);
            Assert.Contains("ordinal\tname\ttype", schemaOutput, StringComparison.Ordinal);
            Assert.Contains("0\tId\tSystem.Double", schemaOutput, StringComparison.Ordinal);
            Assert.Contains("2\tActive\tSystem.Boolean", schemaOutput, StringComparison.Ordinal);
            Assert.Equal(string.Empty, schemaError);

            (int roundtripExit, _, string roundtripError) = await RunAsync(
                "tabular", "convert", workbookPath, tsvPath, "--sheet", "Data");
            Assert.Equal((int)OfficeImoToolExitCode.Success, roundtripExit);
            Assert.Equal(string.Empty, roundtripError);
            string roundtrip = await File.ReadAllTextAsync(tsvPath);
            Assert.Contains("Id\tName\tActive", roundtrip, StringComparison.Ordinal);
            Assert.Contains("1\tAlpha\tTrue", roundtrip, StringComparison.OrdinalIgnoreCase);
        } finally {
            Directory.Delete(directory, recursive: true);
        }
    }

    [Fact]
    public async Task TabularCliDoesNotReplaceExistingOutputWithoutForce() {
        string directory = CreateTestDirectory();
        string csvPath = Path.Combine(directory, "input.csv");
        string outputPath = Path.Combine(directory, "output.xlsx");
        await File.WriteAllTextAsync(csvPath, "Id\n1\n");
        await File.WriteAllTextAsync(outputPath, "sentinel");

        try {
            (int exitCode, _, string error) = await RunAsync(
                "tabular", "convert", csvPath, outputPath);

            Assert.Equal((int)OfficeImoToolExitCode.OutputFailed, exitCode);
            Assert.Contains("--force", error, StringComparison.Ordinal);
            Assert.Equal("sentinel", await File.ReadAllTextAsync(outputPath));
        } finally {
            Directory.Delete(directory, recursive: true);
        }
    }

    [Fact]
    public async Task TabularCliCopiesTheSameWorkbookFormatThroughAtomicOutput() {
        string directory = CreateTestDirectory();
        string csvPath = Path.Combine(directory, "input.csv");
        string sourcePath = Path.Combine(directory, "source.xls");
        string copyPath = Path.Combine(directory, "copy.xls");
        await File.WriteAllTextAsync(csvPath, "Id,Name\n1,Alpha\n");

        try {
            Assert.Equal(
                (int)OfficeImoToolExitCode.Success,
                (await RunAsync("tabular", "convert", csvPath, sourcePath)).ExitCode);

            (int copyExit, _, string copyError) = await RunAsync(
                "tabular", "convert", sourcePath, copyPath);

            Assert.Equal((int)OfficeImoToolExitCode.Success, copyExit);
            Assert.Equal(string.Empty, copyError);
            Assert.Equal(await File.ReadAllBytesAsync(sourcePath), await File.ReadAllBytesAsync(copyPath));
        } finally {
            Directory.Delete(directory, recursive: true);
        }
    }

    [Fact]
    public async Task TabularCliDelimitedConversionPreservesLexicalValues() {
        string directory = CreateTestDirectory();
        string csvPath = Path.Combine(directory, "input.csv");
        string tsvPath = Path.Combine(directory, "output.tsv");
        await File.WriteAllTextAsync(
            csvPath,
            "Code,Flag,Date\n00123,TRUE,01/02/2026\n00007,false,2026-08-31\n");

        try {
            (int exitCode, _, string error) = await RunAsync(
                "tabular", "convert", csvPath, tsvPath);

            Assert.Equal((int)OfficeImoToolExitCode.Success, exitCode);
            Assert.Equal(string.Empty, error);
            string converted = await File.ReadAllTextAsync(tsvPath);
            Assert.Contains("00123\tTRUE\t01/02/2026", converted, StringComparison.Ordinal);
            Assert.Contains("00007\tfalse\t2026-08-31", converted, StringComparison.Ordinal);
        } finally {
            Directory.Delete(directory, recursive: true);
        }
    }

    [Fact]
    public async Task SameFormatCopyObservesCancellationDuringStagedIo() {
        string directory = CreateTestDirectory();
        string sourcePath = Path.Combine(directory, "source.xlsx");
        string stagedPath = Path.Combine(directory, "staged.xlsx");
        await File.WriteAllBytesAsync(sourcePath, new byte[1024 * 1024]);
        using var cancellation = new CancellationTokenSource();
        cancellation.Cancel();

        try {
            await Assert.ThrowsAnyAsync<OperationCanceledException>(() =>
                TabularCommand.CopyFileAsync(sourcePath, stagedPath, cancellation.Token));
        } finally {
            Directory.Delete(directory, recursive: true);
        }
    }

    [Fact]
    public async Task TabularCliRejectsUnsupportedWorkbookOutputAndCsvSheetSelection() {
        string directory = CreateTestDirectory();
        string csvPath = Path.Combine(directory, "input.csv");
        await File.WriteAllTextAsync(csvPath, "Id\n1\n");

        try {
            (int outputExit, _, string outputError) = await RunAsync(
                "tabular", "convert", csvPath, Path.Combine(directory, "output.xlsm"));
            Assert.Equal((int)OfficeImoToolExitCode.UnsupportedInput, outputExit);
            Assert.Contains("XLSX", outputError, StringComparison.Ordinal);

            (int sheetExit, _, string sheetError) = await RunAsync(
                "tabular", "schema", csvPath, "--sheet", "Data");
            Assert.Equal((int)OfficeImoToolExitCode.UnsupportedInput, sheetExit);
            Assert.Contains("workbook input", sheetError, StringComparison.Ordinal);
        } finally {
            Directory.Delete(directory, recursive: true);
        }
    }

    private static string CreateTestDirectory() {
        string directory = Path.Combine(
            Path.GetTempPath(),
            "OfficeIMO.Tool.Tests",
            Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        return directory;
    }

    private static async Task<(int ExitCode, string Output, string Error)> RunAsync(params string[] args) {
        await using var input = new MemoryStream();
        await using var output = new MemoryStream();
        using var error = new StringWriter();
        int exitCode = await OfficeImoToolApp.RunAsync(args, input, output, error);
        return (exitCode, Encoding.UTF8.GetString(output.ToArray()), error.ToString());
    }
}
