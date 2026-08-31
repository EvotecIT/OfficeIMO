using Xunit;
using OfficeIMO.Tool.Commands.Tabular;
using OfficeIMO.Excel;
using System.IO.Compression;

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

    [Theory]
    [InlineData(".xlsx")]
    [InlineData(".xlsb")]
    [InlineData(".xls")]
    public void SheetMetadataDiscoveryDoesNotDecodeSharedStringsOrStyles(string workbookExtension) {
        string directory = CreateTestDirectory();
        string workbookPath = Path.Combine(directory, "metadata" + workbookExtension);
        try {
            using (ExcelDocument document = ExcelDocument.Create()) {
                ExcelSheet sheet = document.AddWorksheet("Inventory");
                sheet.CellAt(1, 1).SetValue("Alpha");
                sheet.CellAt(2, 1).SetValue("Beta");
                sheet.CellAt(3, 1).SetValue("Gamma");
                document.Save(workbookPath);
            }

            IReadOnlyList<string> names = ExcelDocument.GetSheetNames(
                workbookPath,
                new ExcelReadOptions {
                    MaxSharedStringItems = 1,
                    MaxSharedStringCharacters = 1,
                    MaxSharedStringItemCharacters = 1
                });

            Assert.Equal(new[] { "Inventory" }, names);
        } finally {
            Directory.Delete(directory, recursive: true);
        }
    }

    [Fact]
    public void SheetMetadataDiscoveryEnforcesCountAndCancellationLimits() {
        string directory = CreateTestDirectory();
        string workbookPath = Path.Combine(directory, "limits.xlsx");
        try {
            using (ExcelDocument document = ExcelDocument.Create()) {
                document.AddWorksheet("First");
                document.AddWorksheet("Second");
                document.Save(workbookPath);
            }

            InvalidDataException countError = Assert.Throws<InvalidDataException>(() =>
                ExcelDocument.GetSheetNames(
                    workbookPath,
                    new ExcelReadOptions { MaxWorksheets = 1 }));
            Assert.Contains("worksheet", countError.Message, StringComparison.OrdinalIgnoreCase);

            Assert.Throws<InvalidDataException>(() =>
                ExcelDocument.GetSheetNames(
                    workbookPath,
                    new ExcelReadOptions { MaxMetadataPartBytes = 32 }));

            using var cancellation = new CancellationTokenSource();
            cancellation.Cancel();
            Assert.ThrowsAny<OperationCanceledException>(() =>
                ExcelDocument.GetSheetNames(
                    workbookPath,
                    new ExcelReadOptions { CancellationToken = cancellation.Token }));
        } finally {
            Directory.Delete(directory, recursive: true);
        }
    }

    [Fact]
    public async Task SheetMetadataDiscoverySupportsPercentEncodedOpcTargets() {
        string directory = CreateTestDirectory();
        string workbookPath = Path.Combine(directory, "encoded-target.xlsx");
        try {
            using (ExcelDocument document = ExcelDocument.Create()) {
                document.AddWorksheet("Encoded Target");
                document.Save(workbookPath);
            }

            await RenameZipEntryAsync(
                workbookPath,
                "xl/worksheets/sheet1.xml",
                "xl/worksheets/My Sheet.xml");
            await ReplaceZipEntryTextAsync(
                workbookPath,
                "xl/_rels/workbook.xml.rels",
                static xml => xml.Replace(
                    "worksheets/sheet1.xml",
                    "worksheets/My%20Sheet.xml",
                    StringComparison.Ordinal));
            await ReplaceZipEntryTextAsync(
                workbookPath,
                "[Content_Types].xml",
                static xml => xml.Replace(
                    "/xl/worksheets/sheet1.xml",
                    "/xl/worksheets/My%20Sheet.xml",
                    StringComparison.Ordinal));

            Assert.Equal(
                new[] { "Encoded Target" },
                ExcelDocument.GetSheetNames(workbookPath));
        } finally {
            Directory.Delete(directory, recursive: true);
        }
    }

    [Fact]
    public void XlsbMetadataLimitDoesNotApplyToUnusedWorksheetParts() {
        string directory = CreateTestDirectory();
        string workbookPath = Path.Combine(directory, "large-sheet.xlsb");
        const int MetadataLimit = 2 * 1024;
        try {
            using (ExcelDocument document = ExcelDocument.Create()) {
                ExcelSheet sheet = document.AddWorksheet("Inventory");
                for (int row = 1; row <= 500; row++) {
                    sheet.CellAt(row, 1).SetValue("Value " + row);
                }
                document.Save(workbookPath);
            }
            using (ZipArchive archive = ZipFile.OpenRead(workbookPath)) {
                ZipArchiveEntry worksheet = archive.Entries.Single(
                    static entry => entry.FullName.EndsWith("/worksheets/sheet1.bin", StringComparison.OrdinalIgnoreCase));
                Assert.True(worksheet.Length > MetadataLimit);
            }

            Assert.Equal(
                new[] { "Inventory" },
                ExcelDocument.GetSheetNames(
                    workbookPath,
                    new ExcelReadOptions { MaxMetadataPartBytes = MetadataLimit }));
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
    public async Task TabularCliKeepsInputAndOutputDelimitersIndependent() {
        string directory = CreateTestDirectory();
        string inputPath = Path.Combine(directory, "pipe-input.csv");
        string defaultOutputPath = Path.Combine(directory, "default-output.csv");
        string explicitOutputPath = Path.Combine(directory, "explicit-output.csv");
        await File.WriteAllTextAsync(inputPath, "Id|Name\n1|Alpha\n");

        try {
            (int defaultExit, _, string defaultError) = await RunAsync(
                "tabular", "convert", inputPath, defaultOutputPath, "--delimiter", "|");
            Assert.Equal((int)OfficeImoToolExitCode.Success, defaultExit);
            Assert.Equal(string.Empty, defaultError);
            Assert.Contains("Id,Name", await File.ReadAllTextAsync(defaultOutputPath), StringComparison.Ordinal);

            (int explicitExit, _, string explicitError) = await RunAsync(
                "tabular", "convert", inputPath, explicitOutputPath,
                "--delimiter", "|", "--output-delimiter", ";");
            Assert.Equal((int)OfficeImoToolExitCode.Success, explicitExit);
            Assert.Equal(string.Empty, explicitError);
            Assert.Contains("Id;Name", await File.ReadAllTextAsync(explicitOutputPath), StringComparison.Ordinal);
        } finally {
            Directory.Delete(directory, recursive: true);
        }
    }

    [Theory]
    [InlineData("--delimiter")]
    [InlineData("--output-delimiter")]
    public async Task TabularCliRejectsTheCsvQuoteCharacterAsDelimiter(string option) {
        string directory = CreateTestDirectory();
        string inputPath = Path.Combine(directory, "input.csv");
        string outputPath = Path.Combine(directory, "output.csv");
        await File.WriteAllTextAsync(inputPath, "Id,Name\n1,Alpha\n");

        try {
            (int exitCode, _, string error) = await RunAsync(
                "tabular", "convert", inputPath, outputPath, option, "\"");

            Assert.Equal((int)OfficeImoToolExitCode.Usage, exitCode);
            Assert.Contains("quote", error, StringComparison.OrdinalIgnoreCase);
            Assert.False(File.Exists(outputPath));
        } finally {
            Directory.Delete(directory, recursive: true);
        }
    }

    [Theory]
    [InlineData("--delimiter", "\r")]
    [InlineData("--delimiter", "\n")]
    [InlineData("--output-delimiter", "\r")]
    [InlineData("--output-delimiter", "\n")]
    public async Task TabularCliRejectsRecordSeparatorsAsDelimiters(string option, string delimiter) {
        string directory = CreateTestDirectory();
        string inputPath = Path.Combine(directory, "input.csv");
        string outputPath = Path.Combine(directory, "output.csv");
        await File.WriteAllTextAsync(inputPath, "Id,Name\n1,Alpha\n");

        try {
            (int exitCode, _, string error) = await RunAsync(
                "tabular", "convert", inputPath, outputPath, option, delimiter);

            Assert.Equal((int)OfficeImoToolExitCode.Usage, exitCode);
            Assert.Contains("record separator", error, StringComparison.OrdinalIgnoreCase);
            Assert.False(File.Exists(outputPath));
        } finally {
            Directory.Delete(directory, recursive: true);
        }
    }

    [Fact]
    public async Task TabularCliClassifiesLockedInputAsDocumentedIoFailure() {
        if (!OperatingSystem.IsWindows()) return;

        string directory = CreateTestDirectory();
        string inputPath = Path.Combine(directory, "locked.xlsx");
        using (ExcelDocument document = ExcelDocument.Create()) {
            document.AddWorksheet("Data");
            document.Save(inputPath);
        }

        try {
            await using (var locked = new FileStream(
                inputPath,
                FileMode.Open,
                FileAccess.ReadWrite,
                FileShare.None)) {
                (int exitCode, _, string error) = await RunAsync(
                    "tabular", "schema", inputPath);

                Assert.Equal((int)OfficeImoToolExitCode.UnsupportedInput, exitCode);
                Assert.Contains("I/O failed", error, StringComparison.OrdinalIgnoreCase);
            }
        } finally {
            Directory.Delete(directory, recursive: true);
        }
    }

    [Fact]
    public async Task TabularCliHonorsTheTsvDelimiterWhenFieldsContainCommas() {
        string directory = CreateTestDirectory();
        string tsvPath = Path.Combine(directory, "input.tsv");
        await File.WriteAllTextAsync(tsvPath, "Person,Display\tAge\nAlice,Smith\t42\n");

        try {
            (int exitCode, string output, string error) = await RunAsync(
                "tabular", "schema", tsvPath);

            Assert.Equal((int)OfficeImoToolExitCode.Success, exitCode);
            Assert.Equal(string.Empty, error);
            Assert.Contains("0\tPerson,Display\tSystem.String", output, StringComparison.Ordinal);
            Assert.Contains("1\tAge\t", output, StringComparison.Ordinal);
        } finally {
            Directory.Delete(directory, recursive: true);
        }
    }

    [Fact]
    public async Task TabularCliEscapesSchemaControlCharactersWithoutAddingRows() {
        string directory = CreateTestDirectory();
        string csvPath = Path.Combine(directory, "input.csv");
        await File.WriteAllTextAsync(
            csvPath,
            "\"Name\tPart\",\"Line\r\nBreak\",\"Slash\\Path\"\r\nOne,Two,Three\r\n");

        try {
            (int exitCode, string output, string error) = await RunAsync(
                "tabular", "schema", csvPath);

            Assert.Equal((int)OfficeImoToolExitCode.Success, exitCode);
            Assert.Equal(string.Empty, error);
            Assert.Contains("0\tName\\tPart\tSystem.String", output, StringComparison.Ordinal);
            Assert.Contains("1\tLine\\r\\nBreak\tSystem.String", output, StringComparison.Ordinal);
            Assert.Contains("2\tSlash\\\\Path\tSystem.String", output, StringComparison.Ordinal);
            Assert.Equal(4, output.Split('\n', StringSplitOptions.RemoveEmptyEntries).Length);
        } finally {
            Directory.Delete(directory, recursive: true);
        }
    }

    [Fact]
    public async Task TabularCliListsSheetMetadataWithoutOpeningTheFirstWorksheet() {
        string directory = CreateTestDirectory();
        string workbookPath = Path.Combine(directory, "broken-first-sheet.xlsx");
        using (ExcelDocument document = ExcelDocument.Create()) {
            document.AddWorksheet("BrokenData");
            document.AddWorksheet("MetadataOnly");
            document.Save(workbookPath);
        }
        using (ZipArchive archive = ZipFile.Open(workbookPath, ZipArchiveMode.Update)) {
            ZipArchiveEntry worksheet = archive.GetEntry("xl/worksheets/sheet1.xml")
                ?? throw new InvalidDataException("The test workbook does not contain its first worksheet part.");
            worksheet.Delete();
            worksheet = archive.CreateEntry("xl/worksheets/sheet1.xml");
            await using Stream stream = worksheet.Open();
            await using var writer = new StreamWriter(stream, new System.Text.UTF8Encoding(false));
            await writer.WriteAsync("<worksheet");
        }
        await ReplaceZipEntryTextAsync(
            workbookPath,
            "xl/workbook.xml",
            static xml => xml.Replace(
                "name=\"MetadataOnly\"",
                "name=\"MetadataOnly&#10;Injected\"",
                StringComparison.Ordinal));

        try {
            (int exitCode, string output, string error) = await RunAsync(
                "tabular", "sheets", workbookPath);

            Assert.Equal((int)OfficeImoToolExitCode.Success, exitCode);
            Assert.Equal(string.Empty, error);
            Assert.Equal(
                new[] { "BrokenData", "MetadataOnly\\nInjected" },
                output.Split(new[] { "\r\n", "\n" }, StringSplitOptions.RemoveEmptyEntries));
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

    [Fact]
    public async Task TabularCliRejectsOptionsThatWorkbookToWorkbookConversionWouldIgnore() {
        string directory = CreateTestDirectory();
        string inputPath = Path.Combine(directory, "input.xlsx");
        using (ExcelDocument document = ExcelDocument.Create()) {
            document.AddWorksheet("Data");
            document.Save(inputPath);
        }

        try {
            foreach (string[] option in new[] {
                         new[] { "--no-header" },
                         new[] { "--delimiter", ";" },
                         new[] { "--output-delimiter", ";" }
                     }) {
                string outputPath = Path.Combine(directory, Guid.NewGuid().ToString("N") + ".xlsb");
                string[] arguments = new[] { "tabular", "convert", inputPath, outputPath }
                    .Concat(option)
                    .ToArray();
                (int exitCode, _, string error) = await RunAsync(arguments);

                Assert.Equal((int)OfficeImoToolExitCode.UnsupportedInput, exitCode);
                Assert.Contains(option[0], error, StringComparison.Ordinal);
                Assert.False(File.Exists(outputPath));
            }
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

    private static async Task ReplaceZipEntryTextAsync(
        string packagePath,
        string entryName,
        Func<string, string> replace) {
        using ZipArchive archive = ZipFile.Open(packagePath, ZipArchiveMode.Update);
        ZipArchiveEntry entry = archive.GetEntry(entryName)
            ?? throw new InvalidDataException("The test package does not contain '" + entryName + "'.");
        string contents;
        await using (Stream input = entry.Open()) {
            using var reader = new StreamReader(input, Encoding.UTF8, detectEncodingFromByteOrderMarks: true, leaveOpen: true);
            contents = await reader.ReadToEndAsync();
        }
        entry.Delete();
        ZipArchiveEntry replacement = archive.CreateEntry(entryName);
        await using Stream output = replacement.Open();
        await using var writer = new StreamWriter(output, new UTF8Encoding(false));
        await writer.WriteAsync(replace(contents));
    }

    private static async Task RenameZipEntryAsync(
        string packagePath,
        string existingEntryName,
        string replacementEntryName) {
        using ZipArchive archive = ZipFile.Open(packagePath, ZipArchiveMode.Update);
        ZipArchiveEntry entry = archive.GetEntry(existingEntryName)
            ?? throw new InvalidDataException("The test package does not contain '" + existingEntryName + "'.");
        using var buffer = new MemoryStream();
        await using (Stream input = entry.Open()) {
            await input.CopyToAsync(buffer);
        }
        entry.Delete();
        ZipArchiveEntry replacement = archive.CreateEntry(replacementEntryName);
        buffer.Position = 0;
        await using Stream output = replacement.Open();
        await buffer.CopyToAsync(output);
    }

    private static async Task<(int ExitCode, string Output, string Error)> RunAsync(params string[] args) {
        await using var input = new MemoryStream();
        await using var output = new MemoryStream();
        using var error = new StringWriter();
        int exitCode = await OfficeImoToolApp.RunAsync(args, input, output, error);
        return (exitCode, Encoding.UTF8.GetString(output.ToArray()), error.ToString());
    }
}
