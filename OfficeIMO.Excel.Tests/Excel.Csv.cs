using System.Data;
using System.Globalization;
using System.Text;
using OfficeIMO.CSV;
using OfficeIMO.Excel;
using OfficeIMO.Excel.Csv;
using Xunit;

namespace OfficeIMO.Excel.Tests;

public class ExcelCsvExtensionsTests {
    [Fact]
    public void TsvFilesConvertToWorkbooksWithQuotesAsData() {
        string directory = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        try {
            string source = Path.Combine(directory, "notes.tsv");
            File.WriteAllText(source, "Id\tNote\n1\t\"open\n2\tsays \"hi\", twice\n");
            using var imported = ExcelDocument.Create();
            imported.ImportCsvFile(source);
            using ExcelDocument converted = CsvDocument.Load(source).ToExcel();
            foreach (ExcelSheet sheet in new[] { imported["Import"], converted["Import"] }) {
                Assert.True(sheet.TryGetCellText(2, 2, out string? first));
                Assert.Equal("\"open", first);
                Assert.True(sheet.TryGetCellText(3, 2, out string? second));
                Assert.Equal("says \"hi\", twice", second);
            }

            string exported = Path.Combine(directory, "export.tsv");
            imported["Import"].SaveAsCsv(exported);
            Assert.Equal("Id\tNote\n1\t\"open\n2\tsays \"hi\", twice\n", File.ReadAllText(exported));
        } finally { Directory.Delete(directory, recursive: true); }
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task CanonicalConversionsRetainValuesReportsAndCsvSavePolicies(bool asynchronous) {
        using var workbook = ExcelDocument.Create();
        ExcelSheet sheet = workbook.AddSheet("Data");
        sheet.Cell(1, 1).SetValue("Name");
        sheet.Cell(2, 1).SetValue("Alpha");
        var options = new ExcelCsvExportOptions { Range = "A1:A2" };
        OfficeConversionResult<CsvDocument> converted = await sheet.ToCsvAsync(options);
        Assert.Equal("Alpha", Assert.Single(converted.Value.AsEnumerable())[0]);
        Assert.Null(converted.Report.OutputPath);
        Assert.Contains(converted.Report.Diagnostics, diagnostic => diagnostic.LossKind == OfficeConversionLossKind.Unassessed);
        OfficeConversionResult<ExcelDocument> restored = await converted.Value.ToExcelAsync(new CsvExcelExportOptions {
            ImportOptions = new ExcelCsvImportOptions { ReaderOptions = new CsvReaderOptions { InferSchema = false } }
        });
        using (restored.Value) {
            Assert.False(restored.Report.HasLoss);
            Assert.True(restored.Value["Import"].TryGetCellText(2, 1, out string? value));
            Assert.Equal("Alpha", value);
        }
        using var output = new MemoryStream();
        OfficeConversionException loss = await Assert.ThrowsAsync<OfficeConversionException>(() =>
            sheet.SaveAsCsvAsync(output, new ExcelCsvExportOptions { FailOnLoss = true }));
        Assert.True(loss.Report.HasLoss);
        Assert.Equal(0, output.Length);
        Assert.True(output.CanWrite);

        string directory = Path.Combine(Path.GetTempPath(), "officeimo-csv-save-report-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        try {
            string csvPath = Path.Combine(directory, "output.csv");
            string csvDestination = Path.Combine(directory, ".", "output.csv");
            OfficeConversionReport csvReport = asynchronous
                ? await sheet.SaveAsCsvAsync(csvDestination, options)
                : sheet.SaveAsCsv(csvDestination, options);
            Assert.Equal(Path.GetFullPath(csvPath), csvReport.OutputPath);
            Assert.Contains(csvReport.Diagnostics, diagnostic => diagnostic.Code == "EXCEL_CSV_VALUES_ONLY"
                && diagnostic.LossKind == OfficeConversionLossKind.Unassessed);
            Assert.Equal("Alpha", Assert.Single(CsvDocument.Load(csvPath).AsEnumerable())[0]);
            using var csvStream = new MemoryStream();
            OfficeConversionReport csvStreamReport = asynchronous
                ? await sheet.SaveAsCsvAsync(csvStream, options)
                : sheet.SaveAsCsv(csvStream, options);
            Assert.Null(csvStreamReport.OutputPath);
            Assert.True(csvStream.CanWrite);
            Assert.Equal(csvReport.Diagnostics.Select(diagnostic => (diagnostic.Code, diagnostic.LossKind)),
                csvStreamReport.Diagnostics.Select(diagnostic => (diagnostic.Code, diagnostic.LossKind)));
            csvStream.Position = 0;
            Assert.Equal("Alpha", Assert.Single(CsvDocument.Load(csvStream).AsEnumerable())[0]);

            var excelOptions = new CsvExcelExportOptions {
                ImportOptions = new ExcelCsvImportOptions { ReaderOptions = new CsvReaderOptions { InferSchema = false } }
            };
            string excelPath = Path.Combine(directory, "output.xlsx");
            OfficeConversionReport excelReport = asynchronous
                ? await converted.Value.SaveAsExcelAsync(Path.Combine(directory, ".", "output.xlsx"), excelOptions)
                : converted.Value.SaveAsExcel(Path.Combine(directory, ".", "output.xlsx"), excelOptions);
            Assert.Equal(Path.GetFullPath(excelPath), excelReport.OutputPath);
            Assert.False(excelReport.HasLoss);
            using ExcelDocument saved = ExcelDocument.Load(excelPath);
            Assert.True(saved["Import"].TryGetCellText(2, 1, out string? savedValue));
            Assert.Equal("Alpha", savedValue);
            using var excelStream = new MemoryStream();
            OfficeConversionReport excelStreamReport = asynchronous
                ? await converted.Value.SaveAsExcelAsync(excelStream, excelOptions)
                : converted.Value.SaveAsExcel(excelStream, excelOptions);
            Assert.Null(excelStreamReport.OutputPath);
            Assert.True(excelStream.CanWrite);
            Assert.False(excelStreamReport.HasLoss);
            excelStream.Position = 0;
            using ExcelDocument streamed = ExcelDocument.Load(excelStream);
            Assert.True(streamed["Import"].TryGetCellText(2, 1, out string? streamedValue));
            Assert.Equal("Alpha", streamedValue);
        } finally {
            Directory.Delete(directory, recursive: true);
        }

        string path = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".csv");
        File.WriteAllText(path, "retained");
        try {
            OfficeConversionException conflict = Assert.Throws<OfficeConversionException>(() =>
                sheet.SaveAsCsv(path, new ExcelCsvExportOptions { SaveOptions = new CsvSaveOptions { NoClobber = true } }));
            Assert.IsAssignableFrom<IOException>(conflict.InnerException);
            Assert.Contains(conflict.Report.Diagnostics, diagnostic => diagnostic.LossKind == OfficeConversionLossKind.Failure);
            Assert.Equal("retained", File.ReadAllText(path));
        } finally { File.Delete(path); }
    }

    [Fact]
    public void CsvImportChecksMethodCancellationBeforeCreatingReaders() {
        using var cancellation = new CancellationTokenSource();
        cancellation.Cancel();
        var loadOptions = new CsvReaderOptions();
        var options = new ExcelCsvImportOptions { ReaderOptions = loadOptions };
        CsvDocument csv = CsvDocument.Parse("Name\r\nAlpha");
        using var stream = new MemoryStream();
        using var document = ExcelDocument.Create(stream);
        ExcelSheet sheet = document.AddSheet("Existing");

        Assert.Throws<OperationCanceledException>(() =>
            document.ImportCsv(csv, options, cancellation.Token));
        Assert.Throws<OperationCanceledException>(() =>
            sheet.ImportCsvText("Name\r\nAlpha", options, cancellation.Token));

        Assert.False(loadOptions.CancellationToken.CanBeCanceled);
        Assert.Single(document.Sheets);
    }

    [Fact]
    public void EmptyCsvImportReturnsAnEmptyRange() {
        using var stream = new MemoryStream();
        using var document = ExcelDocument.Create(stream);

        ExcelCsvImportResult result = document.ImportCsvText(
            string.Empty,
            new ExcelCsvImportOptions { SheetName = "Empty" });

        Assert.Equal("Empty", result.SheetName);
        Assert.Null(result.TableName);
        Assert.Equal(string.Empty, result.Range);
        Assert.Equal("A1:A1", document["Empty"].UsedRange.Address);
    }

    [Fact]
    public void CsvTextImportPreservesUnicodeRegardlessOfFileEncodingOption() {
        var options = new ExcelCsvImportOptions {
            SheetName = "Imported",
            ReaderOptions = new CsvReaderOptions { Encoding = Encoding.ASCII, InferSchema = false }
        };
        using var stream = new MemoryStream();
        using var document = ExcelDocument.Create(stream);

        document.ImportCsvText("Name\r\nélève", options);
        ExcelSheet existing = document.AddSheet("Existing");
        existing.ImportCsvText("Name\r\n東京", options);

        Assert.True(document["Imported"].TryGetCellText(2, 1, out string? imported));
        Assert.Equal("élève", imported);
        Assert.True(existing.TryGetCellText(2, 1, out string? existingValue));
        Assert.Equal("東京", existingValue);
    }

    [Fact]
    public void CsvTextImportReportsTheEffectiveDelimiterTextCharacter() {
        var options = new ExcelCsvImportOptions {
            SheetName = "Imported",
            ReaderOptions = new CsvReaderOptions { DelimiterText = "||", InferSchema = false }
        };
        using var stream = new MemoryStream();
        using var document = ExcelDocument.Create(stream);

        ExcelCsvImportResult result = document.ImportCsvText("Name||Value\r\nAlpha||1", options);

        Assert.Equal('|', result.Delimiter);
        Assert.Equal("||", result.DelimiterText);
        Assert.True(document["Imported"].TryGetCellText(2, 2, out string? value));
        Assert.Equal("1", value);
    }

    [Fact]
    public void CsvDocumentImportReportsFallbackForEmptyDelimiterText() {
        var csv = CsvDocument.Parse("Name,Value\nAlpha,1\n", new CsvLoadOptions { DelimiterText = "" });
        using var document = ExcelDocument.Create();
        ExcelCsvImportResult result = document.ImportCsv(csv);
        Assert.Equal(",", result.DelimiterText);
        Assert.Equal(',', result.Delimiter);
    }

    [Fact]
    public void CsvTextImportEnforcesMaxInputBytesBeforeCreatingOrMutatingSheets() {
        var options = new ExcelCsvImportOptions {
            SheetName = "Rejected",
            ReaderOptions = new CsvReaderOptions { MaxInputBytes = 8 }
        };
        using var stream = new MemoryStream();
        using var document = ExcelDocument.Create(stream);
        ExcelSheet existing = document.AddSheet("Existing");

        Assert.Throws<InvalidDataException>(() =>
            document.ImportCsvText("Name\r\nAlpha", options));
        Assert.Throws<InvalidDataException>(() =>
            existing.ImportCsvText("Name\r\nAlpha", options));

        Assert.DoesNotContain(document.Sheets, sheet => sheet.Name == "Rejected");
        Assert.Equal("A1:A1", existing.UsedRange.Address);
    }

    [Fact]
    public void SaveAsExcelPassesCancellationIntoWorkbookSerialization() {
        CsvDocument csv = CsvDocument.Parse("Name\r\nAlpha");
        using var cancellation = new CancellationTokenSource();
        using var destination = new CancelOnAsyncWriteStream(cancellation);

        Assert.ThrowsAny<OperationCanceledException>(() => csv.SaveAsExcel(
            destination,
            new CsvExcelExportOptions { SaveOptions = new ExcelSaveOptions { DisableFastPackageWriter = true } },
            cancellationToken: cancellation.Token));
        Assert.Equal(0, destination.Length);
    }

    [Fact]
    public void WorksheetCsvExportHonorsReadOptionsCancellationBeforeWriting() {
        using var stream = new MemoryStream();
        using var document = ExcelDocument.Create(stream);
        ExcelSheet sheet = document.AddSheet("Data");
        sheet.Cell(1, 1).SetValue("Name");
        sheet.Cell(2, 1).SetValue("Alpha");
        using var cancellation = new CancellationTokenSource();
        cancellation.Cancel();
        var readOptions = new ExcelReaderOptions { CancellationToken = cancellation.Token };
        using var destination = new MemoryStream();

        Assert.ThrowsAny<OperationCanceledException>(() =>
            sheet.SaveAsCsv(destination, new ExcelCsvExportOptions { ReadOptions = readOptions }));
        Assert.Equal(0, destination.Length);
    }

    [Fact]
    public void WorksheetCsvRoundTripPreservesQuotedAndMultilineFields() {
        const string csv = "Name,Note,Amount\r\nAlpha,\"Hello, \"\"world\"\"\",10.5\r\nBeta,\"Line\r\nbreak\",20\r\n";
        using var stream = new MemoryStream();
        using var document = ExcelDocument.Create(stream);
        ExcelSheet sheet = document.AddSheet("Data");

        ExcelCsvImportResult imported = sheet.ImportCsvText(csv, new ExcelCsvImportOptions {
            ReaderOptions = new CsvReaderOptions { InferSchema = false }
        });
        using DataTable table = sheet.ToDataTable("A1:C3");
        string exported = sheet.ToCsv(new ExcelCsvExportOptions { Range = "A1:C3" }).ToString();

        Assert.Equal("A1:C3", imported.Range);
        Assert.Equal("Hello, \"world\"", table.Rows[0]["Note"]);
        Assert.Equal("Line\r\nbreak", table.Rows[1]["Note"]);
        Assert.Contains("\"Hello, \"\"world\"\"\"", exported);
        Assert.Contains("\"Line\r\nbreak\"", exported);
    }

    [Fact]
    public void WorksheetCsvExportPreservesFirstDataRowWhenThereIsNoHeader() {
        using var stream = new MemoryStream();
        using var document = ExcelDocument.Create(stream);
        ExcelSheet sheet = document.AddSheet("Data");
        sheet.Cell(1, 1).SetValue("Alpha");
        sheet.Cell(1, 2).SetValue(10);
        sheet.Cell(2, 1).SetValue("Beta");
        sheet.Cell(2, 2).SetValue(20);
        var saveOptions = new CsvSaveOptions { IncludeHeader = true, NewLine = "\n" };

        string exported = sheet.ToCsv(new ExcelCsvExportOptions { Range = "A1:B2", HeadersInFirstRow = false }).ToString(new CsvSaveOptions { IncludeHeader = false, NewLine = saveOptions.NewLine });
        CsvDocument csv = sheet.ToCsv(new ExcelCsvExportOptions { Range = "A1:B2", HeadersInFirstRow = false });
        CsvRow[] rows = csv.AsEnumerable().ToArray();

        Assert.Equal("Alpha,10\nBeta,20\n", exported);
        Assert.True(saveOptions.IncludeHeader);
        Assert.Equal(2, rows.Length);
        Assert.Equal("Alpha", rows[0][0]);
        Assert.Equal("Beta", rows[1][0]);
    }

    [Fact]
    public void WorksheetCsvImportUsesCanonicalMismatchPolicy() {
        const string csv = "Name,Value\r\nAlpha,1,Extra\r\nBeta,\r\n";
        using var stream = new MemoryStream();
        using var document = ExcelDocument.Create(stream);
        ExcelSheet sheet = document.AddSheet("Data");

        ExcelCsvImportResult imported = sheet.ImportCsvText(csv, new ExcelCsvImportOptions {
            ReaderOptions = new CsvReaderOptions { InferSchema = false }
        });
        using DataTable table = sheet.ToDataTable("A1:B3");

        Assert.Equal("A1:B3", imported.Range);
        Assert.Equal(new[] { "Name", "Value" }, table.Columns.Cast<DataColumn>().Select(column => column.ColumnName));
        Assert.Equal(string.Empty, table.Rows[1]["Value"]);
    }

    [Fact]
    public void DelimitedImportDetectsDelimiterAndCreatesRequestedTable() {
        using var stream = new MemoryStream();
        using var document = ExcelDocument.Create(stream);

        ExcelCsvImportResult result = document.ImportCsvText(
            "Name;Amount\r\nAlpha;10.5\r\nBeta;11.75",
            new ExcelCsvImportOptions {
                ReaderOptions = new CsvReaderOptions {
                    DetectDelimiter = true,
                    Culture = CultureInfo.InvariantCulture
                },
                SheetName = "Import",
                TableName = "ImportData"
            });

        Assert.Equal("Import", result.SheetName);
        Assert.Equal("ImportData", result.TableName);
        Assert.Equal("A1:B3", result.Range);
        Assert.Equal(';', result.Delimiter);
        document.Save();
        stream.Position = 0;
        using ExcelDocument reloaded = ExcelDocument.Load(
            stream,
            new ExcelLoadOptions { AccessMode = OfficeIMO.DocumentAccessMode.ReadOnly });
        Assert.Equal("ImportData", reloaded.GetTables().Single().Name);
    }

    [Fact]
    public void CsvImportReportsSanitizedAndUniqueTableNames() {
        using var stream = new MemoryStream();
        using var document = ExcelDocument.Create(stream);

        ExcelCsvImportResult sanitized = document.ImportCsvText(
            "Name,Amount\r\nAlpha,10.5",
            new ExcelCsvImportOptions {
                SheetName = "Sales Data"
            });
        ExcelCsvImportResult unique = document.ImportCsvText(
            "Name,Amount\r\nBeta,11.5",
            new ExcelCsvImportOptions {
                SheetName = "Second",
                TableName = "Sales Data"
            });

        Assert.Equal("Sales_Data", sanitized.TableName);
        Assert.Equal("Sales_Data2", unique.TableName);
        Assert.Equal(sanitized.TableName, document.GetTables().Single(table => table.SheetName == "Sales Data").Name);
        Assert.Equal(unique.TableName, document.GetTables().Single(table => table.SheetName == "Second").Name);
    }

    [Fact]
    public void CsvImportDefaultsTableNameToTheEffectiveWorksheetName() {
        using var stream = new MemoryStream();
        using (var document = ExcelDocument.Create(stream)) {
            document.ImportCsvText(
                "Name,Amount\r\nAlpha,10.5",
                new ExcelCsvImportOptions { SheetName = "Orders" });
            document.Save();
        }

        stream.Position = 0;
        using ExcelDocument reloaded = ExcelDocument.Load(
            stream,
            new ExcelLoadOptions { AccessMode = OfficeIMO.DocumentAccessMode.ReadOnly });
        ExcelTableInfo table = Assert.Single(reloaded.GetTables());
        Assert.Equal("Orders", table.Name);
        Assert.Equal("Orders", table.SheetName);
    }

    [Fact]
    public void DelimitedImportSkipsLogicalRecordsBeforeDetection() {
        using var stream = new MemoryStream();
        using var document = ExcelDocument.Create(stream);

        ExcelCsvImportResult result = document.ImportCsvText(
            "\"generated\r\nstill,has,commas\"\r\nName;Amount\r\nAlpha;10.5",
            new ExcelCsvImportOptions {
                SheetName = "Import",
                ReaderOptions = new CsvReaderOptions {
                    DetectDelimiter = true,
                    SkipInitialRecords = 1
                }
            });

        Assert.Equal("A1:B2", result.Range);
        Assert.Equal(';', result.Delimiter);
        Assert.True(document["Import"].TryGetCellText(2, 2, out string? amount));
        Assert.Equal("10.5", amount);
    }

    [Fact]
    public void CsvFileImportUsesCanonicalHeaderWidth() {
        string path = Path.Combine(Path.GetTempPath(), "OfficeIMO.Reader.Csv." + Guid.NewGuid().ToString("N") + ".csv");
        try {
            File.WriteAllText(path, "Name\r\nAlpha,10.5\r\nBeta,11.75");
            using var stream = new MemoryStream();
            using var document = ExcelDocument.Create(stream);

            ExcelCsvImportResult result = document.ImportCsvFile(path, new ExcelCsvImportOptions {
                SheetName = "Import"
            });

            Assert.Equal("A1:A3", result.Range);
            Assert.True(document["Import"].TryGetCellText(2, 1, out string? name));
            Assert.Equal("Alpha", name);
        } finally {
            if (File.Exists(path)) File.Delete(path);
        }
    }

    [Fact]
    public void DelimitedCompressedFileImportDetectsDelimiterAndReadsRows() {
        string path = Path.Combine(Path.GetTempPath(), "OfficeIMO.Reader.Csv." + Guid.NewGuid().ToString("N") + ".csv.gz");
        try {
            CsvDocument.Create()
                .WithHeader("Name", "Amount")
                .AddRow("Alpha", 10.5m)
                .AddRow("Beta", 11.75m)
                .Save(path, new CsvSaveOptions {
                    Delimiter = ';',
                    CompressionType = CsvCompressionType.GZip
                });

            using var stream = new MemoryStream();
            using var document = ExcelDocument.Create(stream);

            ExcelCsvImportResult result = document.ImportCsvFile(path, new ExcelCsvImportOptions {
                SheetName = "Import",
                ReaderOptions = new CsvReaderOptions {
                    DetectDelimiter = true,
                    Culture = CultureInfo.InvariantCulture
                }
            });

            Assert.Equal("A1:B3", result.Range);
            Assert.Equal(';', result.Delimiter);
            Assert.True(document["Import"].TryGetCellText(2, 2, out string? amount));
            Assert.Equal("10.5", amount);
        } finally {
            if (File.Exists(path)) File.Delete(path);
        }
    }

    private sealed class CancelOnAsyncWriteStream : MemoryStream {
        private readonly CancellationTokenSource _cancellation;

        internal CancelOnAsyncWriteStream(CancellationTokenSource cancellation) {
            _cancellation = cancellation;
        }

        public override Task WriteAsync(
            byte[] buffer,
            int offset,
            int count,
            CancellationToken cancellationToken) {
            _cancellation.Cancel();
            return base.WriteAsync(buffer, offset, count, cancellationToken);
        }
    }
}
