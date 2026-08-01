using System.Data;
using System.Globalization;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using OfficeIMO.CSV;
using OfficeIMO.Excel;
using OfficeIMO.Reader.Csv;
using OfficeIMO.Reader.Excel;
using Xunit;

namespace OfficeIMO.Reader.Tests;

public class ExcelCsvExtensionsTests {
    [Fact]
    public void CsvImportChecksMethodCancellationBeforeCreatingReaders() {
        using var cancellation = new CancellationTokenSource();
        cancellation.Cancel();
        var loadOptions = new CsvLoadOptions();
        var options = new ExcelCsvImportOptions { LoadOptions = loadOptions };
        CsvDocument csv = CsvDocument.Parse("Name\r\nAlpha");
        using var stream = new MemoryStream();
        using var document = ExcelDocument.Create(stream);
        ExcelSheet sheet = document.AddWorksheet("Existing");

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
        Assert.Equal(string.Empty, result.Range);
        Assert.Equal("A1:A1", document["Empty"].GetUsedRangeA1());
    }

    [Fact]
    public void CsvTextImportPreservesUnicodeRegardlessOfFileEncodingOption() {
        var options = new ExcelCsvImportOptions {
            SheetName = "Imported",
            LoadOptions = new CsvLoadOptions { Encoding = Encoding.ASCII },
            ReaderOptions = new CsvDataReaderOptions { InferSchema = false }
        };
        using var stream = new MemoryStream();
        using var document = ExcelDocument.Create(stream);

        document.ImportCsvText("Name\r\nélève", options);
        ExcelSheet existing = document.AddWorksheet("Existing");
        existing.ImportCsvText("Name\r\n東京", options);

        Assert.True(document["Imported"].TryGetCellText(2, 1, out string? imported));
        Assert.Equal("élève", imported);
        Assert.True(existing.TryGetCellText(2, 1, out string? existingValue));
        Assert.Equal("東京", existingValue);
    }

    [Fact]
    public void SaveAsExcelPassesCancellationIntoWorkbookSerialization() {
        CsvDocument csv = CsvDocument.Parse("Name\r\nAlpha");
        using var cancellation = new CancellationTokenSource();
        using var destination = new CancelOnAsyncWriteStream(cancellation);

        Assert.ThrowsAny<OperationCanceledException>(() => csv.SaveAsExcel(
            destination,
            saveOptions: new ExcelSaveOptions { DisableFastPackageWriter = true },
            cancellationToken: cancellation.Token));
        Assert.Equal(0, destination.Length);
    }

    [Fact]
    public void WorksheetCsvExportHonorsReadOptionsCancellationBeforeWriting() {
        using var stream = new MemoryStream();
        using var document = ExcelDocument.Create(stream);
        ExcelSheet sheet = document.AddWorksheet("Data");
        sheet.CellValue(1, 1, "Name");
        sheet.CellValue(2, 1, "Alpha");
        using var cancellation = new CancellationTokenSource();
        cancellation.Cancel();
        var readOptions = new ExcelReadOptions { CancellationToken = cancellation.Token };
        using var destination = new MemoryStream();

        Assert.ThrowsAny<OperationCanceledException>(() =>
            sheet.SaveAsCsv(destination, readOptions: readOptions));
        Assert.Equal(0, destination.Length);
    }

    [Fact]
    public void WorksheetCsvRoundTripPreservesQuotedAndMultilineFields() {
        const string csv = "Name,Note,Amount\r\nAlpha,\"Hello, \"\"world\"\"\",10.5\r\nBeta,\"Line\r\nbreak\",20\r\n";
        using var stream = new MemoryStream();
        using var document = ExcelDocument.Create(stream);
        ExcelSheet sheet = document.AddWorksheet("Data");

        ExcelCsvImportResult imported = sheet.ImportCsvText(csv, new ExcelCsvImportOptions {
            ReaderOptions = new CsvDataReaderOptions { InferSchema = false }
        });
        using DataTable table = sheet.ToDataTable("A1:C3");
        string exported = sheet.ToCsv("A1:C3");

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
        ExcelSheet sheet = document.AddWorksheet("Data");
        sheet.CellValue(1, 1, "Alpha");
        sheet.CellValue(1, 2, 10);
        sheet.CellValue(2, 1, "Beta");
        sheet.CellValue(2, 2, 20);
        var saveOptions = new CsvSaveOptions { IncludeHeader = true, NewLine = "\n" };

        string exported = sheet.ToCsv("A1:B2", headersInFirstRow: false, csvOptions: saveOptions);
        CsvDocument csv = sheet.ToCsvDocument("A1:B2", headersInFirstRow: false);
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
        ExcelSheet sheet = document.AddWorksheet("Data");

        ExcelCsvImportResult imported = sheet.ImportCsvText(csv, new ExcelCsvImportOptions {
            ReaderOptions = new CsvDataReaderOptions { InferSchema = false }
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
                LoadOptions = new CsvLoadOptions {
                    DetectDelimiter = true,
                    Culture = CultureInfo.InvariantCulture
                },
                SheetName = "Import",
                TableName = "ImportData"
            });

        Assert.Equal("Import", result.SheetName);
        Assert.Equal("A1:B3", result.Range);
        document.Save();
        stream.Position = 0;
        using ExcelDocument reloaded = ExcelDocument.Load(
            stream,
            new ExcelLoadOptions { AccessMode = OfficeIMO.Drawing.DocumentAccessMode.ReadOnly });
        Assert.Equal("ImportData", reloaded.GetTables().Single().Name);
    }

    [Fact]
    public void DelimitedImportSkipsLogicalRecordsBeforeDetection() {
        using var stream = new MemoryStream();
        using var document = ExcelDocument.Create(stream);

        ExcelCsvImportResult result = document.ImportCsvText(
            "\"generated\r\nstill,has,commas\"\r\nName;Amount\r\nAlpha;10.5",
            new ExcelCsvImportOptions {
                SheetName = "Import",
                LoadOptions = new CsvLoadOptions {
                    DetectDelimiter = true,
                    SkipInitialRecords = 1
                }
            });

        Assert.Equal("A1:B2", result.Range);
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
            using (TextWriter writer = CsvFile.CreateTextWriter(
                       path,
                       new CsvSaveOptions { CompressionType = CsvCompressionType.GZip })) {
                writer.Write("Name;Amount\r\nAlpha;10.5\r\nBeta;11.75");
            }

            using var stream = new MemoryStream();
            using var document = ExcelDocument.Create(stream);

            ExcelCsvImportResult result = document.ImportCsvFile(path, new ExcelCsvImportOptions {
                SheetName = "Import",
                LoadOptions = new CsvLoadOptions {
                    DetectDelimiter = true,
                    Culture = CultureInfo.InvariantCulture
                }
            });

            Assert.Equal("A1:B3", result.Range);
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
