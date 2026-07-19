using System.Data;
using System.Globalization;
using OfficeIMO.CSV;
using OfficeIMO.Excel;
using OfficeIMO.Reader.Csv;
using OfficeIMO.Reader.Excel;
using Xunit;

namespace OfficeIMO.Reader.Tests;

public class ExcelCsvExtensionsTests {
    [Fact]
    public void WorksheetCsvRoundTripPreservesQuotedAndMultilineFields() {
        const string csv = "Name,Note,Amount\r\nAlpha,\"Hello, \"\"world\"\"\",10.5\r\nBeta,\"Line\r\nbreak\",20\r\n";
        using var stream = new MemoryStream();
        using var document = ExcelDocument.Create(stream);
        ExcelSheet sheet = document.AddWorksheet("Data");

        string range = sheet.FromCsv(csv);
        using DataTable table = sheet.ToDataTable("A1:C3");
        string exported = sheet.ToCsv("A1:C3");

        Assert.Equal("A1:C3", range);
        Assert.Equal("Hello, \"world\"", table.Rows[0]["Note"]);
        Assert.Equal("Line\r\nbreak", table.Rows[1]["Note"]);
        Assert.Contains("\"Hello, \"\"world\"\"\"", exported);
        Assert.Contains("\"Line\r\nbreak\"", exported);
    }

    [Fact]
    public void WorksheetCsvImportPreservesWideAndSparseRows() {
        const string csv = "Name,Value\r\nAlpha,1,Extra\r\nBeta,,\r\n";
        using var stream = new MemoryStream();
        using var document = ExcelDocument.Create(stream);
        ExcelSheet sheet = document.AddWorksheet("Data");

        string range = sheet.FromCsv(csv);
        using DataTable table = sheet.ToDataTable("A1:C3");

        Assert.Equal("A1:C3", range);
        Assert.Equal(new[] { "Name", "Value", "Column3" }, table.Columns.Cast<DataColumn>().Select(column => column.ColumnName));
        Assert.Equal("Extra", table.Rows[0]["Column3"]);
        Assert.Equal(string.Empty, table.Rows[1]["Value"]);
        Assert.Equal(string.Empty, table.Rows[1]["Column3"]);
    }

    [Fact]
    public void DelimitedImportDetectsDelimiterAndCreatesRequestedTable() {
        using var stream = new MemoryStream();
        using var document = ExcelDocument.Create(stream);

        ExcelDelimitedImportResult result = document.ImportDelimitedText(
            "Name;Amount\r\nAlpha;10.5\r\nBeta;11.75",
            new ExcelDelimitedImportOptions {
                Culture = CultureInfo.InvariantCulture,
                SheetName = "Import",
                TableName = "ImportData"
            });

        Assert.Equal(';', result.Delimiter);
        Assert.Equal("Import", result.SheetName);
        Assert.Equal("ImportData", result.TableName);
        Assert.Equal("A1:B3", result.Range);
        Assert.Equal(2, result.RowCount);
        Assert.Equal(2, result.ColumnCount);
    }

    [Fact]
    public void DelimitedImportSkipsLogicalRecordsBeforeDetection() {
        using var stream = new MemoryStream();
        using var document = ExcelDocument.Create(stream);

        ExcelDelimitedImportResult result = document.ImportDelimitedText(
            "\"generated\r\nstill,has,commas\"\r\nName;Amount\r\nAlpha;10.5",
            new ExcelDelimitedImportOptions {
                SheetName = "Import",
                SkipInitialRecords = 1
            });

        Assert.Equal(';', result.Delimiter);
        Assert.Equal("A1:B2", result.Range);
        Assert.Equal(1, result.RowCount);
        Assert.True(document["Import"].TryGetCellText(2, 2, out string? amount));
        Assert.Equal("10.5", amount);
    }

    [Fact]
    public void DelimitedFileImportPreservesFieldsBeyondHeader() {
        string path = Path.Combine(Path.GetTempPath(), "OfficeIMO.Reader.Csv." + Guid.NewGuid().ToString("N") + ".csv");
        try {
            File.WriteAllText(path, "Name\r\nAlpha,10.5\r\nBeta,11.75");
            using var stream = new MemoryStream();
            using var document = ExcelDocument.Create(stream);

            ExcelDelimitedImportResult result = document.ImportDelimitedFile(path, new ExcelDelimitedImportOptions {
                SheetName = "Import"
            });

            Assert.Equal("A1:B3", result.Range);
            Assert.Equal(2, result.ColumnCount);
            Assert.True(document["Import"].TryGetCellText(1, 2, out string? header));
            Assert.Equal("Column2", header);
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

            ExcelDelimitedImportResult result = document.ImportDelimitedFile(path, new ExcelDelimitedImportOptions {
                Culture = CultureInfo.InvariantCulture,
                SheetName = "Import"
            });

            Assert.Equal(';', result.Delimiter);
            Assert.Equal("A1:B3", result.Range);
            Assert.Equal(2, result.RowCount);
            Assert.True(document["Import"].TryGetCellText(2, 2, out string? amount));
            Assert.Equal("10.5", amount);
        } finally {
            if (File.Exists(path)) File.Delete(path);
        }
    }
}
