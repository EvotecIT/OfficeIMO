using System;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class ExcelResourceBudgetSecurityTests {
    [Fact]
    public void AddSparklines_RejectsOversizedRangeExpansionBeforeAllocation() {
        string path = GetTemporaryWorkbookPath();
        try {
            using var document = ExcelDocument.Create(path);
            ExcelSheet sheet = document.AddWorksheet("Data");

            ArgumentOutOfRangeException exception = Assert.Throws<ArgumentOutOfRangeException>(() =>
                sheet.AddSparklines("A1:B10001", "C1:C10001"));

            Assert.Equal("locationRange", exception.ParamName);
        } finally {
            DeleteIfExists(path);
        }
    }

    [Fact]
    public void ReadRange_RejectsDenseRangesBeyondConfiguredCellBudget() {
        string path = CreateWorkbookWithRows(2);
        try {
            var options = new ExcelReadOptions { MaxRangeCells = 100 };
            using var reader = ExcelDocumentReader.Open(path, options);

            InvalidDataException exception = Assert.Throws<InvalidDataException>(() =>
                reader.GetSheet("Data").ReadRange("A1:Z100"));

            Assert.Contains("2600", exception.Message, StringComparison.Ordinal);
        } finally {
            DeleteIfExists(path);
        }
    }

    [Fact]
    public void ReadRangeAsDataTable_RejectsDenseRangesBeyondConfiguredCellBudget() {
        string path = CreateWorkbookWithRows(2);
        try {
            var options = new ExcelReadOptions { MaxRangeCells = 100 };
            using var reader = ExcelDocumentReader.Open(path, options);

            InvalidDataException exception = Assert.Throws<InvalidDataException>(() =>
                reader.GetSheet("Data").ReadRangeAsDataTable("A1:XFD1048576"));

            Assert.Contains("17179869184", exception.Message, StringComparison.Ordinal);
        } finally {
            DeleteIfExists(path);
        }
    }

    [Fact]
    public void ReadObjects_RejectsDenseRangesBeyondConfiguredCellBudget() {
        string path = CreateWorkbookWithRows(2);
        try {
            var options = new ExcelReadOptions { MaxRangeCells = 100 };
            using var reader = ExcelDocumentReader.Open(path, options);

            InvalidDataException exception = Assert.Throws<InvalidDataException>(() =>
                reader.GetSheet("Data").ReadObjects("A1:XFD1048576"));

            Assert.Contains("17179869184", exception.Message, StringComparison.Ordinal);
        } finally {
            DeleteIfExists(path);
        }
    }

    [Fact]
    public void ReadObjectsOfT_RejectsDenseRangesBeyondConfiguredCellBudget() {
        string path = CreateWorkbookWithRows(2);
        try {
            var options = new ExcelReadOptions { MaxRangeCells = 100 };
            using var reader = ExcelDocumentReader.Open(path, options);

            InvalidDataException exception = Assert.Throws<InvalidDataException>(() =>
                reader.GetSheet("Data").ReadObjects<BudgetRow>("A1:Z100").ToList());

            Assert.Contains("2600", exception.Message, StringComparison.Ordinal);
        } finally {
            DeleteIfExists(path);
        }
    }

    [Fact]
    public void ReadUsedRange_RejectsOversizedTableBackedDimensionBeforeAllocation() {
        string path = GetTemporaryWorkbookPath();
        try {
            using (var document = ExcelDocument.Create(path)) {
                ExcelSheet sheet = document.AddWorksheet("Data");
                sheet.CellValue(1, 1, "A");
                sheet.CellValue(1, 2, "B");
                sheet.CellValue(2, 1, 1);
                sheet.CellValue(2, 2, 2);
                sheet.AddTable("A1:B2", hasHeader: true, name: "BudgetTable", style: OfficeIMO.Excel.TableStyle.TableStyleMedium2);
                document.Save();
            }

            using (SpreadsheetDocument package = SpreadsheetDocument.Open(path, true)) {
                WorksheetPart worksheetPart = package.WorkbookPart!.WorksheetParts.Single();
                worksheetPart.Worksheet.SheetDimension!.Reference = "A1:XFD1048576";
                worksheetPart.Worksheet.Save();
                TableDefinitionPart tablePart = worksheetPart.TableDefinitionParts.Single();
                tablePart.Table.Reference = "A1:XFD1048576";
                tablePart.Table.Save();
            }

            var options = new ExcelReadOptions { MaxRangeCells = 100 };
            using var reader = ExcelDocumentReader.Open(path, options);

            InvalidDataException exception = Assert.Throws<InvalidDataException>(() =>
                reader.GetSheet("Data").ReadUsedRange());

            Assert.Contains("17179869184", exception.Message, StringComparison.Ordinal);
        } finally {
            DeleteIfExists(path);
        }
    }

    [Fact]
    public void ReadObjectsStream_RejectsExcessiveOutOfOrderRowBuffering() {
        string path = CreateWorkbookWithRows(4);
        try {
            using (SpreadsheetDocument package = SpreadsheetDocument.Open(path, true)) {
                WorksheetPart worksheetPart = package.WorkbookPart!.WorksheetParts.First();
                SheetData data = worksheetPart.Worksheet.GetFirstChild<SheetData>()!;
                Row[] rows = data.Elements<Row>().OrderByDescending(row => row.RowIndex!.Value).ToArray();
                data.RemoveAllChildren<Row>();
                data.Append(rows);
                worksheetPart.Worksheet.Save();
            }

            var options = new ExcelReadOptions { MaxPendingTypedRows = 1 };
            using var reader = ExcelDocumentReader.Open(path, options);

            Assert.Throws<InvalidDataException>(() =>
                reader.GetSheet("Data").ReadObjectsStream<BudgetRow>("A1:A4").ToList());
        } finally {
            DeleteIfExists(path);
        }
    }

    [Fact]
    public void TypedHeaderBindings_DoNotCacheOversizedAttackerControlledHeaders() {
        string path = GetTemporaryWorkbookPath();
        try {
            string header = new string('H', 25_000);
            using (var document = ExcelDocument.Create(path)) {
                ExcelSheet sheet = document.AddWorksheet("Data");
                sheet.CellValue(1, 1, header + "1");
                sheet.CellValue(1, 2, header + "2");
                sheet.CellValue(1, 3, header + "3");
                sheet.CellValue(2, 1, "value");
                document.Save();
            }

            int before = GetHeaderBindingCacheCount();
            using (var reader = ExcelDocumentReader.Open(path)) {
                _ = reader.GetSheet("Data").ReadObjectsStream<BudgetRow>("A1:C2").Single();
            }

            Assert.Equal(before, GetHeaderBindingCacheCount());
        } finally {
            DeleteIfExists(path);
        }
    }

    private static int GetHeaderBindingCacheCount() {
        Type cacheDefinition = typeof(ExcelSheetReader).GetNestedType("TypedObjectBindingCache`1", BindingFlags.NonPublic)!;
        Type cache = cacheDefinition.MakeGenericType(typeof(BudgetRow));
        object dictionary = cache.GetField("HeaderBindings", BindingFlags.NonPublic | BindingFlags.Static)!.GetValue(null)!;
        return (int)dictionary.GetType().GetProperty("Count")!.GetValue(dictionary)!;
    }

    private static string CreateWorkbookWithRows(int rowCount) {
        string path = GetTemporaryWorkbookPath();
        using var document = ExcelDocument.Create(path);
        ExcelSheet sheet = document.AddWorksheet("Data");
        sheet.CellValue(1, 1, "Name");
        for (int row = 2; row <= rowCount; row++) {
            sheet.CellValue(row, 1, "Row" + row.ToString(CultureInfo.InvariantCulture));
        }

        document.Save();
        return path;
    }

    private static string GetTemporaryWorkbookPath() =>
        Path.Combine(Path.GetTempPath(), "OfficeIMO-" + Guid.NewGuid().ToString("N", CultureInfo.InvariantCulture) + ".xlsx");

    private static void DeleteIfExists(string path) {
        if (File.Exists(path)) {
            File.Delete(path);
        }
    }

    private sealed class BudgetRow {
        public string? Name { get; set; }
    }
}
