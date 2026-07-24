using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel;
using OfficeIMO.Excel.Xlsb;
using OfficeIMO.Excel.Xlsb.Write;
using Xunit;
using ExcelTableStyle = OfficeIMO.Excel.TableStyle;

namespace OfficeIMO.Tests;

public partial class Excel {
    [Fact]
    public void Batch19_MalformedStructuredRangeDoesNotResolveMissingEndpoint() {
        using ExcelDocument document = ExcelDocument.Create();
        ExcelSheet sheet = document.AddWorksheet("Malformed");
        sheet.CellValue(1, 1, "A");
        sheet.CellValue(1, 2, "B");
        sheet.CellValue(2, 1, 1D);
        sheet.CellValue(2, 2, 2D);
        sheet.AddTable("A1:B2", hasHeader: true, name: "Sales", style: ExcelTableStyle.TableStyleMedium2);
        sheet.CellFormula(2, 4, "SUM(Sales[[B]:])");
        sheet.CellFormula(2, 5, "SUM(Sales[@[B]:])");

        ExcelFormulaDependencyGraph graph = document.InspectFormulas().DependencyGraph;

        Assert.Empty(Assert.IsType<ExcelFormulaDependencyNode>(graph.FindNode("Malformed", "D2")).Dependencies);
        Assert.Empty(Assert.IsType<ExcelFormulaDependencyNode>(graph.FindNode("Malformed", "E2")).Dependencies);
    }

    [Fact]
    public void Batch19_LegacyXlsSaveCoalescesDuplicateCellCoordinatesLastWins() {
        using ExcelDocument document = ExcelDocument.Create();
        ExcelSheet sheet = document.AddWorksheet("Duplicates");
        SheetData data = sheet.WorksheetPart.Worksheet.GetFirstChild<SheetData>()!;
        data.RemoveAllChildren();
        data.Append(new Row(
            new Cell { CellReference = "A1", DataType = CellValues.String, CellValue = new CellValue("first") },
            new Cell { CellReference = "A1", DataType = CellValues.String, CellValue = new CellValue("last") }) {
            RowIndex = 1U
        });

        byte[] payload = document.ToBytes(ExcelFileFormat.Xls);
        using ExcelDocument loaded = ExcelDocument.Load(new MemoryStream(payload, writable: false));

        Assert.True(loaded["Duplicates"].TryGetCellText(1, 1, out string? value));
        Assert.Equal("last", value);
    }

    [Fact]
    public void Batch19_LegacyXlsSaveRemovesEarlierDuplicateWhenLastCellIsBlank() {
        using ExcelDocument document = ExcelDocument.Create();
        ExcelSheet sheet = document.AddWorksheet("Duplicates");
        SheetData data = sheet.WorksheetPart.Worksheet.GetFirstChild<SheetData>()!;
        data.RemoveAllChildren();
        data.Append(new Row(
            new Cell { CellReference = "A1", DataType = CellValues.String, CellValue = new CellValue("stale") },
            new Cell { CellReference = "A1" }) {
            RowIndex = 1U
        });

        byte[] payload = document.ToBytes(ExcelFileFormat.Xls);
        using ExcelDocument loaded = ExcelDocument.Load(new MemoryStream(payload, writable: false));

        Assert.False(loaded["Duplicates"].TryGetCellText(1, 1, out _));
    }

    [Fact]
    public void Batch19_XlsbRewriteEnforcesActualUnreferencedPartSize() {
        byte[] source = AddZipEntry(
            CreateMinimalXlsbPackage(),
            "xl/media/unreferenced.bin",
            Enumerable.Repeat((byte)0x41, 2_048).ToArray());

        InvalidDataException exception = Assert.Throws<InvalidDataException>(
            () => XlsbNativePackageWriter.RewritePackage(
                source,
                new Dictionary<string, byte[]>(),
                maxPartBytes: 1_024,
                maxPackageBytes: 128 * 1_024));
        Assert.Contains("configured rewrite limit", exception.Message, StringComparison.OrdinalIgnoreCase);
    }
}
