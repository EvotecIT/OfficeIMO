using System.IO;
using System.Linq;
using OfficeIMO.Excel;
using Xunit;
using Xdr = DocumentFormat.OpenXml.Drawing.Spreadsheet;

namespace OfficeIMO.Tests;

public sealed class ExcelPivotPhysicalGeometryTests {
    [Theory]
    [InlineData(false, false)]
    [InlineData(true, false)]
    [InlineData(false, true)]
    [InlineData(true, true)]
    public void SlicerAndTimelineRetainPhysicalExtentsAfterReload(bool timeline, bool defaults) {
        using ExcelDocument document = ExcelDocument.Create();
        ExcelSheet source = document.AddSheet("Sales");
        source.CellValue(1, 1, "Region");
        source.CellValue(1, 2, "OrderDate");
        source.CellValue(1, 3, "Sales");
        source.CellValue(2, 1, "East");
        source.CellValue(2, 2, new System.DateTime(2026, 1, 2));
        source.CellValue(2, 3, 10D);
        source.CellValue(3, 1, "West");
        source.CellValue(3, 2, new System.DateTime(2026, 2, 3));
        source.CellValue(3, 3, 20D);
        source.AddPivotTable("A1:C3", "E2", options: new ExcelPivotTableOptions {
            Name = "SalesPivot", RowFields = new[] { "Region" },
            DataFields = new[] { new ExcelPivotDataField("Sales", ExcelPivotDataFunction.Sum) }
        });
        document.AddSheet("Dashboard");
        OfficeIMO.OfficeLength width = defaults
            ? OfficeIMO.OfficeLength.Pixels(timeline ? 480 : 180)
            : OfficeIMO.OfficeLength.Emus(1234567);
        OfficeIMO.OfficeLength height = defaults
            ? OfficeIMO.OfficeLength.Pixels(timeline ? 120 : 240)
            : OfficeIMO.OfficeLength.Centimeters(2.345);
        if (timeline) {
            ExcelTimelineViewOptions options = new();
            if (!defaults) { options.Width = width; options.Height = height; }
            document.AddPivotTimeline("SalesPivot", "OrderDate", "Dashboard", options);
        } else {
            ExcelSlicerViewOptions options = new();
            if (!defaults) { options.Width = width; options.Height = height; }
            document.AddPivotSlicer("SalesPivot", "Region", "Dashboard", options);
        }
        using var stream = new MemoryStream(document.ToBytes());
        using ExcelDocument reopened = ExcelDocument.Load(stream);
        var drawing = reopened["Dashboard"].WorksheetPart.DrawingsPart!.WorksheetDrawing!;
        Xdr.Extent extent = Assert.Single(drawing.Descendants<Xdr.Extent>());
        Assert.Equal(width.ToEmus(), extent.Cx!.Value);
        Assert.Equal(height.ToEmus(), extent.Cy!.Value);
        var transform = Assert.Single(drawing.Descendants<DocumentFormat.OpenXml.Drawing.Extents>());
        Assert.Equal(width.ToEmus(), transform.Cx!.Value);
        Assert.Equal(height.ToEmus(), transform.Cy!.Value);
        Assert.True(reopened["Sales"].TryGetCellText(2, 1, out string? region));
        Assert.Equal("East", region);
        Assert.Empty(reopened.ValidateOpenXml());
    }
}
