using System.IO;
using System.Linq;
using OfficeIMO.Excel;
using OfficeIMO.Excel.OpenDocument;
using OfficeIMO.OpenDocument;
using Xunit;

namespace OfficeIMO.OpenDocument.Converters.Tests;

public sealed class SpreadsheetCellLayoutConversionTests {
    [Fact]
    public void ExcelAlignmentAndWrapRoundTripThroughOdsCellStyle() {
        using ExcelDocument source = ExcelDocument.Create();
        ExcelSheet sheet = source.AddWorksheet("Layout");
        sheet.CellAt(1, 1).SetValue("Wrapped text");
        sheet.CellAlign(1, 1, ExcelHorizontalAlignment.Right);
        sheet.CellVerticalAlign(1, 1, ExcelVerticalAlignment.Center);
        sheet.CellWrapText(1, 1);

        OdfConversionResult<OdsDocument> toOds = source.ToOpenDocumentResult();
        Assert.DoesNotContain(toOds.Report.Mappings, mapping => mapping.Feature == "cell-format-details" &&
            mapping.Status == OdfConversionMappingStatus.Unsupported);
        OdsDocument reopened = OdsDocument.Load(new MemoryStream(toOds.Value.ToBytes()));
        Assert.True(reopened.Validate().IsValid);
        OdsCell cell = reopened.GetSheet("Layout")!.Cell(0, 0);
        Assert.Equal("right", cell.TextAlign);
        Assert.Equal("middle", cell.VerticalAlign);
        Assert.Equal("wrap", cell.WrapOption);

        OdfConversionResult<ExcelDocument> toExcel = reopened.ToExcelDocumentResult();
        using ExcelDocument roundTrip = toExcel.Value;
        ExcelCellStyleSnapshot style = roundTrip.CreateInspectionSnapshot().Worksheets.Single()
            .Cells.Single().Style!;
        Assert.Equal("right", style.HorizontalAlignment);
        Assert.Equal("center", style.VerticalAlignment);
        Assert.True(style.WrapText);
        Assert.DoesNotContain(toExcel.Report.Mappings, mapping => mapping.Feature == "cell-layout" &&
            mapping.Status == OdfConversionMappingStatus.Unsupported);
    }

    [Fact]
    public void InheritedOdsCellLayoutProjectsToExcel() {
        OdsDocument source = OdsDocument.Create();
        OdfStyle parent = source.Styles.CreateNamed("LayoutParent", OdfStyleFamily.TableCell);
        parent.TextAlign = "center";
        parent.CellVerticalAlign = "bottom";
        parent.CellWrapOption = "no-wrap";
        OdfStyle child = source.Styles.CreateAutomatic(OdfStyleFamily.TableCell);
        child.ParentStyleName = parent.Name;
        OdsCell cell = source.AddSheet("Layout").Cell(0, 0);
        cell.SetString("Inherited");
        cell.StyleName = child.Name;

        OdsDocument reopened = OdsDocument.Load(new MemoryStream(source.ToBytes()));
        Assert.True(reopened.Validate().IsValid);
        Assert.Equal("center", reopened.GetSheet("Layout")!.Cell(0, 0).TextAlign);
        OdfConversionResult<ExcelDocument> conversion = reopened.ToExcelDocumentResult();
        using ExcelDocument target = conversion.Value;
        ExcelCellStyleSnapshot style = target.CreateInspectionSnapshot().Worksheets.Single()
            .Cells.Single().Style!;
        Assert.Equal("center", style.HorizontalAlignment);
        Assert.Equal("bottom", style.VerticalAlignment);
        Assert.False(style.WrapText);
    }

    [Fact]
    public void UnsupportedAlignmentTokensRemainExplicitLoss() {
        using ExcelDocument excel = ExcelDocument.Create();
        ExcelSheet sheet = excel.AddWorksheet("Layout");
        sheet.CellAt(1, 1).SetValue("Across selection");
        sheet.CellAlign(1, 1, ExcelHorizontalAlignment.CenterContinuous);
        OdfConversionResult<OdsDocument> toOds = excel.ToOpenDocumentResult();
        Assert.Contains(toOds.Report.Mappings, mapping => mapping.Feature == "cell-format-details" &&
            mapping.Status == OdfConversionMappingStatus.Unsupported);

        OdsDocument ods = OdsDocument.Create();
        OdsCell cell = ods.AddSheet("Layout").Cell(0, 0);
        cell.SetString("Direction-sensitive");
        cell.TextAlign = "start";
        OdfConversionResult<ExcelDocument> toExcel = ods.ToExcelDocumentResult();
        using ExcelDocument target = toExcel.Value;
        Assert.Contains(toExcel.Report.Mappings, mapping => mapping.Feature == "cell-layout" &&
            mapping.Status == OdfConversionMappingStatus.Unsupported);
        Assert.Throws<OdfConversionLossException>(() => ods.ToExcelDocumentResult(
            new ExcelOpenDocumentConversionOptions {
                LossPolicy = OdfConversionLossPolicy.ThrowOnSkippedOrUnsupported
            }));
    }
}
