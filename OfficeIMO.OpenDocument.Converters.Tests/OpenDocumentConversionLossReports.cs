using System;
using System.IO;
using System.Linq;
using OfficeIMO.Excel;
using OfficeIMO.Excel.OpenDocument;
using OfficeIMO.OpenDocument;
using OfficeIMO.PowerPoint;
using OfficeIMO.PowerPoint.OpenDocument;
using OfficeIMO.Word.OpenDocument;
using Xunit;

namespace OfficeIMO.OpenDocument.Converters.Tests;

public sealed class OpenDocumentConversionLossReportTests {
    [Fact]
    public void ExcelFormulaConversionDoesNotRewriteQuotedCellLikeText() {
        using ExcelDocument source = ExcelDocument.Create(new MemoryStream(), autoSave: false);
        ExcelSheet sheet = source.AddWorkSheet("Data");
        sheet.CellAt(1, 1).SetValue("B2");
        sheet.CellAt(1, 2).SetFormula("IF(A1=\"B2\",1,0)");

        OdfConversionResult<OdsDocument> conversion = source.ToOpenDocument();
        using OdsDocument target = conversion.Document;

        Assert.Equal("of:=IF([.A1]=\"B2\",1,0)", target.GetSheet("Data")!.Cell(0, 1).Formula);

        target.GetSheet("Data")!.Cell(0, 2).Formula = "of:=IF([.A1]=\"[.B2]\",1,0)";
        OdfConversionResult<ExcelDocument> reverse = target.ToExcelDocument();
        using ExcelDocument roundTrip = reverse.Document;
        ExcelCellSnapshot reverseFormula = roundTrip.CreateInspectionSnapshot().Worksheets.Single().Cells
            .Single(cell => cell.Row == 1 && cell.Column == 3);
        Assert.Equal("IF(A1=\"[.B2]\",1,0)", reverseFormula.Formula);
    }

    [Fact]
    public void UnsupportedPowerPointImageFormatsAreReportedWithoutAbortingConversion() {
        using PowerPointPresentation source = PowerPointPresentation.Create(new MemoryStream(), autoSave: false);
        PowerPointSlide slide = source.AddSlide();
        using var tiff = new MemoryStream(new byte[] { 0x49, 0x49, 0x2A, 0x00, 0x00, 0x00, 0x00, 0x00 });
        slide.AddPicture(tiff, ImagePartType.Tiff);

        OdfConversionResult<OdpPresentation> conversion = source.ToOpenDocument();
        using OdpPresentation target = conversion.Document;

        Assert.Empty(Assert.Single(target.Slides).Shapes.OfType<OdpImage>());
        Assert.Contains(conversion.Report.Mappings, mapping => mapping.Feature == "images" &&
            mapping.Status == OdfConversionMappingStatus.Unsupported && mapping.Count == 1);
    }

    [Fact]
    public void OdtTrackedChangesAreReportedWhenConvertingToWord() {
        using OdtDocument source = OdtDocument.Create();
        source.AddTrackedParagraphInsertion("Inserted", "Author");

        OdfConversionResult<OfficeIMO.Word.WordDocument> conversion = source.ToWordDocument();
        using OfficeIMO.Word.WordDocument target = conversion.Document;

        Assert.True(conversion.Report.HasLoss);
        Assert.Contains(conversion.Report.Mappings, mapping => mapping.Feature == "source-tracked-changes" && mapping.Status == OdfConversionMappingStatus.Unsupported);
    }

    [Fact]
    public void OdpAnimationsAreReportedWhenConvertingToPowerPoint() {
        using OdpPresentation source = OdpPresentation.Create();
        OdpSlide slide = source.AddSlide("Animated");
        OdpRectangle shape = slide.AddRectangle(OdfRect.FromCentimeters(1, 1, 2, 2));
        slide.AddFadeInAnimation(shape, TimeSpan.FromSeconds(1));

        var conversion = source.ToPowerPointPresentation();
        using var target = conversion.Document;

        Assert.True(conversion.Report.HasLoss);
        Assert.Contains(conversion.Report.Mappings, mapping => mapping.Feature == "source-presentation-animations" && mapping.Status == OdfConversionMappingStatus.Unsupported);
    }

    [Fact]
    public void OdsExpansionLimitStillCreatesEveryWorksheetAndReportsTruncation() {
        using OdsDocument source = OdsDocument.Create();
        OdsSheet hidden = source.AddSheet("Hidden");
        hidden.Hidden = true;
        hidden.Cell(0, 0).SetNumber(1);
        hidden.Cell(0, 1).SetNumber(2);
        source.AddSheet("Visible").Cell(0, 0).SetNumber(3);

        OdfConversionResult<ExcelDocument> conversion = source.ToExcelDocument(new ExcelOpenDocumentConversionOptions {
            MaximumExpandedCells = 1
        });
        using ExcelDocument target = conversion.Document;

        Assert.Equal(2, target.CreateInspectionSnapshot().Worksheets.Count);
        Assert.Contains(conversion.Report.Mappings, mapping => mapping.Feature == "worksheets" && mapping.Count == 2);
        Assert.Contains(conversion.Report.Mappings, mapping => mapping.Feature == "expansion-limits" && mapping.Status == OdfConversionMappingStatus.Skipped);
    }

    [Fact]
    public void ExcelToOdsReportsConfiguredCellAndStyleOmissions() {
        using ExcelDocument source = ExcelDocument.Create(new MemoryStream(), autoSave: false);
        ExcelSheet sheet = source.AddWorkSheet("Data");
        sheet.CellAt(1, 1).SetValue("One").SetBold();
        sheet.CellAt(1, 2).SetValue("Two").SetBold();

        OdfConversionResult<OdsDocument> conversion = source.ToOpenDocument(new ExcelOpenDocumentConversionOptions {
            IncludeBasicStyles = false,
            MaximumExpandedCells = 1
        });
        using OdsDocument target = conversion.Document;

        Assert.Contains(conversion.Report.Mappings, mapping => mapping.Feature == "cell-styles" && mapping.Status == OdfConversionMappingStatus.Skipped);
        Assert.Contains(conversion.Report.Mappings, mapping => mapping.Feature == "expansion-limits" && mapping.Status == OdfConversionMappingStatus.Skipped);
    }
}
