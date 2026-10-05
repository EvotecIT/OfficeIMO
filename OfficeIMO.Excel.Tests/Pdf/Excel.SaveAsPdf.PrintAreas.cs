using OfficeIMO.Excel;
using OfficeIMO.Excel.Pdf;
using PdfCore = OfficeIMO.Pdf;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using Xunit;

namespace OfficeIMO.Tests;

public partial class Excel {
    [Theory]
    [InlineData(ExcelPdfWorksheetLayoutMode.WorksheetCanvas)]
    [InlineData(ExcelPdfWorksheetLayoutMode.FlowTable)]
    public void SaveAsPdf_OverlappingPrintAreasLinkToFirstOccurrenceOfTargetCell(ExcelPdfWorksheetLayoutMode layout) {
        string path = Path.Combine(_directoryWithFiles, "OverlappingPrintAreas-" + layout + ".xlsx");
        using ExcelDocument document = ExcelDocument.Create(path, "Report");
        ExcelSheet sheet = document.Sheets[0];
        sheet.Cell(1, 1, "Target");
        sheet.SetInternalLink(2, 1, "A1", display: "Back to target");
        document.SetPrintArea(sheet, "A1:A2,A1:A3");

        byte[] bytes = document.ToPdfBytes(new ExcelToPdfOptions {
            IncludeSheetHeadings = false, HeaderRowCount = 0, WorksheetLayout = layout
        });
        PdfCore.PdfDocumentReadResult logical = PdfCore.PdfDocumentReadResult.Load(bytes);
        PdfCore.PdfNamedDestination destination = Assert.Single(logical.NamedDestinations,
            item => item.Name.EndsWith("-a1", StringComparison.Ordinal));
        Assert.Equal(1, destination.PageNumber);
        Assert.Equal(2, logical.GetLinksByDestinationName(destination.Name).Count);
    }

    [Theory]
    [InlineData(1, 2, false)]
    [InlineData(1, 3, true)]
    [InlineData(3, 3, false)]
    public void FeatureReport_PrintAreaUnionChecksOnlyExportedFormulaCaches(int row, int column, bool canExport) {
        string path = Path.Combine(_directoryWithFiles, "UnionFormula-" + row + "-" + column + ".xlsx");
        using ExcelDocument document = ExcelDocument.Create(path, "Report");
        ExcelSheet sheet = document.Sheets[0];
        sheet.Cell(1, 1, "First");
        sheet.Cell(3, 3, "Second");
        sheet.CellFormula(row, column, "UNIQUE(A1:A1)");
        document.SetPrintArea(sheet, "A1:B2,C3:D4");
        Assert.Equal(canExport, document.InspectFeatures().Can(ExcelPreflightCapability.ExportPdfReport));
    }

    [Theory]
    [InlineData(ExcelPdfWorksheetLayoutMode.WorksheetCanvas)]
    [InlineData(ExcelPdfWorksheetLayoutMode.FlowTable)]
    public void SaveAsPdf_MultiplePrintAreasShareWorksheetHeaderVariantsAndPageCount(ExcelPdfWorksheetLayoutMode layout) {
        string path = Path.Combine(_directoryWithFiles, "PrintAreaHeaderVariants-" + layout + ".xlsx");
        using ExcelDocument document = ExcelDocument.Create(path, "Report");
        ExcelSheet sheet = document.Sheets[0];
        sheet.Cell(1, 1, "FirstArea");
        sheet.Cell(1, 3, "SecondArea");
        sheet.Cell(1, 5, "ThirdArea");
        sheet.SetHeaderFooter(headerCenter: "OddHeader", footerCenter: "Page &P of &N");
        sheet.SetFirstPageHeaderFooter(headerCenter: "FirstHeader", footerCenter: "Page &P of &N");
        sheet.SetEvenPageHeaderFooter(headerCenter: "EvenHeader", footerCenter: "Page &P of &N");
        document.SetPrintArea(sheet, "A1,C1,E1");
        byte[] bytes = document.ToPdfBytes(new ExcelToPdfOptions {
            IncludeSheetHeadings = false, HeaderRowCount = 0, WorksheetLayout = layout
        });
        using PdfPigDocument pdf = PdfPigDocument.Open(bytes);
        Assert.Equal(3, pdf.NumberOfPages);
        Assert.Contains("FirstHeader", pdf.GetPage(1).Text);
        Assert.Contains("EvenHeader", pdf.GetPage(2).Text);
        Assert.Contains("OddHeader", pdf.GetPage(3).Text);
        Assert.Contains("Page 1 of 3", pdf.GetPage(1).Text);
        Assert.Contains("Page 2 of 3", pdf.GetPage(2).Text);
        Assert.Contains("Page 3 of 3", pdf.GetPage(3).Text);
    }

    [Theory]
    [InlineData(ExcelPdfWorksheetLayoutMode.WorksheetCanvas)]
    [InlineData(ExcelPdfWorksheetLayoutMode.FlowTable)]
    public void SaveAsPdf_MultipleQuotedPrintAreasPreserveTitlesAndFilterImages(ExcelPdfWorksheetLayoutMode layout) {
        string path = Path.Combine(_directoryWithFiles, "QuotedPrintAreas-" + layout + ".xlsx");
        byte[] image = CreateMinimalRgbPng();
        using (ExcelDocument document = ExcelDocument.Create(path, "Report, O'Brien")) {
            ExcelSheet sheet = document.Sheets[0];
            sheet.Cell(1, 2, "FirstTitle");
            sheet.Cell(1, 4, "SecondTitle");
            sheet.Cell(2, 2, "FirstArea");
            sheet.Cell(5, 4, "SecondArea");
            sheet.Cell(10, 1, "ExcludedValue");
            sheet.AddImage(2, 2, image, "image/png", widthPixels: 12, heightPixels: 12, name: "First image");
            sheet.AddImage(5, 4, image, "image/png", widthPixels: 12, heightPixels: 12, name: "Second image");
            sheet.AddImage(10, 1, image, "image/png", widthPixels: 12, heightPixels: 12, name: "Excluded image");
            document.SetPrintArea(sheet, "B2:C3,'Report, O''Brien'!$D$5:$E$6", save: false);
            document.SetPrintTitles(sheet, firstRow: 1, lastRow: 1, firstCol: null, lastCol: null, save: false);
            Assert.Equal(2, sheet.GetPrintAreas().Count);
            document.Save();
        }
        using ExcelDocument reopened = ExcelDocument.Load(path);
        Assert.Equal(2, reopened.Sheets[0].GetPrintAreas().Count);
        byte[] bytes = reopened.ToPdfBytes(new ExcelToPdfOptions {
            IncludeSheetHeadings = false, HeaderRowCount = 0, WorksheetLayout = layout
        });
        using PdfPigDocument pdf = PdfPigDocument.Open(bytes);
        Assert.Equal(2, pdf.NumberOfPages);
        string first = pdf.GetPage(1).Text;
        string second = pdf.GetPage(2).Text;
        Assert.Contains("FirstTitle", first);
        Assert.Contains("FirstArea", first);
        Assert.DoesNotContain("SecondArea", first);
        Assert.Contains("SecondTitle", second);
        Assert.Contains("SecondArea", second);
        Assert.DoesNotContain("FirstArea", second);
        Assert.DoesNotContain("ExcludedValue", first + second);
        Assert.Equal(2, PdfCore.PdfImageExtractor.ExtractImages(bytes).Count);
    }

    [Theory]
    [InlineData(ExcelPdfWorksheetLayoutMode.WorksheetCanvas, false)]
    [InlineData(ExcelPdfWorksheetLayoutMode.FlowTable, false)]
    [InlineData(ExcelPdfWorksheetLayoutMode.WorksheetCanvas, true)]
    [InlineData(ExcelPdfWorksheetLayoutMode.FlowTable, true)]
    public void SaveAsPdf_WholeRowAndColumnPrintAreasUseBoundedWorksheetData(ExcelPdfWorksheetLayoutMode layout, bool rows) {
        string path = Path.Combine(_directoryWithFiles, "WholePrintAreas-" + layout + "-" + rows + ".xlsx");
        using ExcelDocument document = ExcelDocument.Create(path, "Report");
        ExcelSheet sheet = document.Sheets[0];
        sheet.Cell(1, 1, "ExcludedTop");
        sheet.Cell(rows ? 2 : 1, rows ? 1 : 2, "FirstArea");
        sheet.Cell(rows ? 4 : 1, rows ? 1 : 4, "SecondArea");
        sheet.Cell(rows ? 3 : 1, rows ? 1 : 3, "ExcludedBetween");
        document.SetPrintArea(sheet, rows ? "2:2,4:4" : "B:B,D:D");
        byte[] bytes = document.ToPdfBytes(new ExcelToPdfOptions {
            IncludeSheetHeadings = false, HeaderRowCount = 0, WorksheetLayout = layout
        });
        using PdfPigDocument pdf = PdfPigDocument.Open(bytes);
        Assert.Equal(2, pdf.NumberOfPages);
        Assert.Contains("FirstArea", pdf.GetPage(1).Text);
        Assert.Contains("SecondArea", pdf.GetPage(2).Text);
        Assert.DoesNotContain("ExcludedTop", pdf.GetPage(1).Text + pdf.GetPage(2).Text);
        Assert.DoesNotContain("ExcludedBetween", pdf.GetPage(1).Text + pdf.GetPage(2).Text);
    }

    [Theory]
    [InlineData("B2,invalid")]
    [InlineData("B2,'Other'!D4")]
    [InlineData("B2,'[External.xlsx]Report'!D4")]
    public void SetPrintArea_InvalidUnionPreservesPreviousSelection(string area) {
        string path = Path.Combine(_directoryWithFiles, "InvalidPrintArea.xlsx");
        using ExcelDocument document = ExcelDocument.Create(path, "Report");
        ExcelSheet sheet = document.Sheets[0];
        document.SetPrintArea(sheet, "A1");
        string? before = sheet.GetPrintArea();
        Assert.ThrowsAny<ArgumentException>(() => document.SetPrintArea(sheet, area));
        Assert.Equal(before, sheet.GetPrintArea());
    }
}
