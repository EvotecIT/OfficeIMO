using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Drawing;
using OfficeIMO.Excel;
using OfficeIMO.Excel.Pdf;
using System.Globalization;
using System.Reflection;
using System.Text;
using System.Text.RegularExpressions;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using Xunit;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Tests;

public partial class Excel {

    [Fact]
    public void SaveAsPdf_ExcelWorkbook_Applies_FirstParty_PageSetup_Options() {
        string workbookPath = Path.Combine(_directoryWithFiles, "ExcelPdfPageSetup.xlsx");

        byte[] bytes;
        using (ExcelDocument document = ExcelDocument.Create(workbookPath, "PageSetup")) {
            ExcelSheet sheet = document.Sheets[0];
            sheet.Cell(1, 1, "Name");
            sheet.Cell(1, 2, "Value");
            sheet.Cell(2, 1, "PageWidth");
            sheet.Cell(2, 2, "Custom");
            document.Save();

            bytes = document.ToPdf(new ExcelPdfSaveOptions {
                PageSize = new PdfCore.PageSize(360, 240),
                Margins = PdfCore.PageMargins.Uniform(24),
                HeaderRowCount = 1
            });
        }

        PdfCore.PdfDocumentInfo info = PdfCore.PdfInspector.Inspect(bytes);
        Assert.Single(info.Pages);
        Assert.Equal(360, info.Pages[0].Width, 1);
        Assert.Equal(240, info.Pages[0].Height, 1);
    }

    [Fact]
    public void SaveAsPdf_ExcelWorkbook_Uses_Worksheet_Print_Area() {
        string workbookPath = Path.Combine(_directoryWithFiles, "ExcelPdfPrintArea.xlsx");

        byte[] bytes;
        using (ExcelDocument document = ExcelDocument.Create(workbookPath, "Report")) {
            ExcelSheet sheet = document.Sheets[0];
            sheet.Cell(1, 1, "OutsideTop");
            sheet.Cell(2, 2, "InsideHeader");
            sheet.Cell(3, 2, "InsideValue");
            sheet.Cell(4, 4, "OutsideRight");
            document.SetPrintArea(sheet, "B2:C3");
            document.Save();

            bytes = document.ToPdf(new ExcelPdfSaveOptions {
                IncludeSheetHeadings = false
            });
        }

        using PdfPigDocument pdf = PdfPigDocument.Open(new MemoryStream(bytes));
        string text = pdf.GetPage(1).Text;
        Assert.Contains("InsideHeader", text);
        Assert.Contains("InsideValue", text);
        Assert.DoesNotContain("OutsideTop", text);
        Assert.DoesNotContain("OutsideRight", text);
    }

    [Fact]
    public void SaveAsPdf_ExcelWorkbook_Uses_Single_Cell_Worksheet_Print_Area() {
        string workbookPath = Path.Combine(_directoryWithFiles, "ExcelPdfSingleCellPrintArea.xlsx");

        byte[] bytes;
        using (ExcelDocument document = ExcelDocument.Create(workbookPath, "Report")) {
            ExcelSheet sheet = document.Sheets[0];
            sheet.Cell(1, 1, "OnlyCell");
            sheet.Cell(2, 1, "OutsideCell");
            document.SetPrintArea(sheet, "A1");
            document.Save();

            bytes = document.ToPdf(new ExcelPdfSaveOptions {
                IncludeSheetHeadings = false
            });
        }

        using PdfPigDocument pdf = PdfPigDocument.Open(new MemoryStream(bytes));
        string text = pdf.GetPage(1).Text;
        Assert.Contains("OnlyCell", text);
        Assert.DoesNotContain("OutsideCell", text);
    }

    [Fact]
    public void SaveAsPdf_ExcelWorkbook_Warns_And_Falls_Back_For_MultiArea_Print_Area() {
        string workbookPath = Path.Combine(_directoryWithFiles, "ExcelPdfMultiAreaPrintArea.xlsx");
        var options = new ExcelPdfSaveOptions {
            IncludeSheetHeadings = false,
            UseWorksheetPrintAreas = true
        };

        byte[] bytes;
        PdfCore.PdfDocumentConversionResult result;
        using (ExcelDocument document = ExcelDocument.Create(workbookPath, "Report")) {
            ExcelSheet sheet = document.Sheets[0];
            sheet.Cell(1, 1, "UsedRangeTop");
            sheet.Cell(2, 2, "AreaOne");
            sheet.Cell(2, 4, "AreaTwo");
            sheet.Cell(5, 5, "UsedRangeBottom");
            document.Save();
        }

        using (SpreadsheetDocument package = SpreadsheetDocument.Open(workbookPath, true)) {
            WorkbookPart workbookPart = package.WorkbookPart ?? throw new InvalidOperationException("Workbook part was not available.");
            Workbook workbook = workbookPart.Workbook ?? throw new InvalidOperationException("Workbook root was not available.");
            workbook.DefinedNames ??= new DefinedNames();
            workbook.DefinedNames.Append(new DefinedName {
                Name = "_xlnm.Print_Area",
                LocalSheetId = 0U,
                Text = "'Report'!$B$2:$B$2,'Report'!$D$2:$D$2"
            });
            workbook.Save();
        }

        using (ExcelDocument document = ExcelDocument.Load(workbookPath)) {
            result = document.ToPdfDocumentResult(options);
            bytes = result.ToBytes();
        }

        using PdfPigDocument pdf = PdfPigDocument.Open(new MemoryStream(bytes));
        string text = pdf.GetPage(1).Text;
        Assert.Contains("UsedRangeTop", text);
        Assert.Contains("AreaOne", text);
        Assert.Contains("AreaTwo", text);
        Assert.Contains("UsedRangeBottom", text);
        Assert.Contains(result.Warnings, warning => warning.Source == "Report" && warning.Code == "WorksheetPrintArea");
    }

    [Fact]
    public void SaveAsPdf_ExcelWorkbook_Filters_Images_And_Charts_Outside_Print_Area() {
        string workbookPath = Path.Combine(_directoryWithFiles, "ExcelPdfPrintAreaMedia.xlsx");
        byte[] imageBytes = CreateMinimalRgbPng();

        byte[] bytes;
        using (ExcelDocument document = ExcelDocument.Create(workbookPath, "Report")) {
            ExcelSheet sheet = document.Sheets[0];
            sheet.Cell(2, 2, "Category");
            sheet.Cell(2, 3, "Value");
            sheet.Cell(3, 2, "Inside");
            sheet.Cell(3, 3, 10);
            sheet.Cell(10, 1, "OutsideData");
            sheet.AddImage(3, 2, imageBytes, "image/png", widthPixels: 12, heightPixels: 12, name: "Inside image");
            sheet.AddImage(10, 1, imageBytes, "image/png", widthPixels: 12, heightPixels: 12, name: "Outside image");
            sheet.AddChartFromRange("B2:C3", row: 3, column: 2, widthPixels: 220, heightPixels: 120, type: ExcelChartType.ColumnClustered, title: "Inside Chart");
            sheet.AddChartFromRange("B2:C3", row: 10, column: 1, widthPixels: 220, heightPixels: 120, type: ExcelChartType.ColumnClustered, title: "Outside Chart");
            document.SetPrintArea(sheet, "B2:C3");
            document.Save();

            bytes = document.ToPdf(new ExcelPdfSaveOptions {
                IncludeSheetHeadings = false,
                UseWorksheetPrintAreas = true,
                PageSize = new PdfCore.PageSize(420, 320),
                Margins = PdfCore.PageMargins.Uniform(24)
            });
        }

        using PdfPigDocument pdf = PdfPigDocument.Open(new MemoryStream(bytes));
        string text = pdf.GetPage(1).Text;
        Assert.Contains("Inside Chart", text);
        Assert.DoesNotContain("Outside Chart", text);
        Assert.DoesNotContain("OutsideData", text);
        Assert.Single(PdfCore.PdfImageExtractor.ExtractImages(bytes));

        OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(
            PdfCore.PdfPageImageRenderer.RenderPage(bytes),
            new OfficeDrawingRasterRenderOptions {
                Background = OfficeColor.White
            });
        int chartInkBelowCellRange = 0;
        for (int y = 60; y < 145; y++) {
            for (int x = 24; x < 210; x++) {
                if (raster.GetPixel(x, y) != OfficeColor.White) {
                    chartInkBelowCellRange++;
                }
            }
        }

        Assert.True(
            chartInkBelowCellRange > 100,
            "Expected the complete chart anchored inside the print area to remain visible below the final exported cell row.");
    }

    [Fact]
    public void SaveAsPdf_ExcelWorkbook_Uses_Worksheet_Orientation_And_Margins() {
        string workbookPath = Path.Combine(_directoryWithFiles, "ExcelPdfWorksheetPageSetup.xlsx");

        byte[] bytes;
        using (ExcelDocument document = ExcelDocument.Create(workbookPath, "PageSetup")) {
            ExcelSheet sheet = document.Sheets[0];
            sheet.Cell(1, 1, "Name");
            sheet.Cell(2, 1, "WorksheetPageSetup");
            sheet.SetOrientation(ExcelPageOrientation.Landscape);
            sheet.SetMargins(left: 0.25, right: 0.25, top: 0.5, bottom: 0.5);
            document.Save();

            bytes = document.ToPdf(new ExcelPdfSaveOptions {
                IncludeSheetHeadings = false
            });
        }

        PdfCore.PdfDocumentInfo info = PdfCore.PdfInspector.Inspect(bytes);
        PdfCore.PdfPageInfo page = Assert.Single(info.Pages);
        Assert.True(page.Width > page.Height, $"Expected worksheet landscape orientation. Width: {page.Width}, height: {page.Height}.");

        using PdfPigDocument pdf = PdfPigDocument.Open(new MemoryStream(bytes));
        double firstLetterX = pdf.GetPage(1).Letters.First(letter => letter.Value == "N").StartBaseLine.X;
        Assert.InRange(firstLetterX, 17D, 36D);
    }

    [Fact]
    public void SaveAsPdf_ExcelWorkbook_Uses_Worksheet_Paper_Size_When_PageSize_Not_Explicit() {
        string workbookPath = Path.Combine(_directoryWithFiles, "ExcelPdfWorksheetPaperSize.xlsx");

        byte[] bytes;
        using (ExcelDocument document = ExcelDocument.Create(workbookPath, "PaperSize")) {
            ExcelSheet sheet = document.Sheets[0];
            sheet.Cell(1, 1, "Name");
            sheet.Cell(2, 1, "WorksheetPaperSize");
            sheet.SetPageSetup(paperSize: ExcelPaperSize.A4);
            document.Save();

            bytes = document.ToPdf(new ExcelPdfSaveOptions {
                IncludeSheetHeadings = false
            });
        }

        PdfCore.PdfDocumentInfo info = PdfCore.PdfInspector.Inspect(bytes);
        PdfCore.PdfPageInfo page = Assert.Single(info.Pages);
        Assert.Equal(595D, page.Width, 1D);
        Assert.Equal(842D, page.Height, 1D);
    }

    [Fact]
    public void SaveAsPdf_ExcelWorkbook_Preserves_PdfOptions_PageSize_When_Excel_PageSize_Not_Explicit() {
        string workbookPath = Path.Combine(_directoryWithFiles, "ExcelPdfPreservePdfOptionsPageSize.xlsx");

        byte[] bytes;
        using (ExcelDocument document = ExcelDocument.Create(workbookPath, "PaperSize")) {
            ExcelSheet sheet = document.Sheets[0];
            sheet.Cell(1, 1, "Name");
            sheet.Cell(2, 1, "PdfOptionsPageSize");
            sheet.SetPageSetup(paperSize: ExcelPaperSize.A4);
            document.Save();

            bytes = document.ToPdf(new ExcelPdfSaveOptions {
                IncludeSheetHeadings = false,
                PdfOptions = new PdfCore.PdfOptions {
                    PageSize = new PdfCore.PageSize(300, 220)
                }
            });
        }

        PdfCore.PdfDocumentInfo info = PdfCore.PdfInspector.Inspect(bytes);
        PdfCore.PdfPageInfo page = Assert.Single(info.Pages);
        Assert.Equal(300D, page.Width, 1D);
        Assert.Equal(220D, page.Height, 1D);
    }

    [Fact]
    public void SaveAsPdf_ExcelWorkbook_Uses_Default_Pdf_PageSize_When_Worksheet_PageSetup_Disabled() {
        string workbookPath = Path.Combine(_directoryWithFiles, "ExcelPdfDisableWorksheetPageSetup.xlsx");

        byte[] bytes;
        using (ExcelDocument document = ExcelDocument.Create(workbookPath, "PaperSize")) {
            ExcelSheet sheet = document.Sheets[0];
            sheet.Cell(1, 1, "Name");
            sheet.Cell(2, 1, "DefaultPdfPageSize");
            sheet.SetPageSetup(paperSize: ExcelPaperSize.A4);
            document.Save();

            bytes = document.ToPdf(new ExcelPdfSaveOptions {
                IncludeSheetHeadings = false,
                UseWorksheetPageSetup = false
            });
        }

        PdfCore.PdfDocumentInfo info = PdfCore.PdfInspector.Inspect(bytes);
        PdfCore.PdfPageInfo page = Assert.Single(info.Pages);
        Assert.Equal(612D, page.Width, 1D);
        Assert.Equal(792D, page.Height, 1D);
    }

    [Fact]
    public void ToPdfDocument_ExcelWorkbook_Maps_Worksheet_FitToHeight_To_Table_Scaling() {
        string workbookPath = Path.Combine(_directoryWithFiles, "ExcelPdfWorksheetFitToHeight.xlsx");

        PdfCore.PdfDocument pdfDocument;
        using (ExcelDocument document = ExcelDocument.Create(workbookPath, "PageSetup")) {
            ExcelSheet sheet = document.Sheets[0];
            sheet.Cell(1, 1, "Name");
            sheet.Cell(1, 2, "Value");
            for (int row = 2; row <= 10; row++) {
                sheet.Cell(row, 1, "Row " + row.ToString(CultureInfo.InvariantCulture));
                sheet.Cell(row, 2, row);
            }

            for (int row = 1; row <= 10; row++) {
                sheet.SetRowHeight(row, 36D);
            }

            sheet.SetPageSetup(fitToWidth: 1U, fitToHeight: 1U);
            document.Save();

            pdfDocument = document.ToPdfDocument(new ExcelPdfSaveOptions {
                IncludeSheetHeadings = false,
                HeaderRowCount = 1,
                WorksheetLayout = ExcelPdfWorksheetLayoutMode.FlowTable,
                PageSize = new PdfCore.PageSize(220, 144),
                Margins = PdfCore.PageMargins.Uniform(18)
            });
        }

        PdfCore.PageBlock page = Assert.IsType<PdfCore.PageBlock>(Assert.Single(pdfDocument.Blocks));
        PdfCore.TableBlock table = Assert.Single(page.Blocks.OfType<PdfCore.TableBlock>());
        Assert.NotNull(table.Style);
        PdfCore.PdfTableStyle style = table.Style!;
        Assert.Equal(1, style.HeaderRowCount);
        Assert.NotNull(style.FixedRowHeights);
        Assert.Equal(10, style.FixedRowHeights!.Count);
        Assert.InRange(style.FixedRowHeights.Sum(height => height ?? 0D), 107D, 108.1D);
        Assert.True(style.CellPaddingY < 3D, "Fit-to-height scaling should shrink default vertical padding along with row heights.");
        Assert.NotNull(style.FontSize);
        Assert.NotNull(style.HeaderFontSize);
        Assert.NotNull(style.FooterFontSize);
        Assert.InRange(style.FontSize!.Value, 3.2D, 3.4D);
        Assert.Equal(style.FontSize.Value, style.HeaderFontSize!.Value);
        Assert.Equal(style.FontSize.Value, style.FooterFontSize!.Value);
    }

    [Fact]
    public void SaveAsPdf_ExcelWorkbook_Uses_Print_Title_Rows_As_Repeating_Table_Header() {
        string workbookPath = Path.Combine(_directoryWithFiles, "ExcelPdfPrintTitles.xlsx");

        byte[] bytes;
        using (ExcelDocument document = ExcelDocument.Create(workbookPath, "LongReport")) {
            ExcelSheet sheet = document.Sheets[0];
            sheet.Cell(1, 1, "RegionHeader");
            sheet.Cell(1, 2, "AmountHeader");
            for (int row = 3; row <= 90; row++) {
                sheet.Cell(row, 1, "Region " + row.ToString(CultureInfo.InvariantCulture));
                sheet.Cell(row, 2, row);
            }

            document.SetPrintArea(sheet, "A3:B90");
            document.SetPrintTitles(sheet, firstRow: 1, lastRow: 1, firstCol: null, lastCol: null);
            document.Save();

            bytes = document.ToPdf(new ExcelPdfSaveOptions {
                IncludeSheetHeadings = false,
                HeaderRowCount = 0,
                PageSize = new PdfCore.PageSize(300, 220),
                Margins = PdfCore.PageMargins.Uniform(18)
            });
        }

        using PdfPigDocument pdf = PdfPigDocument.Open(new MemoryStream(bytes));
        Assert.True(pdf.NumberOfPages > 1);
        Assert.Contains("RegionHeader", pdf.GetPage(1).Text);
        Assert.Contains("RegionHeader", pdf.GetPage(2).Text);
        Assert.Contains("Region 3", pdf.GetPage(1).Text);
    }

    [Fact]
    public void SaveAsPdf_ExcelWorkbook_Honors_Manual_Row_Page_Breaks() {
        string workbookPath = Path.Combine(_directoryWithFiles, "ExcelPdfManualRowPageBreaks.xlsx");

        byte[] bytes;
        byte[] disabledBytes;
        using (ExcelDocument document = ExcelDocument.Create(workbookPath, "ManualBreaks")) {
            ExcelSheet sheet = document.Sheets[0];
            sheet.Cell(1, 1, "Name");
            sheet.Cell(1, 2, "Value");
            sheet.Cell(2, 1, "BeforeBreak");
            sheet.Cell(2, 2, "FirstPage");
            sheet.Cell(3, 1, "BreakTail");
            sheet.Cell(3, 2, "StillFirstPage");
            sheet.Cell(4, 1, "AfterBreak");
            sheet.Cell(4, 2, "SecondPage");
            sheet.Cell(5, 1, "SecondTail");
            sheet.Cell(5, 2, "SecondPageTail");
            sheet.AddManualRowPageBreak(3);

            Assert.Equal(new[] { 3 }, sheet.GetManualRowPageBreaks());
            document.Save();

            var options = new ExcelPdfSaveOptions {
                IncludeSheetHeadings = false,
                HeaderRowCount = 1,
                PageSize = new PdfCore.PageSize(420, 420),
                Margins = PdfCore.PageMargins.Uniform(24)
            };
            bytes = document.ToPdf(options);

            options.UseWorksheetPageBreaks = false;
            disabledBytes = document.ToPdf(options);
        }

        using PdfPigDocument pdf = PdfPigDocument.Open(new MemoryStream(bytes));
        Assert.Equal(2, pdf.NumberOfPages);
        string firstPage = pdf.GetPage(1).Text;
        string secondPage = pdf.GetPage(2).Text;
        Assert.Contains("BeforeBreak", firstPage);
        Assert.Contains("BreakTail", firstPage);
        Assert.DoesNotContain("AfterBreak", firstPage);
        Assert.Contains("AfterBreak", secondPage);
        Assert.Contains("SecondTail", secondPage);
        Assert.Contains("Name", secondPage);
        Assert.Contains("Value", secondPage);

        using PdfPigDocument disabledPdf = PdfPigDocument.Open(new MemoryStream(disabledBytes));
        Assert.Equal(1, disabledPdf.NumberOfPages);
        Assert.Contains("AfterBreak", disabledPdf.GetPage(1).Text);
    }

    [Fact]
    public void SaveAsPdf_ExcelWorkbook_Ignores_Manual_Row_Page_Breaks_Before_Print_Area() {
        string workbookPath = Path.Combine(_directoryWithFiles, "ExcelPdfManualRowBreakBeforePrintArea.xlsx");

        byte[] bytes;
        using (ExcelDocument document = ExcelDocument.Create(workbookPath, "ManualBreaks")) {
            ExcelSheet sheet = document.Sheets[0];
            sheet.Cell(1, 1, "TitleOne");
            sheet.Cell(2, 1, "TitleTwo");
            sheet.Cell(10, 1, "BodyHeader");
            sheet.Cell(10, 2, "ValueHeader");
            sheet.Cell(11, 1, "ExportedBody");
            sheet.Cell(11, 2, "BodyValue");
            sheet.AddManualRowPageBreak(5);
            document.SetPrintArea(sheet, "A10:B11");
            document.SetPrintTitles(sheet, firstRow: 1, lastRow: 2, firstCol: null, lastCol: null);
            document.Save();

            bytes = document.ToPdf(new ExcelPdfSaveOptions {
                IncludeSheetHeadings = false,
                HeaderRowCount = 0,
                PageSize = new PdfCore.PageSize(420, 360),
                Margins = PdfCore.PageMargins.Uniform(24)
            });
        }

        using PdfPigDocument pdf = PdfPigDocument.Open(new MemoryStream(bytes));
        Assert.Equal(1, pdf.NumberOfPages);
        string text = pdf.GetPage(1).Text;
        Assert.Contains("TitleOne", text);
        Assert.Contains("ExportedBody", text);
    }

    [Fact]
    public void SaveAsPdf_ExcelWorkbook_Honors_Manual_Column_Page_Breaks() {
        string workbookPath = Path.Combine(_directoryWithFiles, "ExcelPdfManualColumnPageBreaks.xlsx");

        byte[] bytes;
        byte[] disabledBytes;
        using (ExcelDocument document = ExcelDocument.Create(workbookPath, "ManualColumnBreaks")) {
            ExcelSheet sheet = document.Sheets[0];
            sheet.Cell(1, 1, "LeftHeader");
            sheet.Cell(1, 2, "LeftTail");
            sheet.Cell(1, 3, "RightHeader");
            sheet.Cell(1, 4, "RightTail");
            sheet.Cell(2, 1, "LeftValueA");
            sheet.Cell(2, 2, "LeftValueB");
            sheet.Cell(2, 3, "RightValueC");
            sheet.Cell(2, 4, "RightValueD");
            sheet.SetColumnWidth(1, 16);
            sheet.SetColumnWidth(3, 22);
            sheet.AddManualColumnPageBreak(2);

            Assert.Equal(new[] { 2 }, sheet.GetManualColumnPageBreaks());
            document.Save();

            var options = new ExcelPdfSaveOptions {
                IncludeSheetHeadings = false,
                HeaderRowCount = 1,
                PageSize = new PdfCore.PageSize(560, 320),
                Margins = PdfCore.PageMargins.Uniform(24)
            };
            bytes = document.ToPdf(options);

            options.UseWorksheetPageBreaks = false;
            disabledBytes = document.ToPdf(options);
        }

        using PdfPigDocument pdf = PdfPigDocument.Open(new MemoryStream(bytes));
        Assert.Equal(2, pdf.NumberOfPages);
        string firstPage = pdf.GetPage(1).Text;
        string secondPage = pdf.GetPage(2).Text;
        Assert.Contains("LeftValueA", firstPage);
        Assert.Contains("LeftValueB", firstPage);
        Assert.DoesNotContain("RightValueC", firstPage);
        Assert.Contains("RightValueC", secondPage);
        Assert.Contains("RightValueD", secondPage);
        Assert.DoesNotContain("LeftValueA", secondPage);

        using PdfPigDocument disabledPdf = PdfPigDocument.Open(new MemoryStream(disabledBytes));
        Assert.Equal(1, disabledPdf.NumberOfPages);
        Assert.Contains("LeftValueA", disabledPdf.GetPage(1).Text);
        Assert.Contains("RightValueC", disabledPdf.GetPage(1).Text);
    }

    [Fact]
    public void SaveAsPdf_ExcelWorkbook_Defaults_To_DownThenOver_Page_Order() {
        string workbookPath = Path.Combine(_directoryWithFiles, "ExcelPdfDefaultPageOrder.xlsx");

        byte[] bytes;
        using (ExcelDocument document = ExcelDocument.Create(workbookPath, "PageOrder")) {
            ExcelSheet sheet = document.Sheets[0];
            sheet.Cell(1, 1, "TopLeftPage");
            sheet.Cell(3, 1, "BottomLeftPage");
            sheet.Cell(1, 3, "TopRightPage");
            sheet.Cell(3, 3, "BottomRightPage");
            sheet.AddManualRowPageBreak(2);
            sheet.AddManualColumnPageBreak(2);
            document.Save();

            bytes = document.ToPdf(new ExcelPdfSaveOptions {
                IncludeSheetHeadings = false,
                HeaderRowCount = 0,
                PageSize = new PdfCore.PageSize(360, 260),
                Margins = PdfCore.PageMargins.Uniform(24)
            });
        }

        using PdfPigDocument pdf = PdfPigDocument.Open(new MemoryStream(bytes));
        Assert.Equal(4, pdf.NumberOfPages);
        Assert.Contains("TopLeftPage", pdf.GetPage(1).Text);
        Assert.Contains("BottomLeftPage", pdf.GetPage(2).Text);
        Assert.Contains("TopRightPage", pdf.GetPage(3).Text);
        Assert.Contains("BottomRightPage", pdf.GetPage(4).Text);
    }

}
