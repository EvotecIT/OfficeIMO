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
        using (ExcelDocument document = ExcelDocument.Create(workbookPath, new ExcelCreateOptions { FirstSheetName = "PageSetup" })) {
            ExcelSheet sheet = document.Sheets[0];
            sheet.Cell(1, 1, "Name");
            sheet.Cell(1, 2, "Value");
            sheet.Cell(2, 1, "PageWidth");
            sheet.Cell(2, 2, "Custom");
            document.Save();

            bytes = document.ToPdf(new ExcelPdfExportOptions {
                PageSize = new PdfCore.PdfPageSize(OfficeIMO.OfficeLength.Points(360),OfficeIMO.OfficeLength.Points(240)),
                Margins = PdfCore.PdfPageMargins.Uniform(OfficeIMO.OfficeLength.Points(24)),
                HeaderRowCount = 1
            }).ToBytes(new OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true });
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
        using (ExcelDocument document = ExcelDocument.Create(workbookPath, new ExcelCreateOptions { FirstSheetName = "Report" })) {
            ExcelSheet sheet = document.Sheets[0];
            sheet.Cell(1, 1, "OutsideTop");
            sheet.Cell(2, 2, "InsideHeader");
            sheet.Cell(3, 2, "InsideValue");
            sheet.Cell(4, 4, "OutsideRight");
            document.SetPrintArea(sheet, "B2:C3");
            document.Save();

            bytes = document.ToPdf(new ExcelPdfExportOptions {
                IncludeSheetHeadings = false
            }).ToBytes(new OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true });
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
        using (ExcelDocument document = ExcelDocument.Create(workbookPath, new ExcelCreateOptions { FirstSheetName = "Report" })) {
            ExcelSheet sheet = document.Sheets[0];
            sheet.Cell(1, 1, "OnlyCell");
            sheet.Cell(2, 1, "OutsideCell");
            document.SetPrintArea(sheet, "A1");
            document.Save();

            bytes = document.ToPdf(new ExcelPdfExportOptions {
                IncludeSheetHeadings = false
            }).ToBytes(new OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true });
        }

        using PdfPigDocument pdf = PdfPigDocument.Open(new MemoryStream(bytes));
        string text = pdf.GetPage(1).Text;
        Assert.Contains("OnlyCell", text);
        Assert.DoesNotContain("OutsideCell", text);
    }

    [Theory]
    [InlineData(ExcelPdfWorksheetLayoutMode.WorksheetCanvas)]
    [InlineData(ExcelPdfWorksheetLayoutMode.FlowTable)]
    public void SaveAsPdf_ExcelWorkbook_Exports_MultiArea_Print_Areas_On_Separate_Pages(ExcelPdfWorksheetLayoutMode layout) {
        string workbookPath = Path.Combine(_directoryWithFiles, "ExcelPdfMultiAreaPrintArea.xlsx");
        var options = new ExcelPdfExportOptions {
            IncludeSheetHeadings = false,
            UseWorksheetPrintAreas = true,
            WorksheetLayout = layout
        };

        byte[] bytes;
        PdfCore.PdfDocumentConversionResult result;
        using (ExcelDocument document = ExcelDocument.Create(workbookPath, new ExcelCreateOptions { FirstSheetName = "Report" })) {
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
            result = document.ConvertToPdfWithEvidence(options);
            bytes = result.ToBytes();
        }

        using PdfPigDocument pdf = PdfPigDocument.Open(new MemoryStream(bytes));
        Assert.Equal(2, pdf.NumberOfPages);
        string firstPage = pdf.GetPage(1).Text;
        string secondPage = pdf.GetPage(2).Text;
        Assert.Contains("AreaOne", firstPage);
        Assert.DoesNotContain("AreaTwo", firstPage);
        Assert.Contains("AreaTwo", secondPage);
        Assert.DoesNotContain("AreaOne", secondPage);
        Assert.DoesNotContain("UsedRangeTop", firstPage + secondPage);
        Assert.DoesNotContain("UsedRangeBottom", firstPage + secondPage);
        Assert.DoesNotContain(result.Warnings, warning => warning.Code == "WorksheetPrintArea");
    }

    [Fact]
    public void SaveAsPdf_ExcelWorkbook_Filters_Images_And_Charts_Outside_Print_Area() {
        string workbookPath = Path.Combine(_directoryWithFiles, "ExcelPdfPrintAreaMedia.xlsx");
        byte[] imageBytes = CreateMinimalRgbPng();

        byte[] bytes;
        using (ExcelDocument document = ExcelDocument.Create(workbookPath, new ExcelCreateOptions { FirstSheetName = "Report" })) {
            ExcelSheet sheet = document.Sheets[0];
            sheet.Cell(2, 2, "Category");
            sheet.Cell(2, 3, "Value");
            sheet.Cell(3, 2, "Inside");
            sheet.Cell(3, 3, 10);
            sheet.Cell(10, 1, "OutsideData");
            sheet.AddImage(3, 2, imageBytes, new OfficeIMO.Excel.ExcelImageOptions {
                ContentType = "image/png",
                Width = OfficeIMO.OfficeLength.Pixels(12),
                Height = OfficeIMO.OfficeLength.Pixels(12),
                Name = "Inside image"
            });
            sheet.AddImage(10, 1, imageBytes, new OfficeIMO.Excel.ExcelImageOptions {
                ContentType = "image/png",
                Width = OfficeIMO.OfficeLength.Pixels(12),
                Height = OfficeIMO.OfficeLength.Pixels(12),
                Name = "Outside image"
            });
            sheet.AddChart(OfficeChartType.ColumnClustered, "B2:C3", new OfficeIMO.Excel.ExcelChartOptions { Row = 3, Column = 2, Width = OfficeLength.Pixels(220), Height = OfficeLength.Pixels(120), Title = "Inside Chart" });
            sheet.AddChart(OfficeChartType.ColumnClustered, "B2:C3", new OfficeIMO.Excel.ExcelChartOptions { Row = 10, Column = 1, Width = OfficeLength.Pixels(220), Height = OfficeLength.Pixels(120), Title = "Outside Chart" });
            document.SetPrintArea(sheet, "B2:C3");
            document.Save();

            bytes = document.ToPdf(new ExcelPdfExportOptions {
                IncludeSheetHeadings = false,
                UseWorksheetPrintAreas = true,
                PageSize = new PdfCore.PdfPageSize(OfficeIMO.OfficeLength.Points(420),OfficeIMO.OfficeLength.Points(320)),
                Margins = PdfCore.PdfPageMargins.Uniform(OfficeIMO.OfficeLength.Points(24))
            }).ToBytes(new OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true });
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
        using (ExcelDocument document = ExcelDocument.Create(workbookPath, new ExcelCreateOptions { FirstSheetName = "PageSetup" })) {
            ExcelSheet sheet = document.Sheets[0];
            sheet.Cell(1, 1, "Name");
            sheet.Cell(2, 1, "WorksheetPageSetup");
            sheet.SetOrientation(OfficePageOrientation.Landscape);
            sheet.SetMargins(left: OfficeLength.Inches(0.25), right: OfficeLength.Inches(0.25), top: OfficeLength.Inches(0.5), bottom: OfficeLength.Inches(0.5));
            document.Save();

            bytes = document.ToPdf(new ExcelPdfExportOptions {
                IncludeSheetHeadings = false
            }).ToBytes(new OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true });
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
        using (ExcelDocument document = ExcelDocument.Create(workbookPath, new ExcelCreateOptions { FirstSheetName = "PaperSize" })) {
            ExcelSheet sheet = document.Sheets[0];
            sheet.Cell(1, 1, "Name");
            sheet.Cell(2, 1, "WorksheetPaperSize");
            sheet.SetPageSetup(paperSize: ExcelPaperSize.A4);
            document.Save();

            bytes = document.ToPdf(new ExcelPdfExportOptions {
                IncludeSheetHeadings = false
            }).ToBytes(new OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true });
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
        using (ExcelDocument document = ExcelDocument.Create(workbookPath, new ExcelCreateOptions { FirstSheetName = "PaperSize" })) {
            ExcelSheet sheet = document.Sheets[0];
            sheet.Cell(1, 1, "Name");
            sheet.Cell(2, 1, "PdfOptionsPageSize");
            sheet.SetPageSetup(paperSize: ExcelPaperSize.A4);
            document.Save();

            bytes = document.ToPdf(new ExcelPdfExportOptions {
                IncludeSheetHeadings = false,
                PdfOptions = new PdfCore.PdfCreateOptions {
                    PageSize = new PdfCore.PdfPageSize(OfficeIMO.OfficeLength.Points(300),OfficeIMO.OfficeLength.Points(220))
                }
            }).ToBytes(new OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true });
        }

        PdfCore.PdfDocumentInfo info = PdfCore.PdfInspector.Inspect(bytes);
        PdfCore.PdfPageInfo page = Assert.Single(info.Pages);
        Assert.Equal(300D, page.Width, 1D);
        Assert.Equal(220D, page.Height, 1D);
    }

    [Fact]
    public void SaveAsPdf_ExcelWorkbook_HonorsExplicitProfileBaselinePageSize() {
        string workbookPath = Path.Combine(
            _directoryWithFiles,
            "ExcelPdfExplicitProfileBaselinePageSize.xlsx");

        byte[] bytes;
        using (ExcelDocument document = ExcelDocument.Create(workbookPath, new ExcelCreateOptions { FirstSheetName = "PaperSize" })) {
            ExcelSheet sheet = document.Sheets[0];
            sheet.Cell(1, 1, "Name");
            sheet.Cell(2, 1, "ExplicitLetter");
            sheet.SetPageSetup(paperSize: ExcelPaperSize.A3);
            document.Save();

            var options = new ExcelPdfExportOptions {
                IncludeSheetHeadings = false
            }.UseRenderingProfile(OfficeRenderingProfile.Managed);
            options.PdfOptions!.PageSize = new PdfCore.PdfPageSize(OfficeIMO.OfficeLength.Points(612),OfficeIMO.OfficeLength.Points(792));

            bytes = document.ToPdf(options).ToBytes(new OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true });
        }

        PdfCore.PdfPageInfo page = Assert.Single(PdfCore.PdfInspector.Inspect(bytes).Pages);
        Assert.Equal(612D, page.Width, 1D);
        Assert.Equal(792D, page.Height, 1D);
    }

    [Fact]
    public void SaveAsPdf_ExcelWorkbook_Uses_Default_Pdf_PageSize_When_Worksheet_PageSetup_Disabled() {
        string workbookPath = Path.Combine(_directoryWithFiles, "ExcelPdfDisableWorksheetPageSetup.xlsx");

        byte[] bytes;
        using (ExcelDocument document = ExcelDocument.Create(workbookPath, new ExcelCreateOptions { FirstSheetName = "PaperSize" })) {
            ExcelSheet sheet = document.Sheets[0];
            sheet.Cell(1, 1, "Name");
            sheet.Cell(2, 1, "DefaultPdfPageSize");
            sheet.SetPageSetup(paperSize: ExcelPaperSize.A4);
            document.Save();

            bytes = document.ToPdf(new ExcelPdfExportOptions {
                IncludeSheetHeadings = false,
                UseWorksheetPageSetup = false
            }).ToBytes(new OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true });
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
        using (ExcelDocument document = ExcelDocument.Create(workbookPath, new ExcelCreateOptions { FirstSheetName = "PageSetup" })) {
            ExcelSheet sheet = document.Sheets[0];
            sheet.Cell(1, 1, "Name");
            sheet.Cell(1, 2, "Value");
            for (int row = 2; row <= 10; row++) {
                sheet.Cell(row, 1, "Row " + row.ToString(CultureInfo.InvariantCulture));
                sheet.Cell(row, 2, row);
            }

            for (int row = 1; row <= 10; row++) {
                sheet.SetRowHeight(row, ExcelColumnWidth.Length(OfficeIMO.OfficeLength.Points(36D)));
            }

            sheet.SetPageSetup(fitToWidth: 1U, fitToHeight: 1U);
            document.Save();

            pdfDocument = document.ToPdf(new ExcelPdfExportOptions {
                IncludeSheetHeadings = false,
                HeaderRowCount = 1,
                WorksheetLayout = ExcelPdfWorksheetLayoutMode.FlowTable,
                PageSize = new PdfCore.PdfPageSize(OfficeIMO.OfficeLength.Points(220),OfficeIMO.OfficeLength.Points(144)),
                Margins = PdfCore.PdfPageMargins.Uniform(OfficeIMO.OfficeLength.Points(18))
            });
        }

        PdfCore.PageBlock page = Assert.IsType<PdfCore.PageBlock>(Assert.Single(pdfDocument.Blocks));
        PdfCore.TableBlock table = Assert.Single(page.Blocks.OfType<PdfCore.TableBlock>());
        Assert.NotNull(table.Style);
        PdfCore.PdfTableStyle style = table.Style!;
        Assert.Equal(1, style.HeaderRowCount);
        Assert.Null(style.FixedRowHeights);
        Assert.NotNull(style.RowMinHeights);
        Assert.Equal(10, style.RowMinHeights!.Count);
        Assert.InRange(style.RowMinHeights!.Sum(height => height?.ToPoints() ?? 0D), 107D, 108.1D);
        Assert.True(style.CellPaddingY.ToPoints() < 3D, "Fit-to-height scaling should shrink default vertical padding along with row heights.");
        Assert.NotNull(style.FontSize?.ToPoints());
        Assert.NotNull(style.HeaderFontSize?.ToPoints());
        Assert.NotNull(style.FooterFontSize?.ToPoints());
        Assert.InRange(style.FontSize!.Value.ToPoints(), 3.2D, 3.4D);
        Assert.Equal(style.FontSize.Value.ToPoints(), style.HeaderFontSize!.Value.ToPoints());
        Assert.Equal(style.FontSize.Value.ToPoints(), style.FooterFontSize!.Value.ToPoints());
    }

    [Fact]
    public void SaveAsPdf_ExcelWorkbook_Uses_Print_Title_Rows_As_Repeating_Table_Header() {
        string workbookPath = Path.Combine(_directoryWithFiles, "ExcelPdfPrintTitles.xlsx");

        byte[] bytes;
        using (ExcelDocument document = ExcelDocument.Create(workbookPath, new ExcelCreateOptions { FirstSheetName = "LongReport" })) {
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

            bytes = document.ToPdf(new ExcelPdfExportOptions {
                IncludeSheetHeadings = false,
                HeaderRowCount = 0,
                PageSize = new PdfCore.PdfPageSize(OfficeIMO.OfficeLength.Points(300),OfficeIMO.OfficeLength.Points(220)),
                Margins = PdfCore.PdfPageMargins.Uniform(OfficeIMO.OfficeLength.Points(18))
            }).ToBytes(new OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true });
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
        using (ExcelDocument document = ExcelDocument.Create(workbookPath, new ExcelCreateOptions { FirstSheetName = "ManualBreaks" })) {
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

            var options = new ExcelPdfExportOptions {
                IncludeSheetHeadings = false,
                HeaderRowCount = 1,
                PageSize = new PdfCore.PdfPageSize(OfficeIMO.OfficeLength.Points(420),OfficeIMO.OfficeLength.Points(420)),
                Margins = PdfCore.PdfPageMargins.Uniform(OfficeIMO.OfficeLength.Points(24))
            };
            bytes = document.ToPdf(options).ToBytes(new OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true });

            options.UseWorksheetPageBreaks = false;
            disabledBytes = document.ToPdf(options).ToBytes(new OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true });
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
        using (ExcelDocument document = ExcelDocument.Create(workbookPath, new ExcelCreateOptions { FirstSheetName = "ManualBreaks" })) {
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

            bytes = document.ToPdf(new ExcelPdfExportOptions {
                IncludeSheetHeadings = false,
                HeaderRowCount = 0,
                PageSize = new PdfCore.PdfPageSize(OfficeIMO.OfficeLength.Points(420),OfficeIMO.OfficeLength.Points(360)),
                Margins = PdfCore.PdfPageMargins.Uniform(OfficeIMO.OfficeLength.Points(24))
            }).ToBytes(new OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true });
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
        using (ExcelDocument document = ExcelDocument.Create(workbookPath, new ExcelCreateOptions { FirstSheetName = "ManualColumnBreaks" })) {
            ExcelSheet sheet = document.Sheets[0];
            sheet.Cell(1, 1, "LeftHeader");
            sheet.Cell(1, 2, "LeftTail");
            sheet.Cell(1, 3, "RightHeader");
            sheet.Cell(1, 4, "RightTail");
            sheet.Cell(2, 1, "LeftValueA");
            sheet.Cell(2, 2, "LeftValueB");
            sheet.Cell(2, 3, "RightValueC");
            sheet.Cell(2, 4, "RightValueD");
            sheet.SetColumnWidth(1, ExcelColumnWidth.Characters(16));
            sheet.SetColumnWidth(3, ExcelColumnWidth.Characters(22));
            sheet.AddManualColumnPageBreak(2);

            Assert.Equal(new[] { 2 }, sheet.GetManualColumnPageBreaks());
            document.Save();

            var options = new ExcelPdfExportOptions {
                IncludeSheetHeadings = false,
                HeaderRowCount = 1,
                PageSize = new PdfCore.PdfPageSize(OfficeIMO.OfficeLength.Points(560),OfficeIMO.OfficeLength.Points(320)),
                Margins = PdfCore.PdfPageMargins.Uniform(OfficeIMO.OfficeLength.Points(24))
            };
            bytes = document.ToPdf(options).ToBytes(new OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true });

            options.UseWorksheetPageBreaks = false;
            disabledBytes = document.ToPdf(options).ToBytes(new OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true });
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
        using (ExcelDocument document = ExcelDocument.Create(workbookPath, new ExcelCreateOptions { FirstSheetName = "PageOrder" })) {
            ExcelSheet sheet = document.Sheets[0];
            sheet.Cell(1, 1, "TopLeftPage");
            sheet.Cell(3, 1, "BottomLeftPage");
            sheet.Cell(1, 3, "TopRightPage");
            sheet.Cell(3, 3, "BottomRightPage");
            sheet.AddManualRowPageBreak(2);
            sheet.AddManualColumnPageBreak(2);
            document.Save();

            bytes = document.ToPdf(new ExcelPdfExportOptions {
                IncludeSheetHeadings = false,
                HeaderRowCount = 0,
                PageSize = new PdfCore.PdfPageSize(OfficeIMO.OfficeLength.Points(360),OfficeIMO.OfficeLength.Points(260)),
                Margins = PdfCore.PdfPageMargins.Uniform(OfficeIMO.OfficeLength.Points(24))
            }).ToBytes(new OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true });
        }

        using PdfPigDocument pdf = PdfPigDocument.Open(new MemoryStream(bytes));
        Assert.Equal(4, pdf.NumberOfPages);
        Assert.Contains("TopLeftPage", pdf.GetPage(1).Text);
        Assert.Contains("BottomLeftPage", pdf.GetPage(2).Text);
        Assert.Contains("TopRightPage", pdf.GetPage(3).Text);
        Assert.Contains("BottomRightPage", pdf.GetPage(4).Text);
    }

}
