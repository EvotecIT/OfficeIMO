using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel;
using OfficeIMO.Excel.Pdf;
using System.Globalization;
using System.Text;
using UglyToad.PdfPig;
using Xunit;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Tests;

public partial class Excel {
    [Fact]
    public void SaveAsPdf_ExcelWorkbook_Exports_Worksheet_UsedRange_To_Table() {
        string workbookPath = Path.Combine(_directoryWithFiles, "ExcelPdfWorkbook.xlsx");
        string pdfPath = Path.Combine(_directoryWithFiles, "ExcelPdfWorkbook.pdf");

        using (ExcelDocument document = ExcelDocument.Create(workbookPath, "Sales")) {
            ExcelSheet sheet = document.Sheets[0];
            sheet.Cell(1, 1, "Product");
            sheet.Cell(1, 2, "Amount");
            sheet.Cell(2, 1, "Licenses");
            sheet.Cell(2, 2, 1250.5);
            sheet.Cell(3, 1, "Support");
            sheet.Cell(3, 2, 250);
            document.Save(false);

            document.SaveAsPdf(pdfPath);
        }

        Assert.True(File.Exists(pdfPath));
        using PdfDocument pdf = PdfDocument.Open(pdfPath);
        string text = pdf.GetPage(1).Text;
        Assert.Contains("Sales", text);
        Assert.Contains("Product", text);
        Assert.Contains("Licenses", text);
        Assert.Contains("1250.5", text);
    }

    [Fact]
    public void SaveAsPdf_ExcelWorkbook_Respects_Selected_Sheets() {
        string workbookPath = Path.Combine(_directoryWithFiles, "ExcelPdfSelectedSheets.xlsx");

        byte[] bytes;
        using (ExcelDocument document = ExcelDocument.Create(workbookPath)) {
            ExcelSheet summary = document.AddWorkSheet("Summary");
            summary.Cell(1, 1, "Metric");
            summary.Cell(2, 1, "SelectedValue");
            ExcelSheet internalSheet = document.AddWorkSheet("Internal");
            internalSheet.Cell(1, 1, "HiddenValue");
            document.Save(false);

            bytes = document.SaveAsPdf(new ExcelPdfSaveOptions {
                SheetNames = new[] { "Summary" }
            });
        }

        using PdfDocument pdf = PdfDocument.Open(new MemoryStream(bytes));
        string text = pdf.GetPage(1).Text;
        Assert.Contains("Summary", text);
        Assert.Contains("SelectedValue", text);
        Assert.DoesNotContain("HiddenValue", text);
    }

    [Fact]
    public void SaveAsPdf_ExcelWorkbook_Exports_Worksheet_Images() {
        string workbookPath = Path.Combine(_directoryWithFiles, "ExcelPdfImages.xlsx");

        byte[] imageBytes = CreateMinimalRgbPng();
        byte[] bytes;
        using (ExcelDocument document = ExcelDocument.Create(workbookPath, "Images")) {
            ExcelSheet sheet = document.Sheets[0];
            sheet.Cell(1, 1, "ImageMarker");
            sheet.AddImage(2, 1, imageBytes, "image/png", widthPixels: 24, heightPixels: 16, name: "Pdf Logo", altText: "PDF logo");

            ExcelImage image = Assert.Single(sheet.Images);
            Assert.Equal("Pdf Logo", image.Name);
            Assert.Equal("PDF logo", image.Description);
            Assert.Equal(2, image.RowIndex);
            Assert.Equal(1, image.ColumnIndex);
            Assert.Equal(24, image.WidthPixels);
            Assert.Equal(16, image.HeightPixels);
            Assert.Equal("image/png", image.ContentType);
            Assert.Equal(imageBytes, image.GetBytes());

            document.Save(false);

            bytes = document.SaveAsPdf(new ExcelPdfSaveOptions {
                IncludeSheetHeadings = false,
                HeaderRowCount = 0
            });
        }

        using PdfDocument pdf = PdfDocument.Open(new MemoryStream(bytes));
        Assert.Contains("ImageMarker", pdf.GetPage(1).Text);

        var extractedImages = PdfCore.PdfImageExtractor.ExtractImages(bytes);
        var extractedImage = Assert.Single(extractedImages);
        Assert.Equal(1, extractedImage.PageNumber);
        Assert.Equal("png", extractedImage.FileExtension);
        Assert.Equal("image/png", extractedImage.MimeType);
        Assert.True(extractedImage.IsImageFile);
        Assert.Equal(1, extractedImage.Width);
        Assert.Equal(1, extractedImage.Height);
    }

    [Fact]
    public void SaveAsPdf_ExcelWorkbook_Warns_And_Skips_Invalid_Worksheet_Image_Bytes() {
        string workbookPath = Path.Combine(_directoryWithFiles, "ExcelPdfInvalidImageBytes.xlsx");
        byte[] invalidPngBytes = new byte[] {
            137, 80, 78, 71, 13, 10, 26, 10,
            0, 0, 0, 13,
            73, 72, 68, 82,
            0, 0, 0, 1,
            0, 0, 0, 1,
            16, 2, 0, 0, 0,
            0, 0, 0, 0,
            0, 0, 0, 0,
            73, 69, 78, 68,
            0, 0, 0, 0
        };
        var options = new ExcelPdfSaveOptions {
            IncludeSheetHeadings = false
        };

        byte[] bytes;
        using (ExcelDocument document = ExcelDocument.Create(workbookPath, "Images")) {
            ExcelSheet sheet = document.Sheets[0];
            sheet.Cell(1, 1, "ImageMarker");
            sheet.AddImage(2, 1, invalidPngBytes, "image/png", widthPixels: 24, heightPixels: 16, name: "Invalid PNG");
            document.Save(false);

            bytes = document.SaveAsPdf(options);
        }

        using PdfDocument pdf = PdfDocument.Open(new MemoryStream(bytes));
        Assert.Contains("ImageMarker", pdf.GetPage(1).Text);
        Assert.Empty(PdfCore.PdfImageExtractor.ExtractImages(bytes));
        Assert.Contains(options.Warnings, warning => warning.SheetName == "Images" && warning.Feature == "WorksheetImage");
    }

    [Fact]
    public void SaveAsPdf_ExcelWorkbook_Embeds_Worksheet_Images_In_Anchored_Table_Cells() {
        string workbookPath = Path.Combine(_directoryWithFiles, "ExcelPdfAnchoredImageCell.xlsx");

        byte[] imageBytes = CreateMinimalRgbPng();
        byte[] bytes;
        using (ExcelDocument document = ExcelDocument.Create(workbookPath, "Images")) {
            ExcelSheet sheet = document.Sheets[0];
            sheet.Cell(1, 1, "Label");
            sheet.Cell(1, 2, "Visual");
            sheet.Cell(2, 1, "BeforeImageRow");
            sheet.Cell(3, 1, "AnchoredImageRow");
            sheet.Cell(4, 1, "AfterImageRow");
            sheet.AddImage(3, 2, imageBytes, "image/png", widthPixels: 72, heightPixels: 72, name: "Anchored Cell Image");
            document.Save(false);

            bytes = document.SaveAsPdf(new ExcelPdfSaveOptions {
                IncludeSheetHeadings = false,
                HeaderRowCount = 1,
                PageSize = new PdfCore.PageSize(420, 360),
                Margins = PdfCore.PageMargins.Uniform(24)
            });
        }

        using PdfDocument pdf = PdfDocument.Open(new MemoryStream(bytes));
        UglyToad.PdfPig.Content.Page page = pdf.GetPage(1);
        double beforeRowY = FindWordStartY(page, "BeforeImageRow");
        double anchoredRowY = FindWordStartY(page, "AnchoredImageRow");
        double afterRowY = FindWordStartY(page, "AfterImageRow");

        double gapBeforeAnchoredRow = beforeRowY - anchoredRowY;
        double gapAfterAnchoredRow = anchoredRowY - afterRowY;
        Assert.True(gapAfterAnchoredRow > gapBeforeAnchoredRow + 20, "The image should increase the anchored table row height instead of flowing before the table.");

        var extractedImage = Assert.Single(PdfCore.PdfImageExtractor.ExtractImages(bytes));
        Assert.Equal(1, extractedImage.PageNumber);
        Assert.Equal("image/png", extractedImage.MimeType);
    }

    [Fact]
    public void SaveAsPdf_ExcelWorkbook_Omits_Hidden_Sheets_By_Default() {
        string workbookPath = Path.Combine(_directoryWithFiles, "ExcelPdfHiddenSheets.xlsx");

        byte[] visibleBytes;
        byte[] explicitHiddenBytes;
        using (ExcelDocument document = ExcelDocument.Create(workbookPath)) {
            ExcelSheet visible = document.AddWorkSheet("Visible");
            visible.Cell(1, 1, "VisibleSheetValue");
            ExcelSheet hidden = document.AddWorkSheet("Hidden");
            hidden.Cell(1, 1, "HiddenSheetValue");
            hidden.SetHidden(true);
            Assert.True(hidden.Hidden);
            document.Save(false);

            visibleBytes = document.SaveAsPdf(new ExcelPdfSaveOptions {
                IncludeSheetHeadings = false,
                HeaderRowCount = 0
            });

            explicitHiddenBytes = document.SaveAsPdf(new ExcelPdfSaveOptions {
                SheetNames = new[] { "Hidden" },
                IncludeSheetHeadings = false,
                HeaderRowCount = 0
            });
        }

        using PdfDocument visiblePdf = PdfDocument.Open(new MemoryStream(visibleBytes));
        string visibleText = visiblePdf.GetPage(1).Text;
        Assert.Contains("VisibleSheetValue", visibleText);
        Assert.DoesNotContain("HiddenSheetValue", visibleText);

        using PdfDocument explicitHiddenPdf = PdfDocument.Open(new MemoryStream(explicitHiddenBytes));
        string hiddenText = explicitHiddenPdf.GetPage(1).Text;
        Assert.Contains("HiddenSheetValue", hiddenText);
    }

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
            document.Save(false);

            bytes = document.SaveAsPdf(new ExcelPdfSaveOptions {
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
            document.Save(false);

            bytes = document.SaveAsPdf(new ExcelPdfSaveOptions {
                IncludeSheetHeadings = false
            });
        }

        using PdfDocument pdf = PdfDocument.Open(new MemoryStream(bytes));
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
            document.Save(false);

            bytes = document.SaveAsPdf(new ExcelPdfSaveOptions {
                IncludeSheetHeadings = false
            });
        }

        using PdfDocument pdf = PdfDocument.Open(new MemoryStream(bytes));
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
        using (ExcelDocument document = ExcelDocument.Create(workbookPath, "Report")) {
            ExcelSheet sheet = document.Sheets[0];
            sheet.Cell(1, 1, "UsedRangeTop");
            sheet.Cell(2, 2, "AreaOne");
            sheet.Cell(2, 4, "AreaTwo");
            sheet.Cell(5, 5, "UsedRangeBottom");
            document.Save(false);
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
            bytes = document.SaveAsPdf(options);
        }

        using PdfDocument pdf = PdfDocument.Open(new MemoryStream(bytes));
        string text = pdf.GetPage(1).Text;
        Assert.Contains("UsedRangeTop", text);
        Assert.Contains("AreaOne", text);
        Assert.Contains("AreaTwo", text);
        Assert.Contains("UsedRangeBottom", text);
        Assert.Contains(options.Warnings, warning => warning.SheetName == "Report" && warning.Feature == "WorksheetPrintArea");
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
            document.Save(false);

            bytes = document.SaveAsPdf(new ExcelPdfSaveOptions {
                IncludeSheetHeadings = false,
                UseWorksheetPrintAreas = true,
                PageSize = new PdfCore.PageSize(420, 320),
                Margins = PdfCore.PageMargins.Uniform(24)
            });
        }

        using PdfDocument pdf = PdfDocument.Open(new MemoryStream(bytes));
        string text = pdf.GetPage(1).Text;
        Assert.Contains("Inside Chart", text);
        Assert.DoesNotContain("Outside Chart", text);
        Assert.DoesNotContain("OutsideData", text);
        Assert.Single(PdfCore.PdfImageExtractor.ExtractImages(bytes));
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
            document.Save(false);

            bytes = document.SaveAsPdf(new ExcelPdfSaveOptions {
                IncludeSheetHeadings = false
            });
        }

        PdfCore.PdfDocumentInfo info = PdfCore.PdfInspector.Inspect(bytes);
        PdfCore.PdfPageInfo page = Assert.Single(info.Pages);
        Assert.True(page.Width > page.Height, $"Expected worksheet landscape orientation. Width: {page.Width}, height: {page.Height}.");

        using PdfDocument pdf = PdfDocument.Open(new MemoryStream(bytes));
        double firstLetterX = pdf.GetPage(1).Letters.First(letter => letter.Value == "N").StartBaseLine.X;
        Assert.InRange(firstLetterX, 17D, 36D);
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
            document.Save(false);

            bytes = document.SaveAsPdf(new ExcelPdfSaveOptions {
                IncludeSheetHeadings = false,
                HeaderRowCount = 0,
                PageSize = new PdfCore.PageSize(300, 220),
                Margins = PdfCore.PageMargins.Uniform(18)
            });
        }

        using PdfDocument pdf = PdfDocument.Open(new MemoryStream(bytes));
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
            document.Save(false);

            var options = new ExcelPdfSaveOptions {
                IncludeSheetHeadings = false,
                HeaderRowCount = 1,
                PageSize = new PdfCore.PageSize(420, 420),
                Margins = PdfCore.PageMargins.Uniform(24)
            };
            bytes = document.SaveAsPdf(options);

            options.UseWorksheetPageBreaks = false;
            disabledBytes = document.SaveAsPdf(options);
        }

        using PdfDocument pdf = PdfDocument.Open(new MemoryStream(bytes));
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

        using PdfDocument disabledPdf = PdfDocument.Open(new MemoryStream(disabledBytes));
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
            document.Save(false);

            bytes = document.SaveAsPdf(new ExcelPdfSaveOptions {
                IncludeSheetHeadings = false,
                HeaderRowCount = 0,
                PageSize = new PdfCore.PageSize(420, 360),
                Margins = PdfCore.PageMargins.Uniform(24)
            });
        }

        using PdfDocument pdf = PdfDocument.Open(new MemoryStream(bytes));
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
            document.Save(false);

            var options = new ExcelPdfSaveOptions {
                IncludeSheetHeadings = false,
                HeaderRowCount = 1,
                PageSize = new PdfCore.PageSize(560, 320),
                Margins = PdfCore.PageMargins.Uniform(24)
            };
            bytes = document.SaveAsPdf(options);

            options.UseWorksheetPageBreaks = false;
            disabledBytes = document.SaveAsPdf(options);
        }

        using PdfDocument pdf = PdfDocument.Open(new MemoryStream(bytes));
        Assert.Equal(2, pdf.NumberOfPages);
        string firstPage = pdf.GetPage(1).Text;
        string secondPage = pdf.GetPage(2).Text;
        Assert.Contains("LeftValueA", firstPage);
        Assert.Contains("LeftValueB", firstPage);
        Assert.DoesNotContain("RightValueC", firstPage);
        Assert.Contains("RightValueC", secondPage);
        Assert.Contains("RightValueD", secondPage);
        Assert.DoesNotContain("LeftValueA", secondPage);

        using PdfDocument disabledPdf = PdfDocument.Open(new MemoryStream(disabledBytes));
        Assert.Equal(1, disabledPdf.NumberOfPages);
        Assert.Contains("LeftValueA", disabledPdf.GetPage(1).Text);
        Assert.Contains("RightValueC", disabledPdf.GetPage(1).Text);
    }

    [Fact]
    public void SaveAsPdf_ExcelWorkbook_Maps_Worksheet_HeaderFooter_Text_Zones() {
        string workbookPath = Path.Combine(_directoryWithFiles, "ExcelPdfHeaderFooter.xlsx");

        byte[] bytes;
        using (ExcelDocument document = ExcelDocument.Create(workbookPath, "Dashboard")) {
            ExcelSheet sheet = document.Sheets[0];
            sheet.Cell(1, 1, "Metric");
            sheet.Cell(2, 1, "HeaderFooterBody");
            sheet.SetHeaderFooter(
                headerLeft: "Left Header",
                headerCenter: "Dashboard Header",
                headerRight: "Page &P of &N",
                footerLeft: "Sheet &A",
                footerRight: "Right Footer");
            document.Save(false);

            bytes = document.SaveAsPdf(new ExcelPdfSaveOptions {
                IncludeSheetHeadings = false,
                PageSize = new PdfCore.PageSize(420, 320),
                Margins = PdfCore.PageMargins.Uniform(54)
            });
        }

        using PdfDocument pdf = PdfDocument.Open(new MemoryStream(bytes));
        string text = pdf.GetPage(1).Text;
        Assert.Contains("Left Header", text);
        Assert.Contains("Dashboard Header", text);
        Assert.Contains("Page 1 of 1", text);
        Assert.Contains("Sheet Dashboard", text);
        Assert.Contains("Right Footer", text);
        Assert.Contains("HeaderFooterBody", text);
    }

    [Fact]
    public void SaveAsPdf_ExcelWorkbook_Maps_Worksheet_HeaderFooter_DateTime_And_File_Fields() {
        string workbookPath = Path.Combine(_directoryWithFiles, "ExcelPdfHeaderFooterFields.xlsx");
        DateTime printedAt = new DateTime(2026, 5, 31, 14, 35, 0);

        var options = new ExcelPdfSaveOptions {
            IncludeSheetHeadings = false,
            HeaderRowCount = 0,
            HeaderFooterDateTimeProvider = () => printedAt,
            PageSize = new PdfCore.PageSize(1800, 320),
            Margins = PdfCore.PageMargins.Uniform(54)
        };

        byte[] bytes;
        using (ExcelDocument document = ExcelDocument.Create(workbookPath, "Dashboard")) {
            ExcelSheet sheet = document.Sheets[0];
            sheet.Cell(1, 1, "HeaderFooterFieldBody");
            sheet.SetHeaderFooter(
                headerLeft: "Printed &D &T Dir &Z File &F",
                footerRight: "Page &P of &N");
            document.Save(false);

            bytes = document.SaveAsPdf(options);
        }

        using PdfDocument pdf = PdfDocument.Open(new MemoryStream(bytes));
        string text = pdf.GetPage(1).Text;
        Assert.Contains("Printed " + printedAt.ToString("d", CultureInfo.CurrentCulture), text);
        Assert.Contains(NormalizePdfTextSpaces(printedAt.ToString("t", CultureInfo.CurrentCulture)), NormalizePdfTextSpaces(text));
        Assert.Contains("Dir " + Path.GetDirectoryName(Path.GetFullPath(workbookPath)), text);
        Assert.Contains("File " + Path.GetFileName(workbookPath), text);
        Assert.Contains("Page 1 of 1", text);
        Assert.Contains("HeaderFooterFieldBody", text);
        Assert.DoesNotContain(options.Warnings, warning => warning.Feature == "WorksheetHeaderFooterField");
    }

    [Fact]
    public void SaveAsPdf_ExcelWorkbook_Maps_Simple_Worksheet_HeaderFooter_Formatting() {
        string workbookPath = Path.Combine(_directoryWithFiles, "ExcelPdfHeaderFooterFormatting.xlsx");

        var options = new ExcelPdfSaveOptions {
            IncludeSheetHeadings = false,
            HeaderRowCount = 0,
            PageSize = new PdfCore.PageSize(420, 320),
            Margins = PdfCore.PageMargins.Uniform(54)
        };

        byte[] bytes;
        using (ExcelDocument document = ExcelDocument.Create(workbookPath, "Formatting")) {
            ExcelSheet sheet = document.Sheets[0];
            sheet.Cell(1, 1, "HeaderFooterFormattingBody");
            sheet.SetHeaderFooter(
                headerCenter: "&\"Arial,Bold\"&18&KFF0000Styled Header",
                footerCenter: "&\"Times New Roman,Italic\"&10&K0000FFStyled Footer");
            document.Save(false);

            bytes = document.SaveAsPdf(options);
        }

        using PdfDocument pdf = PdfDocument.Open(new MemoryStream(bytes));
        string text = pdf.GetPage(1).Text;
        Assert.Contains("Styled Header", text);
        Assert.Contains("Styled Footer", text);
        Assert.Contains("HeaderFooterFormattingBody", text);

        string rawPdf = Encoding.ASCII.GetString(bytes);
        Assert.Contains("1 0 0 rg", rawPdf, StringComparison.Ordinal);
        Assert.Contains("0 0 1 rg", rawPdf, StringComparison.Ordinal);
        Assert.Contains("Helvetica-Bold", rawPdf, StringComparison.Ordinal);
        Assert.Contains("Times-Italic", rawPdf, StringComparison.Ordinal);
        Assert.Contains(" 18 Tf", rawPdf, StringComparison.Ordinal);
        Assert.Contains(" 10 Tf", rawPdf, StringComparison.Ordinal);
        Assert.DoesNotContain(options.Warnings, warning => warning.Feature == "WorksheetHeaderFooterFormatting");
    }

    [Fact]
    public void SaveAsPdf_ExcelWorkbook_Warns_For_Mixed_Worksheet_HeaderFooter_Formatting() {
        string workbookPath = Path.Combine(_directoryWithFiles, "ExcelPdfHeaderFooterMixedFormatting.xlsx");

        var options = new ExcelPdfSaveOptions {
            IncludeSheetHeadings = false,
            HeaderRowCount = 0,
            PageSize = new PdfCore.PageSize(420, 320),
            Margins = PdfCore.PageMargins.Uniform(54)
        };

        byte[] bytes;
        using (ExcelDocument document = ExcelDocument.Create(workbookPath, "MixedFormatting")) {
            ExcelSheet sheet = document.Sheets[0];
            sheet.Cell(1, 1, "MixedFormattingBody");
            sheet.SetHeaderFooter(
                headerLeft: "&KFF0000Red Left",
                headerCenter: "Plain Center",
                footerCenter: "Plain Footer");
            document.Save(false);

            bytes = document.SaveAsPdf(options);
        }

        using PdfDocument pdf = PdfDocument.Open(new MemoryStream(bytes));
        string text = pdf.GetPage(1).Text;
        Assert.Contains("Red Left", text);
        Assert.Contains("Plain Center", text);
        Assert.Contains("MixedFormattingBody", text);

        string rawPdf = Encoding.ASCII.GetString(bytes);
        Assert.DoesNotContain("1 0 0 rg", rawPdf, StringComparison.Ordinal);
        Assert.Contains(options.Warnings, warning => warning.SheetName == "MixedFormatting" && warning.Feature == "WorksheetHeaderFooterFormatting");
    }

    [Fact]
    public void SaveAsPdf_ExcelWorkbook_Maps_Worksheet_First_And_Even_HeaderFooter_Text_Zones() {
        string workbookPath = Path.Combine(_directoryWithFiles, "ExcelPdfHeaderFooterVariants.xlsx");

        byte[] bytes;
        using (ExcelDocument document = ExcelDocument.Create(workbookPath, "Ledger")) {
            ExcelSheet sheet = document.Sheets[0];
            sheet.Cell(1, 1, "Entry");
            sheet.Cell(1, 2, "Amount");
            for (int row = 2; row <= 90; row++) {
                sheet.Cell(row, 1, "Ledger row " + row.ToString(CultureInfo.InvariantCulture));
                sheet.Cell(row, 2, row * 10);
            }

            sheet.SetHeaderFooter(headerCenter: "Odd Header &A", footerCenter: "Odd Footer &P");
            sheet.SetFirstPageHeaderFooter(headerCenter: "First Header &A", footerCenter: "First Footer &P");
            sheet.SetEvenPageHeaderFooter(headerCenter: "Even Header &A", footerCenter: "Even Footer &P");

            ExcelSheet.HeaderFooterSnapshot snapshot = sheet.GetHeaderFooter();
            Assert.True(snapshot.DifferentFirstPage);
            Assert.True(snapshot.DifferentOddEven);
            Assert.Equal("Odd Header &A", snapshot.HeaderCenter);
            Assert.Equal("First Header &A", snapshot.FirstHeaderCenter);
            Assert.Equal("Even Header &A", snapshot.EvenHeaderCenter);

            document.Save(false);

            bytes = document.SaveAsPdf(new ExcelPdfSaveOptions {
                IncludeSheetHeadings = false,
                HeaderRowCount = 1,
                PageSize = new PdfCore.PageSize(360, 220),
                Margins = PdfCore.PageMargins.Uniform(48)
            });
        }

        using PdfDocument pdf = PdfDocument.Open(new MemoryStream(bytes));
        Assert.True(pdf.NumberOfPages >= 3);
        string firstPage = pdf.GetPage(1).Text;
        string secondPage = pdf.GetPage(2).Text;
        string thirdPage = pdf.GetPage(3).Text;
        Assert.Contains("First Header Ledger", firstPage);
        Assert.Contains("First Footer 1", firstPage);
        Assert.Contains("Even Header Ledger", secondPage);
        Assert.Contains("Even Footer 2", secondPage);
        Assert.Contains("Odd Header Ledger", thirdPage);
        Assert.Contains("Odd Footer 3", thirdPage);
    }

    [Fact]
    public void SaveAsPdf_ExcelWorkbook_Preserves_Blank_First_And_Even_HeaderFooter_Variants() {
        string workbookPath = Path.Combine(_directoryWithFiles, "ExcelPdfBlankHeaderFooterVariants.xlsx");

        byte[] bytes;
        using (ExcelDocument document = ExcelDocument.Create(workbookPath, "Ledger")) {
            ExcelSheet sheet = document.Sheets[0];
            sheet.Cell(1, 1, "Entry");
            for (int row = 2; row <= 90; row++) {
                sheet.Cell(row, 1, "Ledger row " + row.ToString(CultureInfo.InvariantCulture));
            }

            sheet.SetHeaderFooter(headerCenter: "Odd Header &A", footerCenter: "Odd Footer &P");
            sheet.SetFirstPageHeaderFooter();
            sheet.SetEvenPageHeaderFooter();
            document.Save(false);

            bytes = document.SaveAsPdf(new ExcelPdfSaveOptions {
                IncludeSheetHeadings = false,
                HeaderRowCount = 1,
                PageSize = new PdfCore.PageSize(360, 220),
                Margins = PdfCore.PageMargins.Uniform(48)
            });
        }

        using PdfDocument pdf = PdfDocument.Open(new MemoryStream(bytes));
        Assert.True(pdf.NumberOfPages >= 3);
        string firstPage = pdf.GetPage(1).Text;
        string secondPage = pdf.GetPage(2).Text;
        string thirdPage = pdf.GetPage(3).Text;
        Assert.DoesNotContain("Odd Header Ledger", firstPage);
        Assert.DoesNotContain("Odd Footer 1", firstPage);
        Assert.DoesNotContain("Odd Header Ledger", secondPage);
        Assert.DoesNotContain("Odd Footer 2", secondPage);
        Assert.Contains("Odd Header Ledger", thirdPage);
        Assert.Contains("Odd Footer 3", thirdPage);
    }

    [Fact]
    public void SaveAsPdf_ExcelWorkbook_Maps_Worksheet_HeaderFooter_Images() {
        string workbookPath = Path.Combine(_directoryWithFiles, "ExcelPdfHeaderFooterImages.xlsx");

        byte[] imageBytes = CreateMinimalRgbPng();
        byte[] bytes;
        byte[] disabledBytes;
        using (ExcelDocument document = ExcelDocument.Create(workbookPath, "Dashboard")) {
            ExcelSheet sheet = document.Sheets[0];
            sheet.Cell(1, 1, "Metric");
            sheet.Cell(2, 1, "HeaderFooterImageBody");
            sheet.SetHeaderFooter(headerCenter: "Logo Header");
            sheet.SetHeaderImage(HeaderFooterPosition.Center, imageBytes, "image/png", widthPoints: 24, heightPoints: 16);

            ExcelSheet.HeaderFooterSnapshot snapshot = sheet.GetHeaderFooter();
            Assert.True(snapshot.HeaderHasPicturePlaceholder);
            Assert.NotNull(snapshot.HeaderCenterImage);
            Assert.Equal(HeaderFooterPosition.Center, snapshot.HeaderCenterImage!.Position);
            Assert.Equal("image/png", snapshot.HeaderCenterImage.ContentType);
            Assert.Equal(24, snapshot.HeaderCenterImage.WidthPoints);
            Assert.Equal(16, snapshot.HeaderCenterImage.HeightPoints);
            Assert.Equal(imageBytes, snapshot.HeaderCenterImage.Bytes);

            document.Save(false);

            var options = new ExcelPdfSaveOptions {
                IncludeSheetHeadings = false,
                PageSize = new PdfCore.PageSize(420, 320),
                Margins = PdfCore.PageMargins.Uniform(54)
            };
            bytes = document.SaveAsPdf(options);

            disabledBytes = document.SaveAsPdf(new ExcelPdfSaveOptions {
                IncludeSheetHeadings = false,
                UseWorksheetHeaderFooterImages = false,
                PageSize = new PdfCore.PageSize(420, 320),
                Margins = PdfCore.PageMargins.Uniform(54)
            });
        }

        using PdfDocument pdf = PdfDocument.Open(new MemoryStream(bytes));
        string text = pdf.GetPage(1).Text;
        Assert.Contains("Logo Header", text);
        Assert.Contains("HeaderFooterImageBody", text);

        var extractedImages = PdfCore.PdfImageExtractor.ExtractImages(bytes);
        var extractedImage = Assert.Single(extractedImages);
        Assert.Equal(1, extractedImage.PageNumber);
        Assert.Equal("png", extractedImage.FileExtension);
        Assert.Equal("image/png", extractedImage.MimeType);
        Assert.True(extractedImage.IsImageFile);

        Assert.Empty(PdfCore.PdfImageExtractor.ExtractImages(disabledBytes));
    }

    [Fact]
    public void SaveAsPdf_ExcelWorkbook_Can_Disable_Worksheet_HeaderFooter_Text_Zones() {
        string workbookPath = Path.Combine(_directoryWithFiles, "ExcelPdfHeaderFooterDisabled.xlsx");

        byte[] bytes;
        using (ExcelDocument document = ExcelDocument.Create(workbookPath, "Dashboard")) {
            ExcelSheet sheet = document.Sheets[0];
            sheet.Cell(1, 1, "Metric");
            sheet.Cell(2, 1, "BodyOnly");
            sheet.SetHeaderFooter(headerCenter: "DoNotExportHeader", footerCenter: "DoNotExportFooter");
            document.Save(false);

            bytes = document.SaveAsPdf(new ExcelPdfSaveOptions {
                IncludeSheetHeadings = false,
                UseWorksheetHeadersAndFooters = false
            });
        }

        using PdfDocument pdf = PdfDocument.Open(new MemoryStream(bytes));
        string text = pdf.GetPage(1).Text;
        Assert.Contains("BodyOnly", text);
        Assert.DoesNotContain("DoNotExportHeader", text);
        Assert.DoesNotContain("DoNotExportFooter", text);
    }

    [Fact]
    public void SaveAsPdf_ExcelWorkbook_Maps_Basic_Cell_Font_And_Fill_Styles() {
        string workbookPath = Path.Combine(_directoryWithFiles, "ExcelPdfCellStyles.xlsx");

        byte[] bytes;
        using (ExcelDocument document = ExcelDocument.Create(workbookPath, "Styled")) {
            ExcelSheet sheet = document.Sheets[0];
            sheet.CellAt(1, 1)
                .SetValue("StyledCell")
                .SetBold()
                .SetItalic()
                .SetUnderline()
                .SetFontColor("112233")
                .SetFillColor("DDEEFF");
            sheet.CellAt(1, 2).SetValue("PlainCell");

            ExcelCellStyleSnapshot style = sheet.CellAt(1, 1).GetStyle();
            Assert.True(style.Bold);
            Assert.True(style.Italic);
            Assert.True(style.Underline);
            Assert.Equal("112233", style.FontColorHex);
            Assert.Equal("DDEEFF", style.FillColorHex);

            document.Save(false);

            bytes = document.SaveAsPdf(new ExcelPdfSaveOptions {
                IncludeSheetHeadings = false,
                HeaderRowCount = 0,
                PageSize = new PdfCore.PageSize(360, 220),
                Margins = PdfCore.PageMargins.Uniform(24)
            });
        }

        using PdfDocument pdf = PdfDocument.Open(new MemoryStream(bytes));
        string text = pdf.GetPage(1).Text;
        Assert.Contains("StyledCell", text);
        Assert.Contains("PlainCell", text);

        string rawPdf = Encoding.ASCII.GetString(bytes);
        Assert.Contains("0.067 0.133 0.2 rg", rawPdf, StringComparison.Ordinal);
        Assert.Contains("0.867 0.933 1 rg", rawPdf, StringComparison.Ordinal);
    }

    [Fact]
    public void SaveAsPdf_ExcelWorkbook_Maps_Conditional_ColorScale_Fills() {
        string workbookPath = Path.Combine(_directoryWithFiles, "ExcelPdfConditionalColorScale.xlsx");

        byte[] bytes;
        using (ExcelDocument document = ExcelDocument.Create(workbookPath, "Conditional")) {
            ExcelSheet sheet = document.Sheets[0];
            sheet.Cell(1, 1, "Score");
            sheet.Cell(2, 1, 0);
            sheet.Cell(3, 1, 50);
            sheet.Cell(4, 1, 100);
            sheet.AddConditionalColorScale("A2:A4", "FFFF0000", "FF00FF00");

            ExcelConditionalFormattingInfo rule = Assert.Single(sheet.GetConditionalFormattingRules("A2:A4"));
            Assert.Equal("ColorScale", rule.Type);
            Assert.Equal(new[] { "FFFF0000", "FF00FF00" }, rule.ColorScaleColors);

            document.Save(false);

            bytes = document.SaveAsPdf(new ExcelPdfSaveOptions {
                IncludeSheetHeadings = false,
                HeaderRowCount = 1,
                PageSize = new PdfCore.PageSize(360, 240),
                Margins = PdfCore.PageMargins.Uniform(24)
            });
        }

        using PdfDocument pdf = PdfDocument.Open(new MemoryStream(bytes));
        string text = pdf.GetPage(1).Text;
        Assert.Contains("Score", text);
        Assert.Contains("100", text);

        string rawPdf = Encoding.ASCII.GetString(bytes);
        Assert.Contains("1 0 0 rg", rawPdf, StringComparison.Ordinal);
        Assert.Contains("0.502 0.502 0 rg", rawPdf, StringComparison.Ordinal);
        Assert.Contains("0 1 0 rg", rawPdf, StringComparison.Ordinal);
    }

    [Fact]
    public void SaveAsPdf_ExcelWorkbook_Maps_Conditional_DataBar_Overlays() {
        string workbookPath = Path.Combine(_directoryWithFiles, "ExcelPdfConditionalDataBar.xlsx");

        byte[] bytes;
        using (ExcelDocument document = ExcelDocument.Create(workbookPath, "Conditional")) {
            ExcelSheet sheet = document.Sheets[0];
            sheet.Cell(1, 1, "Score");
            sheet.Cell(2, 1, 0);
            sheet.Cell(3, 1, 50);
            sheet.Cell(4, 1, 100);
            sheet.AddConditionalDataBar("A2:A4", "FF5B9BD5");

            ExcelConditionalFormattingInfo rule = Assert.Single(sheet.GetConditionalFormattingRules("A2:A4"));
            Assert.Equal("DataBar", rule.Type);
            Assert.Equal("FF5B9BD5", rule.DataBarColor);

            document.Save(false);

            bytes = document.SaveAsPdf(new ExcelPdfSaveOptions {
                IncludeSheetHeadings = false,
                HeaderRowCount = 1,
                PageSize = new PdfCore.PageSize(360, 240),
                Margins = PdfCore.PageMargins.Uniform(24)
            });
        }

        using PdfDocument pdf = PdfDocument.Open(new MemoryStream(bytes));
        string text = pdf.GetPage(1).Text;
        Assert.Contains("Score", text);
        Assert.Contains("100", text);

        string rawPdf = Encoding.ASCII.GetString(bytes);
        int barFillCount = rawPdf.Split(new[] { "0.357 0.608 0.835 rg" }, StringSplitOptions.None).Length - 1;

        Assert.Equal(2, barFillCount);
        Assert.Contains(" re f", rawPdf, StringComparison.Ordinal);
    }

    [Fact]
    public void SaveAsPdf_ExcelWorkbook_Maps_Conditional_IconSet_Indicators() {
        string workbookPath = Path.Combine(_directoryWithFiles, "ExcelPdfConditionalIconSet.xlsx");

        byte[] bytes;
        using (ExcelDocument document = ExcelDocument.Create(workbookPath, "Conditional")) {
            ExcelSheet sheet = document.Sheets[0];
            sheet.Cell(1, 1, "Score");
            sheet.Cell(2, 1, 0);
            sheet.Cell(3, 1, 50);
            sheet.Cell(4, 1, 100);
            sheet.AddConditionalIconSet("A2:A4", IconSetValues.ThreeTrafficLights1, showValue: true, reverseIconOrder: false);

            ExcelConditionalFormattingInfo rule = Assert.Single(sheet.GetConditionalFormattingRules("A2:A4"));
            Assert.Equal("IconSet", rule.Type);
            Assert.Equal("ThreeTrafficLights1", rule.IconSet);
            Assert.True(rule.IconSetShowValue);
            Assert.False(rule.IconSetReverse);

            document.Save(false);

            bytes = document.SaveAsPdf(new ExcelPdfSaveOptions {
                IncludeSheetHeadings = false,
                HeaderRowCount = 1,
                PageSize = new PdfCore.PageSize(360, 240),
                Margins = PdfCore.PageMargins.Uniform(24)
            });
        }

        using PdfDocument pdf = PdfDocument.Open(new MemoryStream(bytes));
        string text = pdf.GetPage(1).Text;
        Assert.Contains("Score", text);
        Assert.Contains("100", text);

        string rawPdf = Encoding.ASCII.GetString(bytes);
        Assert.Contains("0.753 0.314 0.302 rg", rawPdf, StringComparison.Ordinal);
        Assert.Contains("1 0.753 0 rg", rawPdf, StringComparison.Ordinal);
        Assert.Contains("0.388 0.608 0.278 rg", rawPdf, StringComparison.Ordinal);
        Assert.Contains(" c ", rawPdf, StringComparison.Ordinal);
    }

    [Fact]
    public void SaveAsPdf_ExcelWorkbook_Maps_External_Cell_Hyperlinks() {
        const string linkUri = "https://github.com/EvotecIT/OfficeIMO";
        string workbookPath = Path.Combine(_directoryWithFiles, "ExcelPdfHyperlinks.xlsx");

        byte[] bytes;
        using (ExcelDocument document = ExcelDocument.Create(workbookPath, "Links")) {
            ExcelSheet sheet = document.Sheets[0];
            sheet.Cell(1, 1, "Name");
            sheet.SetHyperlink(2, 1, linkUri, display: "OfficeIMO");

            ExcelHyperlinkSnapshot hyperlink = Assert.Single(sheet.GetHyperlinks()).Value;
            Assert.True(hyperlink.IsExternal);
            Assert.Equal(linkUri, hyperlink.Target);

            document.Save(false);

            bytes = document.SaveAsPdf(new ExcelPdfSaveOptions {
                IncludeSheetHeadings = false,
                HeaderRowCount = 1,
                PageSize = new PdfCore.PageSize(360, 220),
                Margins = PdfCore.PageMargins.Uniform(24)
            });
        }

        using PdfDocument pdf = PdfDocument.Open(new MemoryStream(bytes));
        string text = pdf.GetPage(1).Text;
        Assert.Contains("OfficeIMO", text);

        PdfCore.PdfLogicalDocument logical = PdfCore.PdfLogicalDocument.Load(bytes);
        PdfCore.PdfLogicalLinkAnnotation link = Assert.Single(logical.GetLinksByUri(linkUri));
        Assert.Equal("OfficeIMO", link.Contents);

        string rawPdf = Encoding.ASCII.GetString(bytes);
        Assert.Contains("/Subtype /Link", rawPdf, StringComparison.Ordinal);
        Assert.Contains("/URI (https://github.com/EvotecIT/OfficeIMO)", rawPdf, StringComparison.Ordinal);
    }

    [Fact]
    public void SaveAsPdf_ExcelWorkbook_Maps_Internal_Cell_Hyperlinks_To_Cell_Destinations() {
        string workbookPath = Path.Combine(_directoryWithFiles, "ExcelPdfInternalHyperlinks.xlsx");

        byte[] bytes;
        byte[] summaryOnlyBytes;
        using (ExcelDocument document = ExcelDocument.Create(workbookPath)) {
            ExcelSheet summary = document.AddWorkSheet("Summary");
            summary.Cell(1, 1, "Name");
            summary.SetInternalLink(2, 1, "Details!B3", display: "Open Details B3");
            ExcelSheet details = document.AddWorkSheet("Details");
            details.Cell(1, 1, "Details Target");
            details.Cell(2, 1, "DestinationValue");
            details.Cell(3, 2, "CellSpecificTarget");

            ExcelHyperlinkSnapshot hyperlink = Assert.Single(summary.GetHyperlinks()).Value;
            Assert.False(hyperlink.IsExternal);
            Assert.Equal("'Details'!B3", hyperlink.Target);

            document.Save(false);

            bytes = document.SaveAsPdf(new ExcelPdfSaveOptions {
                IncludeSheetHeadings = true,
                HeaderRowCount = 1,
                PageSize = new PdfCore.PageSize(360, 220),
                Margins = PdfCore.PageMargins.Uniform(24)
            });

            summaryOnlyBytes = document.SaveAsPdf(new ExcelPdfSaveOptions {
                SheetNames = new[] { "Summary" },
                IncludeSheetHeadings = true,
                HeaderRowCount = 1,
                PageSize = new PdfCore.PageSize(360, 220),
                Margins = PdfCore.PageMargins.Uniform(24)
            });
        }

        PdfCore.PdfLogicalDocument logical = PdfCore.PdfLogicalDocument.Load(bytes);
        PdfCore.PdfNamedDestination destination = Assert.Single(logical.NamedDestinations, item => item.Name.EndsWith("-b3", StringComparison.Ordinal));
        PdfCore.PdfLogicalLinkAnnotation link = Assert.Single(logical.GetLinksByDestinationName(destination.Name));
        Assert.True(link.IsNamedDestinationLink);
        Assert.Equal("Open Details B3", link.Contents);
        Assert.Equal(destination.Name, link.DestinationName);
        Assert.Contains("details", destination.Name, StringComparison.Ordinal);

        string rawPdf = Encoding.ASCII.GetString(bytes);
        Assert.Contains("/Subtype /Link", rawPdf, StringComparison.Ordinal);
        Assert.Contains("/S /GoTo", rawPdf, StringComparison.Ordinal);

        PdfCore.PdfLogicalDocument summaryOnly = PdfCore.PdfLogicalDocument.Load(summaryOnlyBytes);
        Assert.DoesNotContain(summaryOnly.NamedDestinations, item => item.Name.IndexOf("details", StringComparison.Ordinal) >= 0);
        Assert.DoesNotContain(summaryOnly.Links, link => link.IsNamedDestinationLink);
    }

    [Fact]
    public void SaveAsPdf_ExcelWorkbook_Maps_SameSheet_Internal_Cell_Hyperlinks_To_Cell_Destinations() {
        string workbookPath = Path.Combine(_directoryWithFiles, "ExcelPdfSameSheetInternalHyperlinks.xlsx");

        byte[] bytes;
        using (ExcelDocument document = ExcelDocument.Create(workbookPath, "Links")) {
            ExcelSheet sheet = document.Sheets[0];
            sheet.Cell(1, 1, "Top Target");
            sheet.SetInternalLink(2, 1, "A1", display: "Back to Top");
            document.Save(false);

            bytes = document.SaveAsPdf(new ExcelPdfSaveOptions {
                IncludeSheetHeadings = true,
                HeaderRowCount = 1,
                PageSize = new PdfCore.PageSize(360, 220),
                Margins = PdfCore.PageMargins.Uniform(24)
            });
        }

        PdfCore.PdfLogicalDocument logical = PdfCore.PdfLogicalDocument.Load(bytes);
        PdfCore.PdfNamedDestination destination = Assert.Single(logical.NamedDestinations, item => item.Name.EndsWith("-a1", StringComparison.Ordinal));
        PdfCore.PdfLogicalLinkAnnotation link = Assert.Single(logical.GetLinksByDestinationName(destination.Name));
        Assert.Equal("Back to Top", link.Contents);
        Assert.Equal(destination.Name, link.DestinationName);
    }

    [Fact]
    public void SaveAsPdf_ExcelWorkbook_Drops_Internal_Cell_Hyperlinks_To_Unexported_Target_Cells() {
        string workbookPath = Path.Combine(_directoryWithFiles, "ExcelPdfInternalHyperlinkHiddenTarget.xlsx");

        byte[] bytes;
        using (ExcelDocument document = ExcelDocument.Create(workbookPath)) {
            ExcelSheet summary = document.AddWorkSheet("Summary");
            summary.Cell(1, 1, "Name");
            summary.SetInternalLink(2, 1, "Details!B200", display: "Open Details B200");
            ExcelSheet details = document.AddWorkSheet("Details");
            details.Cell(1, 1, "Details Header");
            details.Cell(2, 1, "Visible Detail");
            details.Cell(200, 2, "Hidden Target");
            document.SetPrintArea(details, "A1:B2");
            document.Save(false);

            bytes = document.SaveAsPdf(new ExcelPdfSaveOptions {
                IncludeSheetHeadings = true,
                HeaderRowCount = 1,
                PageSize = new PdfCore.PageSize(360, 220),
                Margins = PdfCore.PageMargins.Uniform(24)
            });
        }

        PdfCore.PdfLogicalDocument logical = PdfCore.PdfLogicalDocument.Load(bytes);
        Assert.Contains(logical.NamedDestinations, item => item.Name.IndexOf("details", StringComparison.Ordinal) >= 0);
        Assert.DoesNotContain(logical.NamedDestinations, item => item.Name.EndsWith("-b200", StringComparison.Ordinal));
        Assert.DoesNotContain(logical.Links, link => link.IsNamedDestinationLink && link.Contents == "Open Details B200");
    }

    [Fact]
    public void SaveAsPdf_ExcelWorkbook_Maps_Cell_Alignment_And_Borders() {
        string workbookPath = Path.Combine(_directoryWithFiles, "ExcelPdfCellAlignmentBorders.xlsx");

        byte[] bytes;
        using (ExcelDocument document = ExcelDocument.Create(workbookPath, "StyleLayout")) {
            ExcelSheet sheet = document.Sheets[0];
            sheet.Cell(1, 1, "Label");
            sheet.Cell(1, 2, "ZZ");
            sheet.Cell(2, 1, "Reference");
            sheet.Cell(2, 2, "LeftInColumn");
            sheet.CellAlign(1, 2, DocumentFormat.OpenXml.Spreadsheet.HorizontalAlignmentValues.Right);
            sheet.CellVerticalAlign(1, 2, DocumentFormat.OpenXml.Spreadsheet.VerticalAlignmentValues.Bottom);
            sheet.CellBorder(1, 2, DocumentFormat.OpenXml.Spreadsheet.BorderStyleValues.Medium, "445566");

            ExcelCellStyleSnapshot style = sheet.CellAt(1, 2).GetStyle();
            Assert.Equal("right", style.HorizontalAlignment);
            Assert.Equal("bottom", style.VerticalAlignment);
            Assert.NotNull(style.Border);
            Assert.Equal("medium", style.Border!.Left!.Style);
            Assert.Equal("FF445566", style.Border.Left.ColorArgb);

            document.Save(false);

            bytes = document.SaveAsPdf(new ExcelPdfSaveOptions {
                IncludeSheetHeadings = false,
                HeaderRowCount = 0,
                PageSize = new PdfCore.PageSize(360, 220),
                Margins = PdfCore.PageMargins.Uniform(24)
            });
        }

        using PdfDocument pdf = PdfDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(1);
        string text = page.Text;
        Assert.Contains("ZZ", text);
        Assert.Contains("LeftInColumn", text);

        double rightAlignedX = FindWordStartX(page, "ZZ");
        double sameColumnLeftX = FindWordStartX(page, "LeftInColumn");
        Assert.True(rightAlignedX > sameColumnLeftX + 70D, $"Expected right-aligned cell text to move toward the cell's right edge. Right x: {rightAlignedX:0.##}, left-reference x: {sameColumnLeftX:0.##}.");

        string rawPdf = Encoding.ASCII.GetString(bytes);
        Assert.Contains("0.267 0.333 0.4 RG", rawPdf, StringComparison.Ordinal);
        Assert.Contains("1.25 w", rawPdf, StringComparison.Ordinal);
    }

    [Fact]
    public void SaveAsPdf_ExcelWorkbook_Maps_Dashed_Cell_Border_Styles() {
        string workbookPath = Path.Combine(_directoryWithFiles, "ExcelPdfCellBorderDashStyles.xlsx");

        byte[] bytes;
        using (ExcelDocument document = ExcelDocument.Create(workbookPath, "BorderStyles")) {
            ExcelSheet sheet = document.Sheets[0];
            sheet.Cell(1, 1, "Dashed");
            sheet.Cell(1, 2, "Dotted");
            sheet.Cell(1, 3, "DashDot");
            sheet.CellBorder(1, 1, BorderStyleValues.Dashed, "123456");
            sheet.CellBorder(1, 2, BorderStyleValues.Dotted, "654321");
            sheet.CellBorder(1, 3, BorderStyleValues.MediumDashDot, "445566");

            Assert.Equal("dashed", sheet.CellAt(1, 1).GetStyle().Border!.Left!.Style);
            Assert.Equal("dotted", sheet.CellAt(1, 2).GetStyle().Border!.Left!.Style);
            Assert.Equal("mediumdashdot", sheet.CellAt(1, 3).GetStyle().Border!.Left!.Style);

            document.Save(false);

            bytes = document.SaveAsPdf(new ExcelPdfSaveOptions {
                IncludeSheetHeadings = false,
                HeaderRowCount = 0,
                PageSize = new PdfCore.PageSize(360, 220),
                Margins = PdfCore.PageMargins.Uniform(24)
            });
        }

        using (PdfDocument pdf = PdfDocument.Open(new MemoryStream(bytes))) {
            string text = pdf.GetPage(1).Text;
            Assert.Contains("Dashed", text);
            Assert.Contains("Dotted", text);
            Assert.Contains("DashDot", text);
        }

        string rawPdf = Encoding.ASCII.GetString(bytes);
        Assert.Contains("[1.5 0.75] 0 d", rawPdf, StringComparison.Ordinal);
        Assert.Contains("[0.5 0.75] 0 d", rawPdf, StringComparison.Ordinal);
        Assert.Contains("[3.75 1.875 1.25 1.875] 0 d", rawPdf, StringComparison.Ordinal);
        Assert.Contains("1 J", rawPdf, StringComparison.Ordinal);
    }

    [Fact]
    public void SaveAsPdf_ExcelWorkbook_Maps_Double_And_Diagonal_Cell_Borders() {
        string workbookPath = Path.Combine(_directoryWithFiles, "ExcelPdfCellBorderDoubleDiagonal.xlsx");

        byte[] bytes;
        using (ExcelDocument document = ExcelDocument.Create(workbookPath, "BorderStyles")) {
            ExcelSheet sheet = document.Sheets[0];
            sheet.Cell(1, 1, "Double");
            sheet.Cell(1, 2, "Diagonal");
            sheet.CellBorder(1, 1, BorderStyleValues.Double, "123456");
            sheet.CellDiagonalBorder(1, 2, BorderStyleValues.Double, "654321", diagonalUp: true, diagonalDown: true);

            ExcelCellStyleSnapshot doubleStyle = sheet.CellAt(1, 1).GetStyle();
            Assert.Equal("double", doubleStyle.Border!.Top!.Style);
            ExcelCellStyleSnapshot diagonalStyle = sheet.CellAt(1, 2).GetStyle();
            Assert.True(diagonalStyle.Border!.DiagonalUp);
            Assert.True(diagonalStyle.Border.DiagonalDown);
            Assert.Equal("double", diagonalStyle.Border.Diagonal!.Style);

            document.Save(false);

            bytes = document.SaveAsPdf(new ExcelPdfSaveOptions {
                IncludeSheetHeadings = false,
                HeaderRowCount = 0,
                PageSize = new PdfCore.PageSize(320, 220),
                Margins = PdfCore.PageMargins.Uniform(30)
            });
        }

        string rawPdf = Encoding.ASCII.GetString(bytes);
        Assert.Contains("0.071 0.204 0.337 RG", rawPdf, StringComparison.Ordinal);
        Assert.Contains("0.396 0.263 0.129 RG", rawPdf, StringComparison.Ordinal);
        Assert.True(rawPdf.Split(new[] { " S" }, StringSplitOptions.None).Length - 1 >= 10, "Expected Excel double and diagonal borders to emit multiple stroked lines.");
        Assert.True(rawPdf.Contains(" m ", StringComparison.Ordinal) && rawPdf.Contains(" l S", StringComparison.Ordinal), "Expected Excel diagonal borders to emit PDF line segments.");

        using (PdfDocument pdf = PdfDocument.Open(new MemoryStream(bytes))) {
            string text = pdf.GetPage(1).Text;
            Assert.Contains("Double", text);
            Assert.Contains("Diagonal", text);
        }
    }

    [Fact]
    public void SaveAsPdf_ExcelWorkbook_Maps_Common_Number_Formats() {
        string workbookPath = Path.Combine(_directoryWithFiles, "ExcelPdfNumberFormats.xlsx");

        byte[] bytes;
        using (ExcelDocument document = ExcelDocument.Create(workbookPath, "Formats")) {
            ExcelSheet sheet = document.Sheets[0];
            sheet.Cell(1, 1, "Kind");
            sheet.Cell(1, 2, "Value");
            sheet.Cell(2, 1, "Currency");
            sheet.CellAt(2, 2).SetValue(1234.5).Currency(2, CultureInfo.GetCultureInfo("en-US"));
            sheet.Cell(3, 1, "Percent");
            sheet.CellAt(3, 2).SetValue(0.257).Percent(1);
            sheet.Cell(4, 1, "Date");
            sheet.CellAt(4, 2).SetValue(new DateTime(2026, 1, 15)).Date("yyyy-mm-dd");
            sheet.Cell(5, 1, "Minutes");
            sheet.CellAt(5, 2).SetValue(new DateTime(2026, 1, 15, 0, 30, 5)).SetNumberFormat("mm:ss");
            sheet.Cell(6, 1, "Negative");
            sheet.CellAt(6, 2).SetValue(-1234).SetNumberFormat("#,##0;(#,##0)");

            ExcelCellStyleSnapshot currencyStyle = sheet.CellAt(2, 2).GetStyle();
            Assert.Equal("\"$\"#,##0.00", currencyStyle.NumberFormatCode);
            ExcelCellStyleSnapshot percentStyle = sheet.CellAt(3, 2).GetStyle();
            Assert.Equal("0.0%", percentStyle.NumberFormatCode);
            ExcelCellStyleSnapshot dateStyle = sheet.CellAt(4, 2).GetStyle();
            Assert.Equal("yyyy-mm-dd", dateStyle.NumberFormatCode);
            Assert.True(dateStyle.IsDateLike);

            document.Save(false);

            bytes = document.SaveAsPdf(new ExcelPdfSaveOptions {
                IncludeSheetHeadings = false,
                HeaderRowCount = 1,
                PageSize = new PdfCore.PageSize(420, 260),
                Margins = PdfCore.PageMargins.Uniform(24)
            });
        }

        using PdfDocument pdf = PdfDocument.Open(new MemoryStream(bytes));
        string text = pdf.GetPage(1).Text;
        Assert.Contains("$1,234.50", text);
        Assert.Contains("25.7%", text);
        Assert.Contains("2026-01-15", text);
        Assert.Contains("30:05", text);
        Assert.Contains("(1,234)", text);
        Assert.DoesNotContain("01:05", text);
        Assert.DoesNotContain("1234.5", text);
        Assert.DoesNotContain("0.257", text);
        Assert.DoesNotContain("-1,234", text);
    }

    [Fact]
    public void SaveAsPdf_ExcelWorkbook_Uses_Worksheet_Column_Widths_And_Print_Scale() {
        string workbookPath = Path.Combine(_directoryWithFiles, "ExcelPdfColumnWidths.xlsx");

        byte[] bytes;
        using (ExcelDocument document = ExcelDocument.Create(workbookPath, "Widths")) {
            ExcelSheet sheet = document.Sheets[0];
            sheet.Cell(1, 1, "ARef");
            sheet.Cell(1, 2, "WideColumn");
            sheet.Cell(1, 3, "Tail");
            sheet.SetColumnWidth(1, 8);
            sheet.SetColumnWidth(2, 32);
            sheet.SetColumnWidth(3, 8);
            sheet.SetPageSetup(scale: 50);

            IReadOnlyList<ExcelColumnSnapshot> columns = sheet.GetColumnDefinitions();
            Assert.Equal(3, columns.Count);
            Assert.Equal(32, columns[1].Width);
            Assert.True(columns[1].CustomWidth);
            Assert.Equal((uint)50, sheet.GetPageSetup().Scale);

            document.Save(false);

            bytes = document.SaveAsPdf(new ExcelPdfSaveOptions {
                IncludeSheetHeadings = false,
                HeaderRowCount = 0,
                PageSize = new PdfCore.PageSize(360, 220),
                Margins = PdfCore.PageMargins.Uniform(24)
            });
        }

        using PdfDocument pdf = PdfDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(1);
        string text = page.Text;
        Assert.Contains("ARef", text);
        Assert.Contains("WideColumn", text);
        Assert.Contains("Tail", text);

        double firstColumnX = page.Letters
            .Where(letter => !string.IsNullOrWhiteSpace(letter.Value))
            .Min(letter => letter.StartBaseLine.X);
        double wideColumnX = FindFirstLetterStartX(page, "W");
        double tailX = FindFirstLetterStartX(page, "T");
        Assert.True(tailX - wideColumnX > (wideColumnX - firstColumnX) * 2D, $"Expected worksheet column width proportions to make the middle column visibly wider. A: {firstColumnX:0.##}, B: {wideColumnX:0.##}, C: {tailX:0.##}.");
        Assert.True(tailX < 190D, $"Expected worksheet print scale to narrow the rendered table. Tail x: {tailX:0.##}.");
    }

    [Fact]
    public void SaveAsPdf_ExcelWorkbook_Uses_Worksheet_Row_Heights() {
        string workbookPath = Path.Combine(_directoryWithFiles, "ExcelPdfRowHeights.xlsx");

        byte[] bytes;
        using (ExcelDocument document = ExcelDocument.Create(workbookPath, "Heights")) {
            ExcelSheet sheet = document.Sheets[0];
            sheet.Cell(1, 1, "TopRow");
            sheet.Cell(2, 1, "TallRow");
            sheet.Cell(3, 1, "AfterTall");
            sheet.SetRowHeight(2, 60);

            ExcelRowSnapshot row = Assert.Single(sheet.GetRowDefinitions());
            Assert.Equal(2, row.Index);
            Assert.Equal(60, row.Height);
            Assert.True(row.CustomHeight);

            document.Save(false);

            bytes = document.SaveAsPdf(new ExcelPdfSaveOptions {
                IncludeSheetHeadings = false,
                HeaderRowCount = 0,
                PageSize = new PdfCore.PageSize(260, 260),
                Margins = PdfCore.PageMargins.Uniform(24)
            });
        }

        using PdfDocument pdf = PdfDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(1);
        string text = page.Text;
        Assert.Contains("TopRow", text);
        Assert.Contains("TallRow", text);
        Assert.Contains("AfterTall", text);

        double topY = FindWordStartY(page, "TopRow");
        double tallY = FindWordStartY(page, "TallRow");
        double afterY = FindWordStartY(page, "AfterTall");
        double defaultGap = topY - tallY;
        double customGap = tallY - afterY;
        Assert.True(customGap > defaultGap * 2D, $"Expected worksheet row height to create a visibly taller second PDF table row. Default gap: {defaultGap:0.##}, custom gap: {customGap:0.##}.");
    }

    [Fact]
    public void SaveAsPdf_ExcelWorkbook_Omits_Hidden_Rows_And_Columns() {
        string workbookPath = Path.Combine(_directoryWithFiles, "ExcelPdfHiddenRowsColumns.xlsx");

        byte[] bytes;
        using (ExcelDocument document = ExcelDocument.Create(workbookPath, "VisibleOnly")) {
            ExcelSheet sheet = document.Sheets[0];
            sheet.Cell(1, 1, "VisibleHeader");
            sheet.Cell(1, 2, "HiddenColumnValue");
            sheet.Cell(2, 1, "HiddenRowValue");
            sheet.Cell(3, 1, "VisibleTail");
            sheet.SetColumnHidden(2, true);
            sheet.SetRowHidden(2, true);

            ExcelColumnSnapshot column = Assert.Single(sheet.GetColumnDefinitions());
            Assert.Equal(2, column.StartIndex);
            Assert.Equal(2, column.EndIndex);
            Assert.True(column.Hidden);

            ExcelRowSnapshot row = Assert.Single(sheet.GetRowDefinitions());
            Assert.Equal(2, row.Index);
            Assert.True(row.Hidden);

            document.Save(false);

            bytes = document.SaveAsPdf(new ExcelPdfSaveOptions {
                IncludeSheetHeadings = false,
                HeaderRowCount = 0,
                PageSize = new PdfCore.PageSize(320, 220),
                Margins = PdfCore.PageMargins.Uniform(24)
            });
        }

        using PdfDocument pdf = PdfDocument.Open(new MemoryStream(bytes));
        string text = pdf.GetPage(1).Text;
        Assert.Contains("VisibleHeader", text);
        Assert.Contains("VisibleTail", text);
        Assert.DoesNotContain("HiddenColumnValue", text);
        Assert.DoesNotContain("HiddenRowValue", text);
    }

    [Fact]
    public void SaveAsPdf_ExcelWorkbook_Maps_Merged_Cells_To_Table_Spans() {
        string workbookPath = Path.Combine(_directoryWithFiles, "ExcelPdfMergedCells.xlsx");

        byte[] bytes;
        using (ExcelDocument document = ExcelDocument.Create(workbookPath, "Merged")) {
            ExcelSheet sheet = document.Sheets[0];
            sheet.Cell(1, 1, "MergedTitle");
            sheet.Cell(1, 3, "TailCell");
            sheet.Cell(2, 1, "ColumnA");
            sheet.Cell(2, 2, "ColumnB");
            sheet.Cell(2, 3, "ColumnC");
            sheet.MergeRange("A1:B1");

            ExcelMergedRangeSnapshot mergedRange = Assert.Single(sheet.GetMergedRanges());
            Assert.Equal("A1:B1", mergedRange.A1Range);
            Assert.Equal(1, mergedRange.StartRow);
            Assert.Equal(1, mergedRange.StartColumn);
            Assert.Equal(1, mergedRange.EndRow);
            Assert.Equal(2, mergedRange.EndColumn);

            document.Save(false);

            bytes = document.SaveAsPdf(new ExcelPdfSaveOptions {
                IncludeSheetHeadings = false,
                HeaderRowCount = 0,
                PageSize = new PdfCore.PageSize(360, 220),
                Margins = PdfCore.PageMargins.Uniform(24)
            });
        }

        using PdfDocument pdf = PdfDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(1);
        string text = page.Text;
        Assert.Contains("MergedTitle", text);
        Assert.Contains("TailCell", text);
        Assert.Contains("ColumnA", text);
        Assert.Contains("ColumnB", text);
        Assert.Contains("ColumnC", text);

        double mergedTitleX = FindWordStartX(page, "MergedTitle");
        double tailCellX = FindWordStartX(page, "TailCell");
        double columnBX = FindWordStartX(page, "ColumnB");
        double columnCX = FindWordStartX(page, "ColumnC");

        Assert.True(tailCellX > columnBX + 30D, $"Expected tail cell after A1:B1 merge to render in the third visual column. Tail x: {tailCellX:0.##}, ColumnB x: {columnBX:0.##}.");
        Assert.InRange(tailCellX, columnCX - 4D, columnCX + 4D);
        Assert.True(mergedTitleX < columnBX, $"Expected merged title to start in the first visual column. Title x: {mergedTitleX:0.##}, ColumnB x: {columnBX:0.##}.");
    }

    [Fact]
    public void SaveAsPdf_ExcelWorkbook_Exports_Worksheet_Chart_Snapshots() {
        string workbookPath = Path.Combine(_directoryWithFiles, "ExcelPdfCharts.xlsx");

        byte[] bytes;
        byte[] disabledBytes;
        using (ExcelDocument document = ExcelDocument.Create(workbookPath, "Charts")) {
            ExcelSheet sheet = document.Sheets[0];
            sheet.Cell(1, 1, "Category");
            sheet.Cell(1, 2, "Actual");
            sheet.Cell(1, 3, "Target");
            sheet.Cell(2, 1, "Jan");
            sheet.Cell(2, 2, 12);
            sheet.Cell(2, 3, 10);
            sheet.Cell(3, 1, "Feb");
            sheet.Cell(3, 2, 18);
            sheet.Cell(3, 3, 16);
            sheet.Cell(4, 1, "Mar");
            sheet.Cell(4, 2, 24);
            sheet.Cell(4, 3, 20);
            sheet.AddChartFromRange("A1:C4", row: 1, column: 5, widthPixels: 360, heightPixels: 220, type: ExcelChartType.ColumnClustered, title: "Revenue Chart");

            ExcelChart chart = Assert.Single(sheet.Charts);
            Assert.True(chart.TryGetSnapshot(out ExcelChartSnapshot snapshot));
            Assert.Equal("Revenue Chart", snapshot.Title);
            Assert.Equal(ExcelChartType.ColumnClustered, snapshot.ChartType);
            Assert.Equal(3, snapshot.Data.Categories.Count);
            Assert.Equal(2, snapshot.Data.Series.Count);

            document.Save(false);

            var options = new ExcelPdfSaveOptions {
                IncludeSheetHeadings = false,
                HeaderRowCount = 1,
                PageSize = new PdfCore.PageSize(480, 360),
                Margins = PdfCore.PageMargins.Uniform(24)
            };
            bytes = document.SaveAsPdf(options);
            disabledBytes = document.SaveAsPdf(new ExcelPdfSaveOptions {
                IncludeSheetHeadings = false,
                HeaderRowCount = 1,
                UseWorksheetCharts = false,
                PageSize = new PdfCore.PageSize(480, 360),
                Margins = PdfCore.PageMargins.Uniform(24)
            });
        }

        using PdfDocument pdf = PdfDocument.Open(new MemoryStream(bytes));
        string text = pdf.GetPage(1).Text;
        Assert.Contains("Revenue Chart", text);
        Assert.Contains("Actual", text);
        Assert.Contains("Target", text);

        string rawPdf = Encoding.ASCII.GetString(bytes);
        Assert.Contains("0.122 0.306 0.475 rg", rawPdf, StringComparison.Ordinal);
        Assert.Contains("0.184 0.435 0.243 rg", rawPdf, StringComparison.Ordinal);

        using PdfDocument disabledPdf = PdfDocument.Open(new MemoryStream(disabledBytes));
        Assert.DoesNotContain("Revenue Chart", disabledPdf.GetPage(1).Text);
    }

    [Fact]
    public void SaveAsPdf_ExcelWorkbook_Exports_Pie_And_Doughnut_Chart_Snapshots() {
        string workbookPath = Path.Combine(_directoryWithFiles, "ExcelPdfPieDoughnutCharts.xlsx");

        byte[] bytes;
        using (ExcelDocument document = ExcelDocument.Create(workbookPath, "Charts")) {
            ExcelSheet sheet = document.Sheets[0];
            sheet.Cell(1, 1, "Category");
            sheet.Cell(1, 2, "Control Share");
            sheet.Cell(2, 1, "Compliant");
            sheet.Cell(2, 2, 62);
            sheet.Cell(3, 1, "Partial");
            sheet.Cell(3, 2, 21);
            sheet.Cell(4, 1, "Non-compliant");
            sheet.Cell(4, 2, 11);
            sheet.Cell(5, 1, "Not assessed");
            sheet.Cell(5, 2, 6);
            sheet.AddChartFromRange("A1:B5", row: 1, column: 4, widthPixels: 280, heightPixels: 180, type: ExcelChartType.Pie, title: "Control Status Pie");
            sheet.AddChartFromRange("A1:B5", row: 12, column: 4, widthPixels: 280, heightPixels: 180, type: ExcelChartType.Doughnut, title: "Control Status Doughnut");

            List<ExcelChart> charts = sheet.Charts.ToList();
            Assert.Equal(2, charts.Count);
            Assert.All(charts, chart => Assert.True(chart.TryGetSnapshot(out _)));
            Assert.Equal(ExcelChartType.Pie, charts[0].ChartType);
            Assert.Equal(ExcelChartType.Doughnut, charts[1].ChartType);

            document.Save(false);

            bytes = document.SaveAsPdf(new ExcelPdfSaveOptions {
                IncludeSheetHeadings = false,
                HeaderRowCount = 1,
                PageSize = new PdfCore.PageSize(480, 520),
                Margins = PdfCore.PageMargins.Uniform(24)
            });
        }

        using PdfDocument pdf = PdfDocument.Open(new MemoryStream(bytes));
        string text = string.Concat(pdf.GetPages().Select(page => page.Text));
        Assert.Contains("Control Status Pie", text);
        Assert.Contains("Control Status Doughnut", text);
        Assert.Contains("Compliant", text);
        Assert.Contains("Non-compliant", text);

        string rawPdf = Encoding.ASCII.GetString(bytes);
        Assert.Contains("0.122 0.306 0.475 rg", rawPdf, StringComparison.Ordinal);
        Assert.Contains("0.184 0.435 0.243 rg", rawPdf, StringComparison.Ordinal);
        Assert.Contains("0.722 0.353 0.137 rg", rawPdf, StringComparison.Ordinal);
    }

    [Fact]
    public void SaveAsPdf_ExcelWorkbook_Exports_Area_Chart_Snapshots() {
        string workbookPath = Path.Combine(_directoryWithFiles, "ExcelPdfAreaCharts.xlsx");

        byte[] bytes;
        using (ExcelDocument document = ExcelDocument.Create(workbookPath, "Charts")) {
            ExcelSheet sheet = document.Sheets[0];
            sheet.Cell(1, 1, "Quarter");
            sheet.Cell(1, 2, "Services");
            sheet.Cell(1, 3, "Licenses");
            sheet.Cell(2, 1, "Q1");
            sheet.Cell(2, 2, 36);
            sheet.Cell(2, 3, 19);
            sheet.Cell(3, 1, "Q2");
            sheet.Cell(3, 2, 44);
            sheet.Cell(3, 3, 25);
            sheet.Cell(4, 1, "Q3");
            sheet.Cell(4, 2, 50);
            sheet.Cell(4, 3, 31);
            sheet.Cell(5, 1, "Q4");
            sheet.Cell(5, 2, 54);
            sheet.Cell(5, 3, 34);
            sheet.AddChartFromRange("A1:C5", row: 1, column: 5, widthPixels: 360, heightPixels: 220, type: ExcelChartType.Area, title: "Revenue Area");
            sheet.AddChartFromRange("A1:C5", row: 14, column: 5, widthPixels: 360, heightPixels: 220, type: ExcelChartType.AreaStacked100, title: "Revenue Mix Area");

            List<ExcelChart> charts = sheet.Charts.ToList();
            Assert.Equal(2, charts.Count);
            Assert.All(charts, chart => Assert.True(chart.TryGetSnapshot(out _)));
            Assert.Equal(ExcelChartType.Area, charts[0].ChartType);
            Assert.Equal(ExcelChartType.AreaStacked100, charts[1].ChartType);

            document.Save(false);

            bytes = document.SaveAsPdf(new ExcelPdfSaveOptions {
                IncludeSheetHeadings = false,
                HeaderRowCount = 1,
                PageSize = new PdfCore.PageSize(520, 620),
                Margins = PdfCore.PageMargins.Uniform(24)
            });
        }

        using PdfDocument pdf = PdfDocument.Open(new MemoryStream(bytes));
        string text = string.Concat(pdf.GetPages().Select(page => page.Text));
        Assert.Contains("Revenue Area", text);
        Assert.Contains("Revenue Mix Area", text);
        Assert.Contains("Services", text);
        Assert.Contains("Licenses", text);

        string rawPdf = Encoding.ASCII.GetString(bytes);
        Assert.Contains("0.122 0.306 0.475 rg", rawPdf, StringComparison.Ordinal);
        Assert.Contains("0.184 0.435 0.243 RG", rawPdf, StringComparison.Ordinal);
    }

    [Fact]
    public void SaveAsPdf_ExcelWorkbook_Exports_Scatter_Chart_Snapshots() {
        string workbookPath = Path.Combine(_directoryWithFiles, "ExcelPdfScatterCharts.xlsx");

        byte[] bytes;
        using (ExcelDocument document = ExcelDocument.Create(workbookPath, "Charts")) {
            ExcelSheet sheet = document.Sheets[0];
            var data = new ExcelChartData(
                new[] { "1", "2", "4", "8", "16" },
                new[] {
                    new ExcelChartSeries("Latency", new[] { 9D, 7D, 5.5D, 4.2D, 3.8D }, ExcelChartType.Scatter),
                    new ExcelChartSeries("Throughput", new[] { 2D, 3.5D, 6D, 7.5D, 9D }, ExcelChartType.Scatter)
                });

            sheet.AddChart(data, row: 1, column: 5, widthPixels: 360, heightPixels: 220, type: ExcelChartType.Scatter, title: "Scale Scatter");

            ExcelChart chart = Assert.Single(sheet.Charts);
            Assert.True(chart.TryGetSnapshot(out ExcelChartSnapshot snapshot));
            Assert.Equal(ExcelChartType.Scatter, snapshot.ChartType);
            Assert.Equal(5, snapshot.Data.Categories.Count);
            Assert.Equal(2, snapshot.Data.Series.Count);

            document.Save(false);

            bytes = document.SaveAsPdf(new ExcelPdfSaveOptions {
                IncludeSheetHeadings = false,
                HeaderRowCount = 1,
                PageSize = new PdfCore.PageSize(480, 360),
                Margins = PdfCore.PageMargins.Uniform(24)
            });
        }

        using PdfDocument pdf = PdfDocument.Open(new MemoryStream(bytes));
        string text = pdf.GetPage(1).Text;
        Assert.Contains("Scale Scatter", text);
        Assert.Contains("Latency", text);
        Assert.Contains("Throughput", text);

        string rawPdf = Encoding.ASCII.GetString(bytes);
        Assert.Contains("0.122 0.306 0.475 RG", rawPdf, StringComparison.Ordinal);
        Assert.Contains("0.184 0.435 0.243 RG", rawPdf, StringComparison.Ordinal);
    }

    [Fact]
    public void SaveAsPdf_ExcelWorkbook_Exports_Radar_Chart_Snapshots() {
        string workbookPath = Path.Combine(_directoryWithFiles, "ExcelPdfRadarCharts.xlsx");

        byte[] bytes;
        using (ExcelDocument document = ExcelDocument.Create(workbookPath, "Charts")) {
            ExcelSheet sheet = document.Sheets[0];
            var data = new ExcelChartData(
                new[] { "Quality", "Speed", "Cost", "Coverage", "Risk" },
                new[] {
                    new ExcelChartSeries("Current", new[] { 7D, 6D, 5D, 8D, 4D }, ExcelChartType.Radar),
                    new ExcelChartSeries("Target", new[] { 9D, 8D, 7D, 9D, 6D }, ExcelChartType.Radar)
                });

            sheet.AddChart(data, row: 1, column: 5, widthPixels: 360, heightPixels: 220, type: ExcelChartType.Radar, title: "Capability Radar");

            ExcelChart chart = Assert.Single(sheet.Charts);
            Assert.True(chart.TryGetSnapshot(out ExcelChartSnapshot snapshot));
            Assert.Equal(ExcelChartType.Radar, snapshot.ChartType);
            Assert.Equal(5, snapshot.Data.Categories.Count);
            Assert.Equal(2, snapshot.Data.Series.Count);

            document.Save(false);

            bytes = document.SaveAsPdf(new ExcelPdfSaveOptions {
                IncludeSheetHeadings = false,
                HeaderRowCount = 1,
                PageSize = new PdfCore.PageSize(480, 360),
                Margins = PdfCore.PageMargins.Uniform(24)
            });
        }

        using PdfDocument pdf = PdfDocument.Open(new MemoryStream(bytes));
        string text = pdf.GetPage(1).Text;
        Assert.Contains("Capability Radar", text);
        Assert.Contains("Current", text);
        Assert.Contains("Target", text);

        string rawPdf = Encoding.ASCII.GetString(bytes);
        Assert.Contains("0.122 0.306 0.475 rg", rawPdf, StringComparison.Ordinal);
        Assert.Contains("0.184 0.435 0.243 rg", rawPdf, StringComparison.Ordinal);
    }

    [Fact]
    public void SaveAsPdf_ExcelWorkbook_Rejects_Invalid_Options() {
        Assert.Throws<ArgumentOutOfRangeException>(() => new ExcelPdfSaveOptions { HeaderRowCount = -1 });
        Assert.Throws<ArgumentOutOfRangeException>(() => new ExcelPdfSaveOptions { MaxRowsPerSheet = 0 });
    }

    [Fact]
    public void SaveAsPdf_ExcelWorkbook_Reports_Unsupported_Export_Features() {
        string workbookPath = Path.Combine(_directoryWithFiles, "ExcelPdfUnsupportedFeatureWarnings.xlsx");

        var options = new ExcelPdfSaveOptions {
            IncludeSheetHeadings = false,
            HeaderRowCount = 1,
            MaxRowsPerSheet = 2,
            PageSize = new PdfCore.PageSize(460, 320),
            Margins = PdfCore.PageMargins.Uniform(24)
        };

        byte[] bytes;
        using (ExcelDocument document = ExcelDocument.Create(workbookPath, "Warnings")) {
            ExcelSheet sheet = document.Sheets[0];
            sheet.Cell(1, 1, "Name");
            sheet.Cell(1, 2, "Value");
            sheet.Cell(2, 1, "Alpha");
            sheet.Cell(2, 2, 10);
            sheet.Cell(3, 1, "Beta");
            sheet.Cell(3, 2, 20);
            sheet.SetHeaderFooter(
                headerCenter: "&U&\"Arial,Bold\"&14&KFF0000Styled &D &T &A",
                footerRight: "Page &P of &N");
            sheet.AddChartFromRange("A1:B3", row: 1, column: 4, widthPixels: 320, heightPixels: 180, type: ExcelChartType.Surface, title: "Unsupported Surface Chart");

            ExcelChart chart = Assert.Single(sheet.Charts);
            Assert.True(chart.TryGetSnapshot(out ExcelChartSnapshot snapshot));
            Assert.Equal(ExcelChartType.Surface, snapshot.ChartType);

            document.Save(false);

            bytes = document.SaveAsPdf(options);
        }

        using (PdfDocument pdf = PdfDocument.Open(new MemoryStream(bytes))) {
            string text = pdf.GetPage(1).Text;
            Assert.Contains("Styled", text);
            Assert.Contains(DateTime.Now.ToString("d", CultureInfo.CurrentCulture), text);
            Assert.Contains("Warnings", text);
            Assert.Contains("Page 1 of", text);
            Assert.Contains("Alpha", text);
            Assert.DoesNotContain("Beta", text);
            Assert.DoesNotContain("Unsupported Surface Chart", text);
        }

        Assert.Contains(options.Warnings, warning => warning.SheetName == "Warnings" && warning.Feature == "WorksheetHeaderFooterFormatting");
        Assert.Contains(options.Warnings, warning => warning.SheetName == "Warnings" && warning.Feature == "WorksheetRows");
        Assert.Contains(options.Warnings, warning => warning.SheetName == "Warnings" && warning.Feature == "WorksheetChart" && warning.Message.Contains("Surface", StringComparison.Ordinal));
        Assert.All(options.Warnings, warning => Assert.Contains("Warnings", warning.ToString(), StringComparison.Ordinal));
    }

    private static double FindWordStartX(UglyToad.PdfPig.Content.Page page, string word) {
        var lines = page.Letters
            .Where(letter => !string.IsNullOrWhiteSpace(letter.Value))
            .GroupBy(letter => Math.Round(letter.StartBaseLine.Y, 1));

        foreach (var line in lines) {
            var ordered = line.OrderBy(letter => letter.StartBaseLine.X).ToList();
            string text = string.Concat(ordered.Select(letter => letter.Value));
            int index = text.IndexOf(word, StringComparison.Ordinal);
            if (index >= 0) {
                return ordered[index].StartBaseLine.X;
            }
        }

        throw new InvalidOperationException("Could not find word '" + word + "' in rendered PDF text.");
    }

    private static double FindWordStartY(UglyToad.PdfPig.Content.Page page, string word) {
        var lines = page.Letters
            .Where(letter => !string.IsNullOrWhiteSpace(letter.Value))
            .GroupBy(letter => Math.Round(letter.StartBaseLine.Y, 1));

        foreach (var line in lines) {
            var ordered = line.OrderBy(letter => letter.StartBaseLine.X).ToList();
            string text = string.Concat(ordered.Select(letter => letter.Value));
            if (text.IndexOf(word, StringComparison.Ordinal) >= 0) {
                return ordered[0].StartBaseLine.Y;
            }
        }

        throw new InvalidOperationException("Could not find word '" + word + "' in rendered PDF text.");
    }

    private static double FindFirstLetterStartX(UglyToad.PdfPig.Content.Page page, string letter) {
        double x = page.Letters
            .Where(pdfLetter => string.Equals(pdfLetter.Value, letter, StringComparison.Ordinal))
            .Select(pdfLetter => pdfLetter.StartBaseLine.X)
            .DefaultIfEmpty(double.NaN)
            .First();

        if (double.IsNaN(x)) {
            throw new InvalidOperationException("Could not find letter '" + letter + "' in rendered PDF text.");
        }

        return x;
    }

    private static byte[] CreateMinimalRgbPng() {
        return new byte[] {
            137, 80, 78, 71, 13, 10, 26, 10,
            0, 0, 0, 13,
            73, 72, 68, 82,
            0, 0, 0, 1,
            0, 0, 0, 1,
            8, 2, 0, 0, 0,
            0, 0, 0, 0,
            0, 0, 0, 12,
            73, 68, 65, 84,
            0x78, 0x9C, 0x63, 0xF8, 0xCF, 0xC0, 0x00, 0x00, 0x03, 0x01, 0x01, 0x00,
            0, 0, 0, 0,
            0, 0, 0, 0,
            73, 69, 78, 68,
            0, 0, 0, 0
        };
    }

    private static string NormalizePdfTextSpaces(string text) =>
        text.Replace('\u00A0', ' ').Replace('\u202F', ' ');
}
