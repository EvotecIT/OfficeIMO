using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
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
    public void SaveAsPdf_ExcelWorkbook_ExportsUsedRangeWithoutInjectingSheetHeading() {
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
            document.Save();

            document.SaveAsPdf(pdfPath);
        }

        Assert.True(File.Exists(pdfPath));
        using PdfPigDocument pdf = PdfPigDocument.Open(pdfPath);
        string text = pdf.GetPage(1).Text;
        Assert.DoesNotContain("Sales", text);
        Assert.Contains("Product", text);
        Assert.Contains("Licenses", text);
        Assert.Contains("1250.5", text);
    }

    [Fact]
    public void SaveAsPdf_ExcelWorkbook_PreservesExplicitMappedDefaultFontFamily() {
        string workbookPath = Path.Combine(_directoryWithFiles, "ExcelPdfExplicitSerif.xlsx");

        byte[] bytes;
        using (ExcelDocument document = ExcelDocument.Create(workbookPath, "Serif")) {
            ExcelSheet sheet = document.Sheets[0];
            sheet.Cell(1, 1, "Title");
            sheet.Cell(2, 1, "Explicit serif default");
            document.Save();

            bytes = document.ToPdf(new ExcelPdfSaveOptions {
                FontFamily = "serif",
                IncludeSheetHeadings = false
            });
        }

        string raw = Encoding.ASCII.GetString(bytes);
        Assert.Contains("/BaseFont /Times", raw, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void ToPdfDocument_ExcelWorkbook_Uses_Configured_Default_Table_Style() {
        string workbookPath = Path.Combine(_directoryWithFiles, "ExcelPdfConfiguredDefaultTableStyle.xlsx");
        var configuredStyle = new PdfCore.PdfTableStyle {
            HeaderRowCount = 0,
            RepeatHeaderRowCount = 0,
            CellPaddingX = 9,
            CellPaddingY = 7,
            BorderColor = null,
            HeaderFill = PdfCore.PdfColor.FromRgb(10, 20, 30),
            HeaderTextColor = PdfCore.PdfColor.FromRgb(240, 245, 250),
            RowStripeFill = null,
            FontSize = 10.5D,
            SpacingAfter = 12D
        };

        PdfCore.PdfDocument pdfDocument;
        using (ExcelDocument document = ExcelDocument.Create(workbookPath, "Styled")) {
            ExcelSheet sheet = document.Sheets[0];
            sheet.Cell(1, 1, "Product");
            sheet.Cell(1, 2, "Amount");
            sheet.Cell(2, 1, "Licenses");
            sheet.Cell(2, 2, 1250.5);
            document.Save();

            pdfDocument = document.ToPdfDocument(new ExcelPdfSaveOptions {
                IncludeSheetHeadings = false,
                HeaderRowCount = 1,
                PdfOptions = new PdfCore.PdfOptions {
                    DefaultTableStyle = configuredStyle
                }
            });
        }

        PdfCore.PageBlock page = Assert.IsType<PdfCore.PageBlock>(Assert.Single(pdfDocument.Blocks));
        PdfCore.TableBlock table = Assert.Single(page.Blocks.OfType<PdfCore.TableBlock>());
        Assert.NotNull(table.Style);
        PdfCore.PdfTableStyle style = table.Style!;

        Assert.Equal(1, style.HeaderRowCount);
        Assert.Equal(1, style.RepeatHeaderRowCount);
        Assert.Equal(9, style.CellPaddingX);
        Assert.Equal(7, style.CellPaddingY);
        Assert.Null(style.BorderColor);
        Assert.Equal(PdfCore.PdfColor.FromRgb(10, 20, 30), style.HeaderFill);
        Assert.Equal(PdfCore.PdfColor.FromRgb(240, 245, 250), style.HeaderTextColor);
        Assert.Null(style.RowStripeFill);
        Assert.Equal(10.5D, style.FontSize);
        Assert.Equal(12D, style.SpacingAfter);

        Assert.Equal(0, configuredStyle.HeaderRowCount);
        Assert.Equal(0, configuredStyle.RepeatHeaderRowCount);
        Assert.Null(configuredStyle.CellFills);
        Assert.Null(configuredStyle.ColumnWidthWeights);
    }

    [Fact]
    public void ToPdfDocument_ExcelWorkbook_Keeps_Excel_Default_Table_Style_For_Unrelated_PdfOptions() {
        string workbookPath = Path.Combine(_directoryWithFiles, "ExcelPdfUnrelatedPdfOptions.xlsx");

        PdfCore.PdfDocument pdfDocument;
        using (ExcelDocument document = ExcelDocument.Create(workbookPath, "Styled")) {
            ExcelSheet sheet = document.Sheets[0];
            sheet.Cell(1, 1, "Product");
            sheet.Cell(2, 1, "Licenses");
            document.Save();

            pdfDocument = document.ToPdfDocument(new ExcelPdfSaveOptions {
                IncludeSheetHeadings = false,
                HeaderRowCount = 1,
                PdfOptions = new PdfCore.PdfOptions {
                    DefaultFontSize = 9D
                }
            });
        }

        PdfCore.PageBlock page = Assert.IsType<PdfCore.PageBlock>(Assert.Single(pdfDocument.Blocks));
        PdfCore.TableBlock table = Assert.Single(page.Blocks.OfType<PdfCore.TableBlock>());
        PdfCore.PdfTableStyle style = table.Style!;

        Assert.Equal(4D, style.CellPaddingX);
        Assert.Equal(3D, style.CellPaddingY);
        Assert.Equal(PdfCore.PdfColor.FromRgb(230, 238, 247), style.HeaderFill);
        Assert.Equal(PdfCore.PdfColor.FromRgb(31, 78, 121), style.HeaderTextColor);
        Assert.Equal(PdfCore.PdfColor.FromRgb(248, 250, 252), style.RowStripeFill);
    }

    [Fact]
    public void SaveAsPdf_ExcelWorkbook_Respects_Selected_Sheets() {
        string workbookPath = Path.Combine(_directoryWithFiles, "ExcelPdfSelectedSheets.xlsx");

        byte[] bytes;
        using (ExcelDocument document = ExcelDocument.Create(workbookPath)) {
            ExcelSheet summary = document.AddWorksheet("Summary");
            summary.Cell(1, 1, "Metric");
            summary.Cell(2, 1, "SelectedValue");
            summary.SetHeaderFooter(headerCenter: "Selected Header &A");
            ExcelSheet internalSheet = document.AddWorksheet("Internal");
            internalSheet.Cell(1, 1, "HiddenValue");
            document.Save();

            bytes = document.ToPdf(new ExcelPdfSaveOptions {
                SheetNames = new[] { "summary" }
            });
        }

        using PdfPigDocument pdf = PdfPigDocument.Open(new MemoryStream(bytes));
        string text = pdf.GetPage(1).Text;
        Assert.Contains("Summary", text);
        Assert.Contains("Selected Header Summary", text);
        Assert.Contains("SelectedValue", text);
        Assert.DoesNotContain("HiddenValue", text);
    }


    [Fact]
    public void SaveAsPdf_ExcelWorkbook_Omits_Hidden_Sheets_By_Default() {
        string workbookPath = Path.Combine(_directoryWithFiles, "ExcelPdfHiddenSheets.xlsx");

        byte[] visibleBytes;
        byte[] explicitHiddenBytes;
        using (ExcelDocument document = ExcelDocument.Create(workbookPath)) {
            ExcelSheet visible = document.AddWorksheet("Visible");
            visible.Cell(1, 1, "VisibleSheetValue");
            ExcelSheet hidden = document.AddWorksheet("Hidden");
            hidden.Cell(1, 1, "HiddenSheetValue");
            hidden.SetHidden(true);
            Assert.True(hidden.Hidden);
            document.Save();

            visibleBytes = document.ToPdf(new ExcelPdfSaveOptions {
                IncludeSheetHeadings = false,
                HeaderRowCount = 0
            });

            explicitHiddenBytes = document.ToPdf(new ExcelPdfSaveOptions {
                SheetNames = new[] { "Hidden" },
                IncludeSheetHeadings = false,
                HeaderRowCount = 0
            });
        }

        using PdfPigDocument visiblePdf = PdfPigDocument.Open(new MemoryStream(visibleBytes));
        string visibleText = visiblePdf.GetPage(1).Text;
        Assert.Contains("VisibleSheetValue", visibleText);
        Assert.DoesNotContain("HiddenSheetValue", visibleText);

        using PdfPigDocument explicitHiddenPdf = PdfPigDocument.Open(new MemoryStream(explicitHiddenBytes));
        string hiddenText = explicitHiddenPdf.GetPage(1).Text;
        Assert.Contains("HiddenSheetValue", hiddenText);
    }

}
