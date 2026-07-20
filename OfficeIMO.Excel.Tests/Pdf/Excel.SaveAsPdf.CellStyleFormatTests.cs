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
    public void SaveAsPdf_ExcelWorkbook_Maps_Basic_Cell_Font_And_Fill_Styles() {
        string workbookPath = Path.Combine(_directoryWithFiles, "ExcelPdfCellStyles.xlsx");

        byte[] bytes;
        using (ExcelDocument document = ExcelDocument.Create(workbookPath, "Styled")) {
            ExcelSheet sheet = document.Sheets[0];
            sheet.CellAt(1, 1)
                .SetValue("StyledCell")
                .SetFontName("Consolas")
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
            Assert.Equal("Consolas", style.FontName);
            Assert.Equal("112233", style.FontColorHex);
            Assert.Equal("DDEEFF", style.FillColorHex);

            document.Save();

            bytes = document.ToPdf(new ExcelPdfSaveOptions {
                IncludeSheetHeadings = false,
                HeaderRowCount = 0,
                PageSize = new PdfCore.PageSize(360, 220),
                Margins = PdfCore.PageMargins.Uniform(24)
            });
        }

        using PdfPigDocument pdf = PdfPigDocument.Open(new MemoryStream(bytes));
        string text = pdf.GetPage(1).Text;
        Assert.Contains("StyledCell", text);
        Assert.Contains("PlainCell", text);

        string rawPdf = Encoding.ASCII.GetString(bytes);
        AssertRawPdfContainsAnyBaseFont(rawPdf, "Courier", "Consolas", "LiberationMono", "DejaVuSansMono");
        Assert.Contains("0.067 0.133 0.2 rg", rawPdf, StringComparison.Ordinal);
        Assert.Contains("0.867 0.933 1 rg", rawPdf, StringComparison.Ordinal);
    }

    [Fact]
    public void SaveAsPdf_ExcelWorkbook_Preserves_Helvetica_Cell_Font_When_Default_Font_Changes() {
        string workbookPath = Path.Combine(_directoryWithFiles, "ExcelPdfHelveticaCellWithTimesDefault.xlsx");

        byte[] bytes;
        using (ExcelDocument document = ExcelDocument.Create(workbookPath, "Fonts")) {
            ExcelSheet sheet = document.Sheets[0];
            sheet.CellAt(1, 1).SetValue("ArialCell").SetFontName("Arial");
            sheet.CellAt(1, 2).SetValue("PlainCell");
            document.Save();

            bytes = document.ToPdf(new ExcelPdfSaveOptions {
                FontFamily = "Times New Roman",
                IncludeSheetHeadings = false,
                HeaderRowCount = 0,
                PageSize = new PdfCore.PageSize(360, 220),
                Margins = PdfCore.PageMargins.Uniform(24)
            });
        }

        using PdfPigDocument pdf = PdfPigDocument.Open(new MemoryStream(bytes));
        string text = pdf.GetPage(1).Text;
        Assert.Contains("ArialCell", text);
        Assert.Contains("PlainCell", text);

        string rawPdf = Encoding.ASCII.GetString(bytes);
        AssertRawPdfContainsAnyBaseFont(rawPdf, "Helvetica", "Arial", "Calibri", "LiberationSans", "Liberation Sans");
        AssertRawPdfContainsAnyBaseFont(rawPdf, "Times");
    }

    [Fact]
    public void SaveAsPdf_ExcelWorkbook_PreservesDistinctSerifCellFontsWithoutSlotCollisions() {
        string workbookPath = Path.Combine(_directoryWithFiles, "ExcelPdfSameFamilyFontSlot.xlsx");

        byte[] bytes;
        using (ExcelDocument document = ExcelDocument.Create(workbookPath, "Fonts")) {
            ExcelSheet sheet = document.Sheets[0];
            sheet.CellAt(1, 1).SetValue("FirstSerif").SetFontName("Times New Roman");
            sheet.CellAt(1, 2).SetValue("SecondSerif").SetFontName("Georgia");
            document.Save();

            bytes = document.ToPdf(new ExcelPdfSaveOptions {
                IncludeSheetHeadings = false,
                HeaderRowCount = 0,
                PageSize = new PdfCore.PageSize(360, 220),
                Margins = PdfCore.PageMargins.Uniform(24)
            });
        }

        using PdfPigDocument pdf = PdfPigDocument.Open(new MemoryStream(bytes));
        string text = pdf.GetPage(1).Text;
        Assert.Contains("FirstSerif", text);
        Assert.Contains("SecondSerif", text);

        string rawPdf = Encoding.ASCII.GetString(bytes);
        AssertRawPdfContainsAnyBaseFont(rawPdf, "Times");
        AssertRawPdfContainsAnyBaseFont(rawPdf, "Georgia");
    }

    [Fact]
    public void SaveAsPdf_ExcelWorkbook_Reports_Unavailable_Cell_Font_Substitution() {
        const string unavailableFamily = "OfficeIMO Missing Font 7F0C9D";
        string workbookPath = Path.Combine(_directoryWithFiles, "ExcelPdfUnavailableCellFont.xlsx");
        PdfCore.PdfDocumentConversionResult result;
        using (ExcelDocument document = ExcelDocument.Create(workbookPath, "Fonts")) {
            document.Sheets[0]
                .CellAt(1, 1)
                .SetValue("Unavailable font marker")
                .SetFontName(unavailableFamily);
            document.Save();

            result = document.ToPdfDocumentResult(new ExcelPdfSaveOptions {
                IncludeSheetHeadings = false,
                HeaderRowCount = 0,
                ResourcePolicy = PdfCore.PdfResourcePolicy.CreateTrustedHost()
            });
        }

        PdfCore.PdfConversionWarning warning = Assert.Single(
            result.Warnings,
            item => item.Code == "WorksheetFontFamilySubstituted");
        Assert.Equal("Fonts", warning.Source);
        Assert.Throws<InvalidOperationException>(() => result.Report.RequireNoLoss());
    }

    [Fact]
    public void SaveAsPdf_ExcelWorkbook_Reports_Explicit_Font_Substitution_When_Host_Fonts_Are_Disabled() {
        string workbookPath = Path.Combine(_directoryWithFiles, "ExcelPdfPortableCellFont.xlsx");
        PdfCore.PdfDocumentConversionResult result;
        using (ExcelDocument document = ExcelDocument.Create(workbookPath, "Fonts")) {
            document.Sheets[0]
                .CellAt(1, 1)
                .SetValue("Portable font marker")
                .SetFontName("Arial");
            document.Save();

            result = document.ToPdfDocumentResult(new ExcelPdfSaveOptions {
                IncludeSheetHeadings = false,
                HeaderRowCount = 0,
                ResourcePolicy = PdfCore.PdfResourcePolicy.CreatePortableDeterministic()
            });
        }

        PdfCore.PdfConversionWarning warning = Assert.Single(
            result.Warnings,
            item => item.Code == "WorksheetFontFamilySubstituted");
        Assert.Equal("Arial", warning.Details["fontFamily"]);
        Assert.Equal("Helvetica", warning.Details["fallbackSlot"]);
    }

    [Fact]
    public void SaveAsPdf_ExcelWorkbook_PreservesConfiguredDefaultFontSlot() {
        string workbookPath = Path.Combine(_directoryWithFiles, "ExcelPdfDefaultFontSlot.xlsx");

        byte[] bytes;
        using (ExcelDocument document = ExcelDocument.Create(workbookPath, "Fonts")) {
            ExcelSheet sheet = document.Sheets[0];
            sheet.CellAt(1, 1).SetValue("StyledSerif").SetFontName("Georgia");
            sheet.CellAt(1, 2).SetValue("DefaultSerif");
            document.Save();

            bytes = document.ToPdf(new ExcelPdfSaveOptions {
                FontFamily = "Times New Roman",
                IncludeSheetHeadings = false,
                HeaderRowCount = 0,
                PageSize = new PdfCore.PageSize(360, 220),
                Margins = PdfCore.PageMargins.Uniform(24)
            });
        }

        using PdfPigDocument pdf = PdfPigDocument.Open(new MemoryStream(bytes));
        string text = pdf.GetPage(1).Text;
        Assert.Contains("StyledSerif", text);
        Assert.Contains("DefaultSerif", text);

        string rawPdf = Encoding.ASCII.GetString(bytes);
        AssertRawPdfContainsAnyBaseFont(rawPdf, "Times");
        AssertRawPdfContainsAnyBaseFont(rawPdf, "Georgia");
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

            document.Save();

            bytes = document.ToPdf(new ExcelPdfSaveOptions {
                IncludeSheetHeadings = false,
                HeaderRowCount = 1,
                PageSize = new PdfCore.PageSize(360, 240),
                Margins = PdfCore.PageMargins.Uniform(24)
            });
        }

        using PdfPigDocument pdf = PdfPigDocument.Open(new MemoryStream(bytes));
        string text = pdf.GetPage(1).Text;
        Assert.Contains("Score", text);
        Assert.Contains("100", text);

        string rawPdf = Encoding.ASCII.GetString(bytes);
        Assert.Contains("1 0 0 rg", rawPdf, StringComparison.Ordinal);
        Assert.Contains("0.502 0.502 0 rg", rawPdf, StringComparison.Ordinal);
        Assert.Contains("0 1 0 rg", rawPdf, StringComparison.Ordinal);
    }

    private static void AssertRawPdfContainsAnyBaseFont(string rawPdf, params string[] fontNameParts) {
        string[] baseFonts = Regex.Matches(rawPdf, @"/BaseFont /([^\s/<>\[\]()]+)")
            .Cast<Match>()
            .Select(match => match.Groups[1].Value)
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToArray();
        Assert.True(
            fontNameParts.Any(fontNamePart => rawPdf.Contains("/BaseFont /" + fontNamePart, StringComparison.OrdinalIgnoreCase)),
            "Expected raw PDF to contain one of these BaseFont names: " + string.Join(", ", fontNameParts) + ". Actual BaseFont names: " + string.Join(", ", baseFonts));
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

            document.Save();

            bytes = document.ToPdf(new ExcelPdfSaveOptions {
                IncludeSheetHeadings = false,
                HeaderRowCount = 1,
                PageSize = new PdfCore.PageSize(360, 240),
                Margins = PdfCore.PageMargins.Uniform(24)
            });
        }

        using PdfPigDocument pdf = PdfPigDocument.Open(new MemoryStream(bytes));
        string text = pdf.GetPage(1).Text;
        Assert.Contains("Score", text);
        Assert.Contains("100", text);

        string rawPdf = Encoding.ASCII.GetString(bytes);
        int barFillCount = rawPdf.Split(new[] { "0.357 0.608 0.835 rg" }, StringSplitOptions.None).Length - 1;

        Assert.Equal(2, barFillCount);
        Assert.Contains(" re f", rawPdf, StringComparison.Ordinal);
    }

    [Fact]
    public void SaveAsPdf_ExcelWorkbook_Preserves_Negative_Conditional_DataBars() {
        string workbookPath = Path.Combine(_directoryWithFiles, "ExcelPdfConditionalNegativeDataBar.xlsx");

        byte[] bytes;
        using (ExcelDocument document = ExcelDocument.Create(workbookPath, "Conditional")) {
            ExcelSheet sheet = document.Sheets[0];
            sheet.Cell(1, 1, "Delta");
            sheet.Cell(2, 1, -100);
            sheet.Cell(3, 1, 0);
            sheet.Cell(4, 1, 100);
            sheet.AddConditionalDataBar("A2:A4", "FF5B9BD5");
            document.Save();

            bytes = document.ToPdf(new ExcelPdfSaveOptions {
                IncludeSheetHeadings = false,
                HeaderRowCount = 1,
                PageSize = new PdfCore.PageSize(360, 240),
                Margins = PdfCore.PageMargins.Uniform(24)
            });
        }

        string rawPdf = Encoding.ASCII.GetString(bytes);
        MatchCollection barRects = Regex.Matches(rawPdf, @"0\.357 0\.608 0\.835 rg\s+(?<x>-?\d+(?:\.\d+)?) (?<y>-?\d+(?:\.\d+)?) (?<width>-?\d+(?:\.\d+)?) (?<height>-?\d+(?:\.\d+)?) re f");

        Assert.Equal(2, barRects.Count);
        double firstX = double.Parse(barRects[0].Groups["x"].Value, CultureInfo.InvariantCulture);
        double secondX = double.Parse(barRects[1].Groups["x"].Value, CultureInfo.InvariantCulture);
        double firstWidth = double.Parse(barRects[0].Groups["width"].Value, CultureInfo.InvariantCulture);
        double secondWidth = double.Parse(barRects[1].Groups["width"].Value, CultureInfo.InvariantCulture);

        Assert.True(firstX < secondX);
        Assert.InRange(Math.Abs(firstWidth - secondWidth), 0D, 1D);
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

            document.Save();

            bytes = document.ToPdf(new ExcelPdfSaveOptions {
                IncludeSheetHeadings = false,
                HeaderRowCount = 1,
                PageSize = new PdfCore.PageSize(360, 240),
                Margins = PdfCore.PageMargins.Uniform(24)
            });
        }

        using PdfPigDocument pdf = PdfPigDocument.Open(new MemoryStream(bytes));
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

            document.Save();

            bytes = document.ToPdf(new ExcelPdfSaveOptions {
                IncludeSheetHeadings = false,
                HeaderRowCount = 0,
                PageSize = new PdfCore.PageSize(360, 220),
                Margins = PdfCore.PageMargins.Uniform(24)
            });
        }

        using PdfPigDocument pdf = PdfPigDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(1);
        string text = page.Text;
        Assert.Contains("ZZ", text);
        Assert.Contains("LeftInColumn", text);

        double rightAlignedX = FindWordStartX(page, "ZZ");
        double sameColumnLeftX = FindWordStartX(page, "LeftInColumn");
        Assert.True(rightAlignedX > sameColumnLeftX + 20D, $"Expected right-aligned cell text to move toward the authored worksheet column's right edge. Right x: {rightAlignedX:0.##}, left-reference x: {sameColumnLeftX:0.##}.");

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

            document.Save();

            bytes = document.ToPdf(new ExcelPdfSaveOptions {
                IncludeSheetHeadings = false,
                HeaderRowCount = 0,
                PageSize = new PdfCore.PageSize(360, 220),
                Margins = PdfCore.PageMargins.Uniform(24)
            });
        }

        using (PdfPigDocument pdf = PdfPigDocument.Open(new MemoryStream(bytes))) {
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

            document.Save();

            bytes = document.ToPdf(new ExcelPdfSaveOptions {
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

        using (PdfPigDocument pdf = PdfPigDocument.Open(new MemoryStream(bytes))) {
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
            sheet.Cell(7, 1, "Zero");
            sheet.CellAt(7, 2).SetValue(0).SetNumberFormat("#,##0;(#,##0);-");

            ExcelCellStyleSnapshot currencyStyle = sheet.CellAt(2, 2).GetStyle();
            Assert.Equal("\"$\"#,##0.00", currencyStyle.NumberFormatCode);
            ExcelCellStyleSnapshot percentStyle = sheet.CellAt(3, 2).GetStyle();
            Assert.Equal("0.0%", percentStyle.NumberFormatCode);
            ExcelCellStyleSnapshot dateStyle = sheet.CellAt(4, 2).GetStyle();
            Assert.Equal("yyyy-mm-dd", dateStyle.NumberFormatCode);
            Assert.True(dateStyle.IsDateLike);

            document.Save();

            bytes = document.ToPdf(new ExcelPdfSaveOptions {
                IncludeSheetHeadings = false,
                HeaderRowCount = 1,
                PageSize = new PdfCore.PageSize(420, 360),
                Margins = PdfCore.PageMargins.Uniform(24)
            });
        }

        using PdfPigDocument pdf = PdfPigDocument.Open(new MemoryStream(bytes));
        string text = string.Concat(Enumerable.Range(1, pdf.NumberOfPages).Select(page => pdf.GetPage(page).Text));
        Assert.Contains("$1,234.50", text);
        Assert.Contains("25.7%", text);
        Assert.Contains("2026-01-15", text);
        Assert.Contains("30:05", text);
        Assert.Contains("(1,234)", text);
        Assert.Contains("Zero", text);
        Assert.Contains("-", text);
        Assert.DoesNotContain("01:05", text);
        Assert.DoesNotContain("1234.5", text);
        Assert.DoesNotContain("0.257", text);
        Assert.DoesNotContain("-1,234", text);
    }

    [Fact]
    public void SaveAsPdf_ExcelWorkbook_NumberFormatSections_DoNotExpandAllSemicolonSegments() {
        string workbookPath = Path.Combine(_directoryWithFiles, "ExcelPdfHostileNumberFormatSections.xlsx");
        string extraSections = string.Concat(Enumerable.Repeat(";[Red]0", 20000));

        byte[] bytes;
        using (ExcelDocument document = ExcelDocument.Create(workbookPath, "Formats")) {
            ExcelSheet sheet = document.Sheets[0];
            sheet.Cell(1, 1, "Kind");
            sheet.Cell(1, 2, "Value");
            sheet.Cell(2, 1, "ManySections");
            sheet.CellAt(2, 2).SetValue(1234).SetNumberFormat("#,##0" + extraSections);
            sheet.Cell(3, 1, "QuotedSectionSeparator");
            sheet.CellAt(3, 2).SetValue(-12).SetNumberFormat("0;\"minus;literal\"0" + extraSections);

            document.Save();

            bytes = document.ToPdf(new ExcelPdfSaveOptions {
                IncludeSheetHeadings = false,
                HeaderRowCount = 1,
                PageSize = new PdfCore.PageSize(420, 220),
                Margins = PdfCore.PageMargins.Uniform(24)
            });
        }

        using PdfPigDocument pdf = PdfPigDocument.Open(new MemoryStream(bytes));
        string text = string.Concat(Enumerable.Range(1, pdf.NumberOfPages).Select(page => pdf.GetPage(page).Text));
        Assert.Contains("ManySections", text);
        Assert.Contains("1,234", text);
        Assert.Contains("QuotedSectionSeparator", text);
        Assert.Contains("minus;literal-12", text);
    }

    [Fact]
    public void SaveAsPdf_ExcelWorkbook_Maps_Elapsed_Time_And_Quoted_Number_Literals() {
        string workbookPath = Path.Combine(_directoryWithFiles, "ExcelPdfElapsedAndQuotedFormats.xlsx");

        byte[] bytes;
        using (ExcelDocument document = ExcelDocument.Create(workbookPath, "Formats")) {
            ExcelSheet sheet = document.Sheets[0];
            sheet.Cell(1, 1, "Kind");
            sheet.Cell(1, 2, "Value");
            sheet.Cell(2, 1, "Elapsed");
            sheet.CellAt(2, 2).SetValue(1.5).SetNumberFormat("[h]:mm");
            sheet.Cell(3, 1, "Units");
            sheet.CellAt(3, 2).SetValue(12).SetNumberFormat("0 \"kg\"");
            sheet.Cell(4, 1, "Elapsed units");
            sheet.CellAt(4, 2).SetValue(1.5).SetNumberFormat("[h] \"hours\"");

            document.Save();

            bytes = document.ToPdf(new ExcelPdfSaveOptions {
                IncludeSheetHeadings = false,
                HeaderRowCount = 1,
                PageSize = new PdfCore.PageSize(360, 220),
                Margins = PdfCore.PageMargins.Uniform(24)
            });
        }

        using PdfPigDocument pdf = PdfPigDocument.Open(new MemoryStream(bytes));
        string text = pdf.GetPage(1).Text;
        Assert.Contains("Elapsed", text);
        Assert.Contains("36:00", text);
        Assert.Contains("Units", text);
        Assert.Contains("12 kg", text);
        Assert.Contains("Elapsed units", text);
        Assert.Contains("36 hours", text);
        Assert.DoesNotContain("12:00", text);
        Assert.DoesNotContain("hour0", text);
        Assert.DoesNotContain("kg12", text);
    }
}
