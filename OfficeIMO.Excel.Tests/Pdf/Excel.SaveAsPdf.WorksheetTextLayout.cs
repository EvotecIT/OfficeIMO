using OfficeIMO.Excel;
using OfficeIMO.Excel.Pdf;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using Xunit;

namespace OfficeIMO.Tests;

public partial class Excel {
    private static PdfPigDocument OpenWorksheetPdf(ExcelDocument document, ExcelPdfExportOptions? options = null) =>
        PdfPigDocument.Open(document.ToPdf(options ?? new ExcelPdfExportOptions()).ToBytes(new OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true }));

    /// <summary>Registers a bundled font as the workbook default so column units and "####" decisions are deterministic.</summary>
    private static ExcelPdfExportOptions CreateDeterministicFontOptions() => new ExcelPdfExportOptions {
        PdfOptions = new OfficeIMO.Pdf.PdfCreateOptions().RegisterNamedFontFamily(new OfficeIMO.Pdf.PdfEmbeddedFontFamily(
            "Calibri", OfficeIMO.TestAssets.PdfTestFontAssets.LoadBundledOpenTypeCffFont()))
    };

    private static string GetPageText(UglyToad.PdfPig.Content.Page page) =>
        string.Join(" ", page.GetWords().Select(word => word.Text));

    private static List<OfficeDiagnostic> ConvertAndCollectClippedCells(ExcelDocument document) {
        using var stream = new MemoryStream();
        OfficeConversionReport report = document.SaveAsPdf(stream);
        return report.Diagnostics.Where(diagnostic => diagnostic.Code == "table-cell-content-clipped").ToList();
    }

    [Fact]
    public void SaveAsPdf_LeftAlignedTextOverflowsIntoEmptyNeighbours() {
        using ExcelDocument spilled = ExcelDocument.Create();
        ExcelSheet sheet = spilled.AddWorksheet("Overflow");
        sheet.Cell(1, 1, "Quarterly sales against target and plan");
        sheet.Cell(1, 5, "Edge");
        Assert.Empty(ConvertAndCollectClippedCells(spilled));

        using ExcelDocument blocked = ExcelDocument.Create();
        ExcelSheet blockedSheet = blocked.AddWorksheet("Blocked");
        blockedSheet.Cell(1, 1, "Quarterly sales against target and plan");
        blockedSheet.Cell(1, 2, "Next");
        // A neighbour with a value stops the overflow, as in Excel.
        Assert.Single(ConvertAndCollectClippedCells(blocked));
    }

    [Fact]
    public void SaveAsPdf_RightAndCentreAlignedTextOverflowTowardsEmptyNeighbours() {
        using ExcelDocument document = ExcelDocument.Create();
        ExcelSheet sheet = document.AddWorksheet("Overflow");
        sheet.Cell(1, 1, "A1");
        sheet.Cell(1, 4, "Right aligned label spills left");
        sheet.CellAlign(1, 4, ExcelHorizontalAlignment.Right);
        sheet.Cell(2, 3, "Centred label spills");
        sheet.CellAlign(2, 3, ExcelHorizontalAlignment.Center);
        sheet.Cell(2, 6, "F2");
        Assert.Empty(ConvertAndCollectClippedCells(document));

        using PdfPigDocument pdf = OpenWorksheetPdf(document);
        var page = pdf.GetPage(1);
        double a1Start = FindWordStartX(page, "A1");
        double fiveColumns = FindWordStartX(page, "F2") - a1Start;
        // Right-aligned text in column D starts left of column D.
        Assert.True(FindWordStartX(page, "Right") < a1Start + 3D * fiveColumns / 5D);
        // Centred text in column C starts left of column C.
        Assert.True(FindWordStartX(page, "Centred") < a1Start + 2D * fiveColumns / 5D);
    }

    [Fact]
    public void SaveAsPdf_WrappedAndNumericCellsDoNotOverflow() {
        using ExcelDocument document = ExcelDocument.Create();
        ExcelSheet sheet = document.AddWorksheet("NoOverflow");
        sheet.Cell(1, 1, 1234567890123.456);
        sheet.Cell(2, 1, new DateTime(2026, 10, 27));
        sheet.FormatCell(2, 1, "yyyy-mm-dd");
        sheet.SetColumnWidth(1, 6);
        using (PdfPigDocument pdf = OpenWorksheetPdf(document, CreateDeterministicFontOptions())) {
            string text = GetPageText(pdf.GetPage(1));
            Assert.DoesNotContain("1234567890123", text);
            Assert.DoesNotContain("2026-10-2", text);
            Assert.Contains("####", text);
        }

        using ExcelDocument wrappedDocument = ExcelDocument.Create();
        ExcelSheet wrappedSheet = wrappedDocument.AddWorksheet("Wrapped");
        wrappedSheet.Cell(1, 2, "Wrapped text stays inside its column");
        wrappedSheet.CellWrapText(1, 2);
        using PdfPigDocument wrappedPdf = OpenWorksheetPdf(wrappedDocument);
        var words = wrappedPdf.GetPage(1).GetWords().ToArray();
        var first = Assert.Single(words, word => word.Text == "Wrapped");
        var last = Assert.Single(words, word => word.Text == "column");
        // Wrapped text breaks into lines inside its cell instead of overflowing.
        Assert.True(first.BoundingBox.Bottom > last.BoundingBox.Top);
    }

    [Fact]
    public void SaveAsPdf_GeneralNumbersRoundToElevenCharacters() {
        using ExcelDocument document = ExcelDocument.Create();
        ExcelSheet sheet = document.AddWorksheet("General");
        sheet.Cell(1, 1, Math.PI);
        sheet.SetColumnWidth(1, 20);

        using PdfPigDocument pdf = OpenWorksheetPdf(document);
        Assert.Contains("3.141592654", GetPageText(pdf.GetPage(1)));
        Assert.DoesNotContain("3.14159265358979", GetPageText(pdf.GetPage(1)));
    }

    [Fact]
    public void SaveAsPdf_RowsWithoutStoredHeightGrowForLargeFonts() {
        using ExcelDocument document = ExcelDocument.Create();
        ExcelSheet sheet = document.AddWorksheet("Heights");
        sheet.Cell(1, 1, "Title");
        sheet.CellFontSize(1, 1, 20);
        sheet.Cell(2, 1, "Body");

        using PdfPigDocument pdf = OpenWorksheetPdf(document);
        var words = pdf.GetPage(1).GetWords().ToArray();
        var title = Assert.Single(words, word => word.Text == "Title");
        var body = Assert.Single(words, word => word.Text == "Body");
        // A 20 pt row is about 27 pt tall; at the default 15 pt the title descends into row 2.
        double gap = title.BoundingBox.Bottom - body.BoundingBox.Bottom;
        Assert.InRange(gap, 13D, 20D);
        Assert.True(title.BoundingBox.Bottom - body.BoundingBox.Top > 1D);
    }

    [Fact]
    public void SaveAsPdf_WrappedRowsWithoutStoredHeightGrowForEveryLine() {
        using ExcelDocument document = ExcelDocument.Create();
        ExcelSheet sheet = document.AddWorksheet("Wrap");
        sheet.Cell(1, 1, "one two three four five six");
        sheet.CellWrapText(1, 1);
        sheet.SetColumnWidth(1, 6);
        sheet.Cell(2, 1, "after");

        using PdfPigDocument pdf = OpenWorksheetPdf(document);
        string text = GetPageText(pdf.GetPage(1));
        foreach (string word in new[] { "one", "two", "three", "four", "five", "six", "after" }) {
            Assert.Contains(word, text);
        }

        var words = pdf.GetPage(1).GetWords().ToArray();
        var six = Assert.Single(words, word => word.Text == "six");
        var after = Assert.Single(words, word => word.Text == "after");
        Assert.True(six.BoundingBox.Bottom > after.BoundingBox.Top);
    }

    [Fact]
    public void SaveAsPdf_MergedTitleExtendingBeyondValuesIsNotClipped() {
        using ExcelDocument document = ExcelDocument.Create();
        ExcelSheet sheet = document.AddWorksheet("Merged");
        sheet.Cell(1, 1, "Quarterly sales against target");
        sheet.CellFontSize(1, 1, 16);
        sheet.MergeRange("A1:F1");
        sheet.Cell(2, 1, "Q1");
        sheet.Cell(2, 2, 42);

        using PdfPigDocument pdf = OpenWorksheetPdf(document);
        Assert.Contains("Quarterly sales against target", GetPageText(pdf.GetPage(1)));
    }

    [Fact]
    public void SaveAsPdf_UncachedFormulasShowCalculatedValues() {
        using ExcelDocument document = ExcelDocument.Create();
        ExcelSheet sheet = document.AddWorksheet("Formulas");
        sheet.Cell(1, 1, 2D);
        sheet.Cell(2, 1, 3D);
        sheet.CellFormula(3, 1, "SUM(A1:A2)");
        sheet.SetColumnWidth(1, 12);

        using PdfPigDocument pdf = OpenWorksheetPdf(document);
        string text = GetPageText(pdf.GetPage(1));
        Assert.DoesNotContain("SUM", text);
        Assert.Contains("5", text.Split(' '));
    }

    [Fact]
    public void SaveAsPdf_CustomNumberFormatKeepsLeadingZeros() {
        using ExcelDocument document = ExcelDocument.Create();
        ExcelSheet sheet = document.AddWorksheet("Formats");
        sheet.Cell(1, 1, 123D);
        sheet.FormatCell(1, 1, "000000");
        sheet.Cell(2, 1, 1234D);
        sheet.FormatCell(2, 1, "#,##0000000");
        sheet.SetColumnWidth(1, 16);

        using PdfPigDocument pdf = OpenWorksheetPdf(document);
        string text = GetPageText(pdf.GetPage(1));
        Assert.Contains("000123", text);
        Assert.Contains("0,001,234", text);
    }

    [Fact]
    public void SaveAsPdf_EmptyWorkbookProducesBlankPageAndDiagnostic() {
        using ExcelDocument document = ExcelDocument.Create();
        document.AddWorksheet("Empty");

        using var stream = new MemoryStream();
        OfficeConversionReport report = document.SaveAsPdf(stream);
        using PdfPigDocument pdf = PdfPigDocument.Open(stream.ToArray());
        Assert.Equal(1, pdf.NumberOfPages);
        Assert.Empty(pdf.GetPage(1).GetWords());
        Assert.Contains(report.Diagnostics, diagnostic => diagnostic.Code.Contains("workbook-empty", StringComparison.Ordinal));
    }
}
