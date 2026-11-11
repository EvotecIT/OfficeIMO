using System.Text.RegularExpressions;
using OfficeIMO.Excel;
using OfficeIMO.Excel.OpenDocument;
using OfficeIMO.OpenDocument.Odp.Pdf;
using OfficeIMO.OpenDocument.Ods.Pdf;
using OfficeIMO.OpenDocument.Odt.Pdf;
using OfficeIMO.Pdf;
using OfficeIMO.PowerPoint;
using OfficeIMO.PowerPoint.OpenDocument;
using OfficeIMO.Word;
using OfficeIMO.Word.OpenDocument;
using Xunit;

namespace OfficeIMO.OpenDocument.Converters.Tests;

/// <summary>
/// Converts the LibreOffice and Microsoft Office fixtures recorded in
/// OfficeIMO.OpenDocument.Tests/Fixtures/producer-manifest.json. Each producer serializes
/// styles, repeated rows and columns, and page masters differently from OfficeIMO.
/// </summary>
public sealed class OpenDocumentIndependentProducerContracts {
    [Theory]
    [InlineData("libreoffice-writer-basic.odt", "note-configuration")]
    [InlineData("microsoft-word-basic.odt", "note-configuration")]
    public void ProducerOdtRoundTripsThroughDocxWithBodyOrderAndTables(string fixture, string expectedLoss) {
        OdtDocument source = OdtDocument.Load(FixturePath(fixture));

        OdfConversionResult<WordDocument> toWord = source.ConvertToWordWithEvidence();
        using WordDocument word = toWord.Value;

        Assert.Empty(word.ValidateDocument());
        Assert.Equal(new[] { "Testing Tables on Load", "table 2x2", "Another table on load", "table 2x2" },
            word.Inspect().Sections.SelectMany(section => section.Elements).Select(DescribeWordBlock)
                .Where(block => block.Length > 0));
        Assert.Equal(expectedLoss.Split('|').OrderBy(value => value, StringComparer.Ordinal), LossFeatures(toWord.Report));

        OdfConversionResult<OdtDocument> toOdt = word.ConvertToOpenDocumentWithEvidence();
        OdtDocument reopened = OdtDocument.Load(new MemoryStream(toOdt.Value.ToBytes()));
        Assert.True(reopened.Validate().IsValid);
        Assert.Equal(NonEmptyParagraphs(source), NonEmptyParagraphs(reopened));
        Assert.Equal(source.Tables.Select(TableShape), reopened.Tables.Select(TableShape));
    }

    [Theory]
    [InlineData("libreoffice-calc-basic.ods", "blank-cell-styles|blank-cell-tail-styles|source-master-pages|source-text-fields|source-unmodeled-style-maps")]
    [InlineData("microsoft-excel-basic.ods", "blank-cell-styles|blank-cell-tail-styles|cell-layout|source-master-pages")]
    public void ProducerOdsRoundTripsThroughXlsxWithoutMaterializingRowPadding(string fixture, string expectedLoss) {
        OdsDocument source = OdsDocument.Load(FixturePath(fixture));

        OdfConversionResult<ExcelDocument> toExcel = source.ConvertToExcelWithEvidence();
        using ExcelDocument excel = toExcel.Value;
        var info = excel.Inspect();

        Assert.Empty(excel.ValidateOpenXml());
        Assert.Equal(new[] { "Sheet1", "Sheet2", "DifferentTest", "Sheet3" }, info.Worksheets.Select(sheet => sheet.Name));
        Assert.Equal(ExpectedWorkbookText, info.Worksheets.SelectMany(sheet => sheet.Cells
            .Where(cell => cell.Value is string text && text.Length > 0)
            .Select(cell => sheet.Name + "!" + SpreadsheetCell(cell.Row, cell.Column) + "=" + cell.Value)));
        // Both producers pad every used row to column XFD with one styled blank run.
        Assert.All(info.Worksheets.SelectMany(sheet => sheet.Cells), cell => Assert.True(cell.Row <= 3 && cell.Column <= 4,
            "Producer padding was materialized at " + SpreadsheetCell(cell.Row, cell.Column)));
        ExcelCellInfo red = info.Worksheets[0].Cells.Single(cell => cell.Row == 1 && cell.Column == 1);
        ExcelCellInfo yellow = info.Worksheets[0].Cells.Single(cell => cell.Row == 2 && cell.Column == 2);
        Assert.EndsWith("FF0000", red.Style!.FontColorArgb, StringComparison.OrdinalIgnoreCase);
        Assert.EndsWith("FFFF00", yellow.Style!.FillColorArgb, StringComparison.OrdinalIgnoreCase);
        Assert.Equal(expectedLoss.Split('|').OrderBy(value => value, StringComparer.Ordinal), LossFeatures(toExcel.Report));

        OdfConversionResult<OdsDocument> toOds = excel.ConvertToOpenDocumentWithEvidence();
        OdsDocument reopened = OdsDocument.Load(new MemoryStream(toOds.Value.ToBytes()));
        Assert.True(reopened.Validate().IsValid);
        Assert.Equal(ExpectedWorkbookText, reopened.Sheets.SelectMany(sheet => OdsText(sheet)));
    }

    [Theory]
    [InlineData("libreoffice-impress-basic.odp", "document-metadata|handout-master|shape-appearance|shape-layers|slide-show-settings|source-presentation-notes|source-text-fields|speaker-notes")]
    [InlineData("microsoft-powerpoint-basic.odp", "document-metadata|shape-accessibility|shape-appearance|slide-backgrounds|source-text-fields")]
    public void ProducerOdpRoundTripsThroughPptxWithSlideGeometryAndPlaceholders(string fixture, string expectedLoss) {
        OdpPresentation source = OdpPresentation.Load(FixturePath(fixture));

        OdfConversionResult<PowerPointDocument> toPowerPoint = source.ConvertToPowerPointWithEvidence();
        using PowerPointDocument presentation = toPowerPoint.Value;

        PowerPointSlide slide = Assert.Single(presentation.Slides);
        Assert.Equal(960, presentation.SlideSize.Width.ToPoints(), 0);
        Assert.Equal(540, presentation.SlideSize.Height.ToPoints(), 0);
        Assert.Equal(2, slide.TextBoxes.Count(box => box.PlaceholderType != null));
        Assert.Equal(expectedLoss.Split('|').OrderBy(value => value, StringComparer.Ordinal), LossFeatures(toPowerPoint.Report));

        OdfConversionResult<OdpPresentation> toOdp = presentation.ConvertToOpenDocumentWithEvidence();
        OdpPresentation reopened = OdpPresentation.Load(new MemoryStream(toOdp.Value.ToBytes()));
        Assert.True(reopened.Validate().IsValid);
        Assert.Equal(2, Assert.Single(reopened.Slides).Shapes.OfType<OdpTextBox>().Count());
        Assert.Equal(960, reopened.PageWidth.ToPoints(), 0);
        Assert.Equal(540, reopened.PageHeight.ToPoints(), 0);
    }

    [Theory]
    [InlineData("libreoffice-writer-basic.odt")]
    [InlineData("microsoft-word-basic.odt")]
    public void ProducerOdtExportsSearchablePdfAndCarriesSourceLoss(string fixture) {
        OdtDocument source = OdtDocument.Load(FixturePath(fixture));
        string[] sourceLoss = LossFeatures(source.ConvertToWordWithEvidence().Report);

        PdfDocument pdf = ExportPdf(source.ToPdf(null, out OfficeConversionReport report));

        Assert.Single(pdf.Pages);
        Assert.Equal("Testing Tables on Load Another table on load", NormalizedText(pdf));
        Assert.Equal(sourceLoss, OmittedLocations(report, sourceLoss));
        Assert.Throws<OfficeConversionException>(() => source.ToPdf(new OdtPdfExportOptions { FailOnLoss = true }, out _));
    }

    [Theory]
    [InlineData("libreoffice-calc-basic.ods")]
    [InlineData("microsoft-excel-basic.ods")]
    public void ProducerOdsExportsOnePdfPagePerUsedSheetAndCarriesSourceLoss(string fixture) {
        OdsDocument source = OdsDocument.Load(FixturePath(fixture));
        string[] sourceLoss = LossFeatures(source.ConvertToExcelWithEvidence().Report);

        PdfDocument pdf = ExportPdf(source.ToPdf(null, out OfficeConversionReport report));

        Assert.Equal(2, pdf.Pages.Count);
        Assert.Equal("This is test Another Test Mmmm", NormalizedText(pdf.Pages[0].ExtractText()));
        Assert.Equal("Another one Test", NormalizedText(pdf.Pages[1].ExtractText()));
        Assert.Equal(sourceLoss, OmittedLocations(report, sourceLoss));
        Assert.Throws<OfficeConversionException>(() => source.ToPdf(new OdsPdfExportOptions { FailOnLoss = true }, out _));
    }

    [Theory]
    [InlineData("libreoffice-impress-basic.odp")]
    [InlineData("microsoft-powerpoint-basic.odp")]
    public void ProducerOdpExportsSlideSizedPdfAndCarriesSourceLoss(string fixture) {
        OdpPresentation source = OdpPresentation.Load(FixturePath(fixture));
        string[] sourceLoss = LossFeatures(source.ConvertToPowerPointWithEvidence().Report);

        PdfDocument pdf = ExportPdf(source.ToPdf(null, out OfficeConversionReport report));

        (double width, double height) = Assert.Single(pdf.Pages).GetPageSize();
        Assert.Equal(960, width, 0);
        Assert.Equal(540, height, 0);
        Assert.Equal(sourceLoss, OmittedLocations(report, sourceLoss));
        Assert.Throws<OfficeConversionException>(() => source.ToPdf(new OdpPdfExportOptions { FailOnLoss = true }, out _));
    }

    [Theory]
    [InlineData("libreoffice-calc-basic.ods")]
    [InlineData("microsoft-excel-basic.ods")]
    public void ProducerOdsImagesRenderEverySheetAndCarryBridgeLoss(string fixture) {
        OdsDocument source = OdsDocument.Load(FixturePath(fixture));

        IReadOnlyList<OfficeIMO.Drawing.OfficeImageExportResult> images =
            source.ToImages(OfficeIMO.Drawing.OfficeImageExportFormat.Png);

        Assert.Equal(new[] { "Sheet1", "Sheet2", "DifferentTest", "Sheet3" }, images.Select(image => image.Name));
        Assert.All(images, image => Assert.Contains(image.Diagnostics, diagnostic =>
            diagnostic.Code == "ODF_IMAGE_BLANK_CELL_TAIL_STYLES_UNSUPPORTED" &&
            diagnostic.LossKind != OfficeConversionLossKind.None));
    }

    private static readonly string[] ExpectedWorkbookText = {
        "Sheet1!A1=This is test",
        "Sheet1!B2=Another Test",
        "Sheet1!C3=Mmmm",
        "Sheet3!A1=Another one",
        "Sheet3!D3=Test"
    };

    private static string FixturePath(string name) => Path.Combine(AppContext.BaseDirectory, "Fixtures", name);

    private static string[] LossFeatures(OdfConversionReport report) => report.Mappings
        .Where(mapping => mapping.Status is OdfConversionMappingStatus.Skipped or OdfConversionMappingStatus.Unsupported)
        .Select(mapping => mapping.Feature)
        // Producer extension namespaces vary by producer version; each is still reported as its own finding.
        .Where(feature => !feature.StartsWith("source-foreign-namespace:", StringComparison.Ordinal))
        .Distinct(StringComparer.Ordinal)
        .OrderBy(feature => feature, StringComparer.Ordinal)
        .ToArray();

    private static string[] OmittedLocations(OfficeConversionReport report, IEnumerable<string> features) => report.Diagnostics
        .Where(diagnostic => diagnostic.LossKind != OfficeConversionLossKind.None && features.Contains(diagnostic.Location, StringComparer.Ordinal))
        .Select(diagnostic => diagnostic.Location!)
        .Distinct(StringComparer.Ordinal)
        .OrderBy(location => location, StringComparer.Ordinal)
        .ToArray();

    private static PdfDocument ExportPdf(PdfDocument pdf) {
        using (pdf) return PdfDocument.Load(new MemoryStream(pdf.ToBytes(new PdfSaveOptions { AllowLoss = true })));
    }

    private static string NormalizedText(PdfDocument pdf) => NormalizedText(pdf.ExtractText());

    private static string NormalizedText(string text) => Regex.Replace(text, @"\s+", " ").Trim();

    private static string DescribeWordBlock(object block) => block switch {
        WordTableInfo table => "table " + table.RowCount + "x" + table.ColumnCount,
        WordParagraphInfo paragraph => paragraph.Text,
        _ => string.Empty
    };

    private static IEnumerable<string> NonEmptyParagraphs(OdtDocument document) =>
        document.Paragraphs.Select(paragraph => paragraph.Text).Where(text => text.Length > 0);

    private static string TableShape(OdtTable table) => table.Rows.Count + "x" + table.Rows[0].Cells.Count;

    private static IEnumerable<string> OdsText(OdsSheet sheet) {
        if (sheet.UsedRange is not { } used) yield break;
        for (long row = used.FirstRow; row <= used.LastRow; row++) {
            for (long column = used.FirstColumn; column <= used.LastColumn; column++) {
                string text = sheet.GetValue(row, column).DisplayText;
                if (text.Length > 0) yield return sheet.Name + "!" + SpreadsheetCell((int)row + 1, (int)column + 1) + "=" + text;
            }
        }
    }

    private static string SpreadsheetCell(int row, int column) => ((char)('A' + column - 1)).ToString() + row;
}
