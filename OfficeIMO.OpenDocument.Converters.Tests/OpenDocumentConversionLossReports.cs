using System;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using OfficeIMO.Excel;
using OfficeIMO.Excel.OpenDocument;
using OfficeIMO.OpenDocument;
using OfficeIMO.PowerPoint;
using OfficeIMO.PowerPoint.OpenDocument;
using OfficeIMO.Word;
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

        Assert.Equal("of:=IF([.A1]=\"B2\";1;0)", target.GetSheet("Data")!.Cell(0, 1).Formula);

        target.GetSheet("Data")!.Cell(0, 2).Formula = "of:=IF([.A1]=\"[.B2]\",1,0)";
        OdfConversionResult<ExcelDocument> reverse = target.ToExcelDocument();
        using ExcelDocument roundTrip = reverse.Document;
        ExcelCellSnapshot reverseFormula = roundTrip.CreateInspectionSnapshot().Worksheets.Single().Cells
            .Single(cell => cell.Row == 1 && cell.Column == 3);
        Assert.Equal("IF(A1=\"[.B2]\",1,0)", reverseFormula.Formula);
    }

    [Fact]
    public void ExcelFormulaConversionMapsSheetQualifiedReferences() {
        using ExcelDocument source = ExcelDocument.Create(new MemoryStream(), autoSave: false);
        ExcelSheet data = source.AddWorkSheet("Data");
        source.AddWorkSheet("Other").CellAt(1, 1).SetValue(1);
        source.AddWorkSheet("Other Sheet").CellAt(2, 2).SetValue(2);
        data.CellAt(1, 1).SetFormula("SUM(Other!A1,'Other Sheet'!B2:C3)");

        OdfConversionResult<OdsDocument> conversion = source.ToOpenDocument();
        using OdsDocument target = conversion.Document;

        Assert.Equal("of:=SUM([$'Other'.A1];[$'Other Sheet'.B2:.C3])",
            target.GetSheet("Data")!.Cell(0, 0).Formula);
        OdfConversionResult<ExcelDocument> reverse = target.ToExcelDocument();
        using ExcelDocument roundTrip = reverse.Document;
        ExcelCellSnapshot formula = roundTrip.CreateInspectionSnapshot().Worksheets.Single(sheet => sheet.Name == "Data").Cells.Single();
        Assert.Equal("SUM('Other'!A1,'Other Sheet'!B2:C3)", formula.Formula);
    }

    [Fact]
    public void OpenFormulaConversionMapsArgumentSeparatorsOutsideStrings() {
        using OdsDocument source = OdsDocument.Create();
        OdsSheet sheet = source.AddSheet("Data");
        sheet.Cell(0, 0).SetNumber(1);
        sheet.Cell(0, 1).Formula = "of:=IF([.A1]>0;\"yes;still\";\"no\")";

        OdfConversionResult<ExcelDocument> conversion = source.ToExcelDocument();
        using ExcelDocument target = conversion.Document;

        ExcelCellSnapshot formula = target.CreateInspectionSnapshot().Worksheets.Single().Cells
            .Single(cell => cell.Row == 1 && cell.Column == 2);
        Assert.Equal("IF(A1>0,\"yes;still\",\"no\")", formula.Formula);
    }

    [Fact]
    public void DuplicateSheetLocalExcelNamesAreDisambiguatedWithoutAborting() {
        using ExcelDocument source = ExcelDocument.Create(new MemoryStream(), autoSave: false);
        ExcelSheet first = source.AddWorkSheet("First Sheet");
        ExcelSheet second = source.AddWorkSheet("Second Sheet");
        first.CellAt(1, 1).SetValue(1);
        second.CellAt(1, 1).SetValue(2);
        first.SetNamedRange("LocalValue", "A1", save: false);
        second.SetNamedRange("LocalValue", "A1", save: false);

        OdfConversionResult<OdsDocument> conversion = source.ToOpenDocument();
        using OdsDocument target = conversion.Document;

        Assert.Equal(2, target.NamedRanges.Count);
        Assert.Equal(2, target.NamedRanges.Select(named => named.Name).Distinct(StringComparer.Ordinal).Count());
        Assert.Contains(target.NamedRanges, named => named.CellRangeAddress.Contains("First Sheet"));
        Assert.Contains(target.NamedRanges, named => named.CellRangeAddress.Contains("Second Sheet"));
        Assert.Contains(conversion.Report.Mappings, mapping => mapping.Feature == "sheet-local-named-ranges" &&
            mapping.Status == OdfConversionMappingStatus.Approximated && mapping.Count == 1);
    }

    [Fact]
    public void WordAutomaticColorsAndUnsupportedImagesDoNotAbortConversion() {
        using WordDocument source = WordDocument.Create();
        source.AddParagraph("Automatic color").ColorHex = "auto";
        using var tiff = new MemoryStream(new byte[] { 0x49, 0x49, 0x2A, 0x00, 0x00, 0x00, 0x00, 0x00 });
        source.AddParagraph().AddImage(tiff, "unsupported.tiff", 10, 10);

        OdfConversionResult<OdtDocument> conversion = source.ToOpenDocument();
        using OdtDocument target = conversion.Document;

        Assert.Equal("Automatic color", target.Paragraphs.First().Text);
        Assert.Null(target.Paragraphs.First().Spans.Single().Color);
        Assert.Empty(target.Paragraphs.SelectMany(paragraph => paragraph.Images));
        Assert.Contains(conversion.Report.Mappings, mapping => mapping.Feature == "images" &&
            mapping.Status == OdfConversionMappingStatus.Unsupported && mapping.Count == 1);
    }

    [Fact]
    public void OdtToWordPreservesRelativeLinksAndSkipsUnsupportedImages() {
        using OdtDocument source = OdtDocument.Create();
        source.AddParagraph().AddHyperlink("Relative", "docs/page.html");
        byte[] webp = { 0x52, 0x49, 0x46, 0x46, 0x04, 0x00, 0x00, 0x00, 0x57, 0x45, 0x42, 0x50 };
        source.AddParagraph("Image").AddImage(webp, "pixel.webp", OdfLength.Centimeters(1), OdfLength.Centimeters(1));

        OdfConversionResult<WordDocument> conversion = source.ToWordDocument();
        using WordDocument target = conversion.Document;

        WordRunSnapshot link = target.CreateInspectionSnapshot().Sections.SelectMany(section => section.Elements)
            .OfType<WordParagraphSnapshot>().Single(paragraph => paragraph.Text == "Relative").Runs.Single();
        Assert.Equal("docs/page.html", link.HyperlinkUri);
        Assert.Contains(conversion.Report.Mappings, mapping => mapping.Feature == "images" &&
            mapping.Status == OdfConversionMappingStatus.Skipped && mapping.Count == 1);
    }

    [Fact]
    public void UnavailableOdfImagesAreReportedWithoutAbortingOfficeExport() {
        XNamespace draw = "urn:oasis:names:tc:opendocument:xmlns:drawing:1.0";
        XNamespace xlink = "http://www.w3.org/1999/xlink";
        XNamespace office = "urn:oasis:names:tc:opendocument:xmlns:office:1.0";
        byte[] png = Convert.FromBase64String(
            "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII=");

        using OdtDocument textSource = OdtDocument.Create();
        textSource.AddParagraph().AddImage(png, "missing.png", OdfLength.Centimeters(1), OdfLength.Centimeters(1));
        using OdtDocument brokenText = OpenBrokenFlat(textSource.ToFlatXml(), draw, xlink, office,
            stream => OdtDocument.OpenFlatXml(stream));
        OdfConversionResult<WordDocument> wordConversion = brokenText.ToWordDocument();
        using WordDocument word = wordConversion.Document;

        Assert.Contains(wordConversion.Report.Mappings, mapping => mapping.Feature == "images" &&
            mapping.Status == OdfConversionMappingStatus.Skipped && mapping.Count == 1);

        using OdpPresentation presentationSource = OdpPresentation.Create();
        presentationSource.AddSlide("Broken").AddImage(png, "missing.png", OdfRect.FromCentimeters(1, 1, 2, 2));
        using OdpPresentation brokenPresentation = OpenBrokenFlat(presentationSource.ToFlatXml(), draw, xlink, office,
            stream => OdpPresentation.OpenFlatXml(stream));
        OdfConversionResult<PowerPointPresentation> powerPointConversion = brokenPresentation.ToPowerPointPresentation();
        using PowerPointPresentation powerPoint = powerPointConversion.Document;

        Assert.Contains(powerPointConversion.Report.Mappings, mapping => mapping.Feature == "images" &&
            mapping.Status == OdfConversionMappingStatus.Unsupported && mapping.Count == 1);
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
    public void OdsToExcelKeepsOneWorksheetVisibleWhenEverySourceSheetIsHidden() {
        using OdsDocument source = OdsDocument.Create();
        source.AddSheet("First").Hidden = true;
        source.AddSheet("Second").Hidden = true;

        OdfConversionResult<ExcelDocument> conversion = source.ToExcelDocument();
        using ExcelDocument target = conversion.Document;
        ExcelWorkbookSnapshot snapshot = target.CreateInspectionSnapshot();

        Assert.Empty(target.ValidateDocument());
        Assert.False(snapshot.Worksheets[0].Hidden);
        Assert.True(snapshot.Worksheets[0].IsActive);
        Assert.True(snapshot.Worksheets[1].Hidden);
        Assert.Contains(conversion.Report.Mappings, mapping => mapping.Feature == "worksheet-visibility" &&
            mapping.Status == OdfConversionMappingStatus.Approximated && mapping.Count == 1);
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

    private static T OpenBrokenFlat<T>(XDocument flat, XNamespace draw, XNamespace xlink, XNamespace office,
        Func<Stream, T> open) where T : IDisposable {
        XElement image = flat.Descendants(draw + "image").Single();
        image.SetAttributeValue(xlink + "href", "https://example.test/missing.png");
        image.Elements(office + "binary-data").Remove();
        var stream = new MemoryStream();
        flat.Save(stream);
        stream.Position = 0;
        try {
            return open(stream);
        } finally {
            stream.Dispose();
        }
    }
}
