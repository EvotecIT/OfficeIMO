using System;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using OfficeIMO.OpenDocument;
using OfficeIMO.OpenDocument.Testing;
using OfficeIMO.Excel;
using OfficeIMO.Excel.OpenDocument;
using OfficeIMO.PowerPoint;
using OfficeIMO.PowerPoint.OpenDocument;
using OfficeIMO.Word;
using OfficeIMO.Word.OpenDocument;
using Xunit;

namespace OfficeIMO.OpenDocument.Converters.Tests;

public sealed class OpenDocumentCurrentReviewLossReportTests {
    private static readonly byte[] TinyPng = Convert.FromBase64String(
        "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII=");

    [Fact]
    public void WordToOdtReportsFlattenedNestedListLevels() {
        using WordDocument source = WordDocument.Create();
        WordList list = source.AddListNumbered();
        list.AddItem("Parent");
        list.AddItem("Child", 1);

        OdfConversionResult<OdtDocument> conversion = source.ToOpenDocumentResult();
        using OdtDocument target = conversion.Value;

        Assert.Contains(conversion.Report.Mappings, mapping => mapping.Feature == "list-levels" &&
            mapping.Status == OdfConversionMappingStatus.Approximated && mapping.Count == 1);
    }

    [Fact]
    public void OdtToWordReportsHeaderAndFooterImagesAsSkipped() {
        using OdtDocument source = OdtDocument.Create();
        source.PageLayout.Header.AddParagraph("Logo").AddImage(TinyPng, "header.png",
            OdfLength.Centimeters(1), OdfLength.Centimeters(1));

        OdfConversionResult<WordDocument> conversion = source.ToWordDocumentResult();
        using WordDocument target = conversion.Value;

        Assert.Contains(conversion.Report.Mappings, mapping => mapping.Feature == "images" &&
            mapping.Status == OdfConversionMappingStatus.Skipped && mapping.Count == 1);
    }

    [Fact]
    public void OdpToPowerPointReportsFlattenedListsAndMixedRuns() {
        using OdpPresentation source = OdpPresentation.Create();
        OdpTextBox textBox = source.AddSlide("Text").AddTextBox(
            OdfRect.FromCentimeters(1, 1, 8, 4), null, "Content");
        OdpParagraph mixed = textBox.AddParagraph("Plain ");
        mixed.AddRun("Bold").Bold = true;
        textBox.AddList().AddItem("Bullet");

        OdfConversionResult<PowerPointPresentation> conversion = source.ToPowerPointPresentationResult();
        using PowerPointPresentation target = conversion.Value;

        Assert.Contains(conversion.Report.Mappings, mapping => mapping.Feature == "text-lists" &&
            mapping.Status == OdfConversionMappingStatus.Approximated && mapping.Count == 1);
        Assert.Contains(conversion.Report.Mappings, mapping => mapping.Feature == "inline-formatting" &&
            mapping.Status == OdfConversionMappingStatus.Approximated && mapping.Count == 1);
        Assert.Contains("Plain Bold", target.Slides.Single().TextBoxes.Single().Text, StringComparison.Ordinal);
    }

    [Fact]
    public void OdtToWordToleratesMissingStylesAndReportsTableCellImages() {
        using OdtDocument template = OdtDocument.Create();
        template.AddParagraph("Minimal");
        template.AddTable(1, 1, "Media").Cell(0, 0).Paragraphs[0].AddImage(TinyPng, "cell.png",
            OdfLength.Centimeters(1), OdfLength.Centimeters(1));
        using OdtDocument source = OdtDocument.Open(new MemoryStream(RemovePackageEntry(template.ToBytes(), "styles.xml")));

        OdfConversionResult<WordDocument> conversion = source.ToWordDocumentResult();
        using WordDocument target = conversion.Value;

        Assert.Contains(target.CreateInspectionSnapshot().Sections.SelectMany(section => section.Elements)
            .OfType<WordParagraphSnapshot>(), paragraph => paragraph.Text == "Minimal");
        Assert.Contains(conversion.Report.Mappings, mapping => mapping.Feature == "images" &&
            mapping.Status == OdfConversionMappingStatus.Skipped && mapping.Count == 1);
    }

    [Fact]
    public void ExcelToOdsPreservesTypedValuesOnHyperlinkedCellsAndFormulaSeparators() {
        using ExcelDocument source = ExcelDocument.Create(new MemoryStream());
        ExcelSheet sheet = source.AddWorkSheet("Data");
        source.AddWorkSheet("Other, Sheet").CellAt(1, 1).SetValue(1);
        sheet.SetHyperlink(1, 1, "https://example.com", "42");
        sheet.CellAt(1, 1).SetValue(42);
        sheet.CellAt(1, 2).SetFormula("IF(A1=42,\"x,y\",\"other\")");
        sheet.CellAt(1, 3).SetFormula("SUM('Other, Sheet'!A1,A1)");

        OdfConversionResult<OdsDocument> conversion = source.ToOpenDocumentResult();
        using OdsDocument target = conversion.Value;
        OdsSheet converted = target.GetSheet("Data")!;

        Assert.Equal(OdsCellValueKind.Number, converted.GetValue(0, 0).Kind);
        Assert.Equal(42D, converted.GetValue(0, 0).AsDouble());
        Assert.Equal("https://example.com", converted.RowRuns[0].CellRuns[0].HyperlinkHref);
        Assert.Equal("of:=IF([.A1]=42;\"x,y\";\"other\")", converted.GetFormula(0, 1));
        Assert.Equal("of:=SUM([$'Other, Sheet'.A1];[.A1])", converted.GetFormula(0, 2));
    }

    [Fact]
    public void OdsToExcelCreatesInternalLinksWithoutLosingTypedValues() {
        using OdsDocument source = OdsDocument.Create();
        source.AddSheet("Target").Cell(0, 0).SetString("Destination");
        OdsCell linked = source.AddSheet("Links").Cell(0, 0);
        linked.SetNumber(42D);
        linked.SetHyperlink("Go", "#$'Target'.A1");

        OdfConversionResult<ExcelDocument> conversion = source.ToExcelDocumentResult();
        using ExcelDocument target = conversion.Value;
        ExcelWorksheetSnapshot links = target.CreateInspectionSnapshot().Worksheets.Single(sheet => sheet.Name == "Links");
        ExcelCellSnapshot cell = Assert.Single(links.Cells);

        Assert.Equal(42m, Convert.ToDecimal(cell.Value));
        Assert.NotNull(cell.Hyperlink);
        Assert.False(cell.Hyperlink!.IsExternal);
        Assert.Equal("'Target'!A1", cell.Hyperlink.Target);
    }

    [Fact]
    public void ExcelToOdsConvertsLowercaseFormulaReferences() {
        using ExcelDocument source = ExcelDocument.Create(new MemoryStream());
        ExcelSheet sheet = source.AddWorkSheet("Data");
        sheet.CellAt(1, 1).SetValue(1);
        sheet.CellAt(1, 2).SetFormula("sum(a1,'Data'!a1)");
        source.SetNamedRange("A1_total", "'Data'!$A$1", save: false);
        sheet.CellAt(1, 3).SetFormula("SUM(A1_total,a1)");

        using OdsDocument target = source.ToOpenDocument();

        Assert.Equal("of:=sum([.a1];[$'Data'.a1])", target.GetSheet("Data")!.GetFormula(0, 1));
        Assert.Equal("of:=SUM(A1_total;[.a1])", target.GetSheet("Data")!.GetFormula(0, 2));
    }

    [Fact]
    public void OdsToExcelPreservesRelativeExternalLinksAndMissingStyles() {
        using OdsDocument template = OdsDocument.Create();
        OdsCell linked = template.AddSheet("Links").Cell(0, 0);
        linked.SetString("Docs");
        linked.SetHyperlink("Docs", "docs/page.html");
        using OdsDocument source = OdsDocument.Open(new MemoryStream(RemovePackageEntry(template.ToBytes(), "styles.xml")));

        using ExcelDocument target = source.ToExcelDocument();
        ExcelCellSnapshot cell = Assert.Single(target.CreateInspectionSnapshot().Worksheets.Single().Cells);

        Assert.NotNull(cell.Hyperlink);
        Assert.True(cell.Hyperlink!.IsExternal);
        Assert.Equal("docs/page.html", cell.Hyperlink.Target);
    }

    [Fact]
    public void OdpToPowerPointPreservesStyledParagraphsNormalizedImagesAndMasterBackground() {
        using OdpPresentation template = OdpPresentation.Create();
        OdpMasterPage master = template.AddMasterPage("Brand");
        master.BackgroundColor = OdfColor.Parse("#445566");
        OdpSlide slide = template.AddSlide("Slide");
        slide.MasterPageName = master.Name;
        OdpTextBox textBox = slide.AddTextBox(OdfRect.FromCentimeters(1, 1, 8, 4));
        textBox.AddParagraph().AddRun("First").Bold = true;
        textBox.AddParagraph().AddRun("Second").Italic = true;
        OdpImage image = slide.AddImage(TinyPng, "pixel.png", OdfRect.FromCentimeters(1, 6, 2, 2));
        string escapedHref = "./" + image.Path.Replace(".png", "%2Epng") + "?cache=1";
        byte[] package = RewriteXmlEntry(template.ToBytes(), "content.xml", document =>
            document.Descendants().Single(element => element.Name.LocalName == "image")
                .SetAttributeValue(XName.Get("href", "http://www.w3.org/1999/xlink"), escapedHref));
        using OdpPresentation source = OdpPresentation.Open(new MemoryStream(package));

        using PowerPointPresentation target = source.ToPowerPointPresentation();
        PowerPointSlide converted = Assert.Single(target.Slides);

        Assert.Equal(new[] { "First", "Second" }, converted.TextBoxes.Single().Paragraphs.Select(paragraph => paragraph.Text));
        Assert.Single(converted.Pictures);
        Assert.Equal("445566", converted.GetBackground().Color);
    }

    [Fact]
    public void OdpToPowerPointToleratesMissingStylesPart() {
        using OdpPresentation template = OdpPresentation.Create();
        template.AddSlide("Minimal").AddTextBox(OdfRect.FromCentimeters(1, 1, 8, 2), "Text");
        using OdpPresentation source = OdpPresentation.Open(new MemoryStream(RemovePackageEntry(template.ToBytes(), "styles.xml")));

        using PowerPointPresentation target = source.ToPowerPointPresentation();

        Assert.Single(target.Slides);
        Assert.Contains("Text", target.Slides.Single().TextBoxes.Single().Text, StringComparison.Ordinal);
    }

    [Fact]
    public void OdsToExcelPrefersContentScopedDuplicateDataStyle() {
        using OdsDocument template = OdsDocument.Create();
        template.AddNumberStyle("Amount", 2);
        template.AddSheet("Data").Cell(0, 0).SetNumber(12.5);
        byte[] package = RewriteXmlEntry(template.ToBytes(), "styles.xml", document => {
            XNamespace office = "urn:oasis:names:tc:opendocument:xmlns:office:1.0";
            XNamespace style = "urn:oasis:names:tc:opendocument:xmlns:style:1.0";
            XNamespace number = "urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0";
            document.Root!.Element(office + "styles")!.Add(
                new XElement(number + "percentage-style", new XAttribute(style + "name", "Amount")));
        });
        using OdsDocument source = OdsDocument.Open(new MemoryStream(package));

        using ExcelDocument target = source.ToExcelDocument();

        Assert.Single(target.CreateInspectionSnapshot().Worksheets);
    }

    [Fact]
    public void WordToOdtReportsHeaderFooterTablesAndLaterSectionDefaults() {
        using WordDocument source = WordDocument.Create();
        source.AddHeadersAndFooters();
        source.Sections[0].Header.Default!.AddTable(1, 1).Rows[0].Cells[0].Paragraphs[0].Text = "Table";
        WordSection second = source.AddSection();
        second.AddHeadersAndFooters();
        second.Header.Default!.AddParagraph("Later header");

        OdfConversionResult<OdtDocument> conversion = source.ToOpenDocumentResult(new WordOpenDocumentConversionOptions {
            IncludeHeadersAndFooters = true
        });
        using OdtDocument target = conversion.Value;
        OdfConversionReport report = conversion.Report;

        Assert.Contains(report.Mappings, mapping => mapping.Feature == "header-footer-tables" &&
            mapping.Status == OdfConversionMappingStatus.Skipped && mapping.Count == 1);
        Assert.Contains(report.Mappings, mapping => mapping.Feature == "section-headers-footers" &&
            mapping.Status == OdfConversionMappingStatus.Skipped && mapping.Count >= 1);
    }

    private static byte[] RemovePackageEntry(byte[] packageBytes, string removedPath) =>
        OdfTestPackageRewriter.Remove(packageBytes, removedPath);

    private static byte[] RewriteXmlEntry(byte[] packageBytes, string path, Action<XDocument> rewrite) {
        return OdfTestPackageRewriter.Rewrite(packageBytes, (name, bytes) => {
            if (name == path) {
                XDocument document = XDocument.Parse(Encoding.UTF8.GetString(bytes));
                rewrite(document);
                return Encoding.UTF8.GetBytes(document.ToString(SaveOptions.DisableFormatting));
            }
            return bytes;
        });
    }
}
