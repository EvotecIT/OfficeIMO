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
    public void PowerPointToOdp_Reports_Formatting_And_Notes_Disabled_By_Options() {
        using PowerPointPresentation source = PowerPointPresentation.Create();
        PowerPointSlide slide = source.AddSlide();
        PowerPointTextBox text = slide.AddTextBoxPoints("Styled", 10, 10, 200, 40);
        text.Paragraphs[0].Runs[0].Bold = true;
        text.FillColor = "112233";
        slide.Notes.Text = "Private note";

        OdfConversionResult<OdpPresentation> conversion = source.ToOpenDocumentResult(
            new PowerPointOpenDocumentConversionOptions {
                IncludeBasicFormatting = false,
                IncludeSpeakerNotes = false
            });

        Assert.Contains(conversion.Report.Mappings, mapping => mapping.Feature == "basic-formatting" &&
            mapping.Status == OdfConversionMappingStatus.Skipped && mapping.Count >= 2);
        Assert.Contains(conversion.Report.Mappings, mapping => mapping.Feature == "speaker-notes" &&
            mapping.Status == OdfConversionMappingStatus.Skipped && mapping.Count == 1);
    }

    [Fact]
    public void OdpToPowerPoint_Reports_Formatting_And_Notes_Disabled_By_Options() {
        OdpPresentation source = OdpPresentation.Create();
        OdpSlide slide = source.AddSlide("Styled");
        OdpTextBox text = slide.AddTextBox(OdfRect.FromCentimeters(1, 1, 8, 2));
        text.FillColor = OdfColor.Parse("#112233");
        text.AddParagraph().AddRun("Styled").Bold = true;
        slide.GetOrCreateSpeakerNotes().AddParagraph("Private note");

        OdfConversionResult<PowerPointPresentation> conversion = source.ToPowerPointPresentationResult(
            new PowerPointOpenDocumentConversionOptions {
                IncludeBasicFormatting = false,
                IncludeSpeakerNotes = false
            });
        using PowerPointPresentation target = conversion.Value;
        OdfConversionReport report = conversion.Report;

        Assert.Contains(report.Mappings, mapping => mapping.Feature == "basic-formatting" &&
            mapping.Status == OdfConversionMappingStatus.Skipped && mapping.Count >= 2);
        Assert.Contains(report.Mappings, mapping => mapping.Feature == "speaker-notes" &&
            mapping.Status == OdfConversionMappingStatus.Skipped && mapping.Count == 1);
    }

    [Fact]
    public void OdpParagraphTextDecorationsFlowToPlainTextRunsAndHyperlinks() {
        OdpPresentation source = OdpPresentation.Create();
        OdpParagraph paragraph = source.AddSlide("Styled")
            .AddTextBox(OdfRect.FromCentimeters(1, 1, 8, 2)).AddParagraph();
        paragraph.Underline = true;
        paragraph.StrikeThrough = true;
        paragraph.BackgroundColor = OdfColor.Parse("#AABBCC");
        paragraph.AddText("Plain");
        paragraph.AddHyperlink("Link", "#slide-2");
        source.AddSlide("Target");

        OdfConversionResult<PowerPointPresentation> conversion = source.ToPowerPointPresentationResult();
        using PowerPointPresentation target = conversion.Value;
        PowerPointTextRun[] runs = target.Slides[0].TextBoxes.Single().Paragraphs.Single().Runs
            .Where(run => run.Text.Length > 0).ToArray();

        Assert.Equal(2, runs.Length);
        Assert.All(runs, run => Assert.True(run.Underline));
        Assert.All(runs, run => Assert.True(run.Strikethrough));
        Assert.All(runs, run => Assert.Equal("AABBCC", run.HighlightColor));
    }

    [Fact]
    public void OdtRelativeParagraphMeasurementsAreOmittedAndReportedInsteadOfThrowing() {
        OdtDocument source = OdtDocument.Create();
        OdtParagraph paragraph = source.AddParagraph("Body");
        paragraph.IndentStart = OdfLength.Parse("10%");
        paragraph.FontSize = OdfLength.Parse("120%");
        paragraph.Underline = true;
        paragraph.StrikeThrough = true;
        paragraph.TextBackgroundColor = OdfColor.Parse("#FFFF00");

        OdfConversionResult<WordDocument> conversion = source.ToWordDocumentResult();
        using WordDocument target = conversion.Value;

        WordParagraph run = Assert.Single(target.Paragraphs.Single().GetRuns());
        Assert.Contains("Body", run.Text, StringComparison.Ordinal);
        Assert.Equal(WordUnderlineStyle.Single, run.Underline);
        Assert.True(run.Strike);
        Assert.Equal(WordHighlightColor.Yellow, run.Highlight);
        Assert.Contains(conversion.Report.Mappings, mapping => mapping.Feature == "relative-measurements"
            && mapping.Status == OdfConversionMappingStatus.Unsupported && mapping.Count == 2);
        Assert.Throws<OdfConversionLossException>(() => conversion.Report.RequireNoSkippedOrUnsupported());
    }

    [Fact]
    public void OdsRelativeTextMeasurementsAreOmittedAndReportedInsteadOfThrowing() {
        OdsDocument source = OdsDocument.Create();
        OdsCell cell = source.AddSheet("Data").Cell(0, 0);
        cell.SetString("Body");
        cell.FontSize = OdfLength.Parse("120%");

        OdfConversionResult<ExcelDocument> conversion = source.ToExcelDocumentResult();
        using ExcelDocument target = conversion.Value;

        Assert.Equal("Body", target.CreateInspectionSnapshot().Worksheets.Single().Cells.Single().Value);
        Assert.Contains(conversion.Report.Mappings, mapping => mapping.Feature == "relative-measurements"
            && mapping.Status == OdfConversionMappingStatus.Unsupported && mapping.Count == 1);
        Assert.Throws<OdfConversionLossException>(() => conversion.Report.RequireNoSkippedOrUnsupported());
    }

    [Fact]
    public void OdpRelativeShapeGeometryIsOmittedAndReportedInsteadOfThrowing() {
        OdpPresentation source = OdpPresentation.Create();
        OdpTextBox textBox = source.AddSlide("Relative").AddTextBox(OdfRect.FromCentimeters(1, 1, 8, 2), "Body");
        textBox.Bounds = new OdfRect(OdfLength.Parse("10%"), OdfLength.Centimeters(1),
            OdfLength.Centimeters(8), OdfLength.Centimeters(2));

        OdfConversionResult<PowerPointPresentation> conversion = source.ToPowerPointPresentationResult();
        using PowerPointPresentation target = conversion.Value;

        Assert.Empty(target.Slides.Single().TextBoxes);
        Assert.Contains(conversion.Report.Mappings, mapping => mapping.Feature == "relative-measurements"
            && mapping.Status == OdfConversionMappingStatus.Unsupported && mapping.Count == 1);
        Assert.Contains(conversion.Report.Mappings, mapping => mapping.Feature == "shapes"
            && mapping.Status == OdfConversionMappingStatus.Unsupported && mapping.Count == 1);
        Assert.Throws<OdfConversionLossException>(() => conversion.Report.RequireNoSkippedOrUnsupported());
    }

    [Fact]
    public void WordToOdtReportsFlattenedNestedListLevels() {
        using WordDocument source = WordDocument.Create();
        WordList list = source.AddListNumbered();
        list.AddItem("Parent");
        list.AddItem("Child", 1);

        OdfConversionResult<OdtDocument> conversion = source.ToOpenDocumentResult();
        OdtDocument target = conversion.Value;

        Assert.Contains(conversion.Report.Mappings, mapping => mapping.Feature == "list-levels" &&
            mapping.Status == OdfConversionMappingStatus.Approximated && mapping.Count == 1);
    }

    [Fact]
    public void OdtToWordPreservesHeaderInlineFormattingAndImages() {
        OdtDocument source = OdtDocument.Create();
        OdtParagraph header = source.PageLayout.Header.AddParagraph();
        header.AddSpan("Styled").Bold = true;
        header.AddImage(TinyPng, "header.png", OdfLength.Centimeters(1), OdfLength.Centimeters(1));

        OdfConversionResult<WordDocument> conversion = source.ToWordDocumentResult();
        using WordDocument target = conversion.Value;
        WordParagraph paragraph = Assert.Single(target.Header!.Default!.Paragraphs, item => item.Text.Contains("Styled", StringComparison.Ordinal));
        WordParagraph[] runs = paragraph.GetRuns().ToArray();

        Assert.Contains(runs, run => run.Text == "Styled" && run.Bold);
        Assert.Contains(runs, run => run.IsImage);
        Assert.DoesNotContain(conversion.Report.Mappings, mapping => mapping.Feature == "images" &&
            mapping.Status == OdfConversionMappingStatus.Skipped);
    }

    [Fact]
    public void OdsConvertedAnnotationsAndNamedRangesPassStrictSkippedOrUnsupportedPolicy() {
        OdsDocument source = OdsDocument.Create();
        source.AddSheet("Data").Cell(0, 0).AddAnnotation("Review", "Alice");
        source.AddNamedRange("Input", "$'Data'.$A$1");

        OdfConversionResult<ExcelDocument> conversion = source.ToExcelDocumentResult(
            new ExcelOpenDocumentConversionOptions {
                LossPolicy = OdfConversionLossPolicy.ThrowOnSkippedOrUnsupported
            });
        using ExcelDocument target = conversion.Value;

        Assert.False(conversion.Report.HasSkippedOrUnsupported, Describe(conversion.Report));
        conversion.Report.RequireNoSkippedOrUnsupported();
        Assert.Single(target.Sheets.Single().GetComments());
        Assert.Contains(target.CreateInspectionSnapshot().NamedRanges, name => name.Name == "Input");
    }

    [Fact]
    public void OdpConvertedLinksNotesAndMasterPassStrictSkippedOrUnsupportedPolicy() {
        OdpPresentation source = OdpPresentation.Create();
        OdpMasterPage master = source.AddMasterPage("Brand");
        OdpSlide slide = source.AddSlide("Source");
        slide.MasterPageName = master.Name;
        slide.AddTextBox(OdfRect.FromCentimeters(1, 1, 8, 2))
            .AddParagraph().AddHyperlink("Web", "https://example.test/");
        slide.GetOrCreateSpeakerNotes().AddParagraph("Presenter note");

        OdfConversionResult<PowerPointPresentation> conversion = source.ToPowerPointPresentationResult(
            new PowerPointOpenDocumentConversionOptions {
                LossPolicy = OdfConversionLossPolicy.ThrowOnSkippedOrUnsupported
            });
        using PowerPointPresentation target = conversion.Value;

        Assert.False(conversion.Report.HasSkippedOrUnsupported, Describe(conversion.Report));
        conversion.Report.RequireNoSkippedOrUnsupported();
        Assert.Equal("Presenter note", target.Slides.Single().GetSpeakerNotesText());
        Assert.Equal("https://example.test/", target.Slides.Single().TextBoxes.Single()
            .Paragraphs.Single().Runs.Single().Hyperlink!.ToString());
    }

    [Fact]
    public void OdpInternalSlideLinkCreatesAnInternalPowerPointRelationship() {
        OdpPresentation source = OdpPresentation.Create();
        source.AddSlide("First").AddTextBox(OdfRect.FromCentimeters(1, 1, 8, 2))
            .AddParagraph().AddHyperlink("Next", "#slide-2");
        source.AddSlide("Second");

        OdfConversionResult<PowerPointPresentation> conversion = source.ToPowerPointPresentationResult();
        using PowerPointPresentation target = conversion.Value;

        Assert.Equal("#slide-2", target.Slides[0].TextBoxes.Single().Paragraphs.Single()
            .Runs.Single().Hyperlink!.ToString());
        byte[] package = target.ToBytes();
        string slideXml = ReadFirstPackageEntry(package, name =>
            name.StartsWith("ppt/slides/slide", StringComparison.Ordinal)
            && name.EndsWith(".xml", StringComparison.Ordinal));
        string relationships = ReadFirstPackageEntry(package, name =>
            name.StartsWith("ppt/slides/_rels/slide", StringComparison.Ordinal)
            && name.EndsWith(".xml.rels", StringComparison.Ordinal));
        Assert.Contains("ppaction://hlinksldjump", slideXml, StringComparison.Ordinal);
        Assert.DoesNotContain("TargetMode=\"External\"", relationships, StringComparison.Ordinal);
    }

    [Fact]
    public void OdpSlideNameFragmentCreatesAnInternalPowerPointRelationship() {
        OdpPresentation source = OdpPresentation.Create();
        source.AddSlide("First").AddTextBox(OdfRect.FromCentimeters(1, 1, 8, 2))
            .AddParagraph().AddHyperlink("Agenda", "#Agenda");
        source.AddSlide("Agenda");

        OdfConversionResult<PowerPointPresentation> conversion = source.ToPowerPointPresentationResult();
        using PowerPointPresentation target = conversion.Value;

        byte[] package = target.ToBytes();
        string slideXml = ReadFirstPackageEntry(package, name =>
            name.StartsWith("ppt/slides/slide", StringComparison.Ordinal)
            && name.EndsWith(".xml", StringComparison.Ordinal));
        string relationships = ReadFirstPackageEntry(package, name =>
            name.StartsWith("ppt/slides/_rels/slide", StringComparison.Ordinal)
            && name.EndsWith(".xml.rels", StringComparison.Ordinal));
        Assert.Contains("ppaction://hlinksldjump", slideXml, StringComparison.Ordinal);
        Assert.DoesNotContain("TargetMode=\"External\"", relationships, StringComparison.Ordinal);
    }

    [Fact]
    public void OdpUnknownSlideFragmentIsOmittedAndReportedUnsupported() {
        OdpPresentation source = OdpPresentation.Create();
        source.AddSlide("Only").AddTextBox(OdfRect.FromCentimeters(1, 1, 8, 2))
            .AddParagraph().AddHyperlink("Missing", "#Missing");

        OdfConversionResult<PowerPointPresentation> conversion = source.ToPowerPointPresentationResult();
        using PowerPointPresentation target = conversion.Value;

        string relationships = ReadFirstPackageEntry(target.ToBytes(), name =>
            name.StartsWith("ppt/slides/_rels/slide", StringComparison.Ordinal)
            && name.EndsWith(".xml.rels", StringComparison.Ordinal));
        Assert.DoesNotContain("TargetMode=\"External\"", relationships, StringComparison.Ordinal);
        Assert.Contains(conversion.Report.Mappings, mapping => mapping.Feature == "hyperlinks"
            && mapping.Status == OdfConversionMappingStatus.Unsupported && mapping.Count == 1);
        Assert.DoesNotContain(conversion.Report.Mappings, mapping => mapping.Feature == "hyperlinks"
            && mapping.Status == OdfConversionMappingStatus.Converted);
        Assert.Throws<OdfConversionLossException>(() => conversion.Report.RequireNoSkippedOrUnsupported());
    }

    [Fact]
    public void OdtInternalLinkDoesNotMaskUnsupportedExternalImage() {
        OdtDocument template = OdtDocument.Create();
        OdtParagraph paragraph = template.AddParagraph();
        paragraph.AddHyperlink("Internal", "#target");
        paragraph.AddImage(TinyPng, "pixel.png", OdfLength.Centimeters(1), OdfLength.Centimeters(1));
        byte[] package = RewriteXmlEntry(template.ToBytes(), "content.xml", document =>
            document.Descendants().Single(element => element.Name.LocalName == "image")
                .SetAttributeValue(XName.Get("href", "http://www.w3.org/1999/xlink"), "https://example.test/external.png"));
        OdtDocument source = OdtDocument.Load(new MemoryStream(package));

        OdfConversionResult<WordDocument> conversion = source.ToWordDocumentResult();
        using WordDocument target = conversion.Value;

        Assert.Contains(conversion.Report.Mappings, mapping => mapping.Feature == "source-external-links" &&
            mapping.Status == OdfConversionMappingStatus.Unsupported && mapping.Count == 1);
        Assert.Throws<OdfConversionLossException>(() => conversion.Report.RequireNoSkippedOrUnsupported());
    }

    [Fact]
    public void OdsInternalLinkDoesNotMaskUnsupportedExternalDrawingResource() {
        OdsDocument template = OdsDocument.Create();
        OdsCell cell = template.AddSheet("Links").Cell(0, 0);
        cell.SetString("Internal");
        cell.SetHyperlink("Internal", "#$'Links'.A1");
        byte[] package = RewriteXmlEntry(template.ToBytes(), "content.xml", document => {
            XNamespace draw = "urn:oasis:names:tc:opendocument:xmlns:drawing:1.0";
            XNamespace xlink = "http://www.w3.org/1999/xlink";
            document.Descendants().First(element => element.Name.LocalName == "table-cell").Add(
                new XElement(draw + "frame",
                    new XElement(draw + "image", new XAttribute(xlink + "href", "https://example.test/external.png"))));
        });
        OdsDocument source = OdsDocument.Load(new MemoryStream(package));

        OdfConversionResult<ExcelDocument> conversion = source.ToExcelDocumentResult();
        using ExcelDocument target = conversion.Value;

        Assert.Contains(conversion.Report.Mappings, mapping => mapping.Feature == "source-external-links" &&
            mapping.Status == OdfConversionMappingStatus.Unsupported && mapping.Count == 1);
        Assert.Throws<OdfConversionLossException>(() => conversion.Report.RequireNoSkippedOrUnsupported());
    }

    [Fact]
    public void OdpInternalLinkDoesNotMaskUnsupportedExternalImage() {
        OdpPresentation template = OdpPresentation.Create();
        template.AddSlide("First").AddTextBox(OdfRect.FromCentimeters(1, 1, 8, 2))
            .AddParagraph().AddHyperlink("Next", "#slide-2");
        OdpSlide second = template.AddSlide("Second");
        second.AddImage(TinyPng, "pixel.png", OdfRect.FromCentimeters(1, 1, 2, 2));
        byte[] package = RewriteXmlEntry(template.ToBytes(), "content.xml", document =>
            document.Descendants().Single(element => element.Name.LocalName == "image")
                .SetAttributeValue(XName.Get("href", "http://www.w3.org/1999/xlink"), "https://example.test/external.png"));
        OdpPresentation source = OdpPresentation.Load(new MemoryStream(package));

        OdfConversionResult<PowerPointPresentation> conversion = source.ToPowerPointPresentationResult();
        using PowerPointPresentation target = conversion.Value;

        Assert.Contains(conversion.Report.Mappings, mapping => mapping.Feature == "source-external-links" &&
            mapping.Status == OdfConversionMappingStatus.Unsupported && mapping.Count == 1);
        Assert.Throws<OdfConversionLossException>(() => conversion.Report.RequireNoSkippedOrUnsupported());
    }

    [Fact]
    public void UnreadableAuxiliaryOdfXmlIsPropagatedIntoStrictConversionReport() {
        OdsDocument template = OdsDocument.Create();
        template.AddSheet("Data").Cell(0, 0).SetString("Visible");
        byte[] package = OdfTestPackageRewriter.Rewrite(template.ToBytes(), new[] {
            new OdfTestPackageEntry("settings.xml", Encoding.UTF8.GetBytes("<broken"))
        });
        OdsDocument source = OdsDocument.Load(new MemoryStream(package));

        OdfConversionResult<ExcelDocument> conversion = source.ToExcelDocumentResult();
        using ExcelDocument target = conversion.Value;

        Assert.Contains(conversion.Report.Mappings, mapping =>
            mapping.Feature == "source-inspection"
            && mapping.Status == OdfConversionMappingStatus.Unsupported);
        Assert.Throws<OdfConversionLossException>(() => conversion.Report.RequireNoSkippedOrUnsupported());
    }

    [Fact]
    public void OdpToPowerPointReportsFlattenedListsButPreservesMixedRuns() {
        OdpPresentation source = OdpPresentation.Create();
        OdpTextBox textBox = source.AddSlide("Text").AddTextBox(
            OdfRect.FromCentimeters(1, 1, 8, 4), null, "Content");
        OdpParagraph mixed = textBox.AddParagraph("Plain ");
        mixed.AddRun("Bold").Bold = true;
        textBox.AddList().AddItem("Bullet");

        OdfConversionResult<PowerPointPresentation> conversion = source.ToPowerPointPresentationResult();
        using PowerPointPresentation target = conversion.Value;

        Assert.Contains(conversion.Report.Mappings, mapping => mapping.Feature == "text-lists" &&
            mapping.Status == OdfConversionMappingStatus.Approximated && mapping.Count == 1);
        Assert.DoesNotContain(conversion.Report.Mappings, mapping => mapping.Feature == "inline-formatting");
        var runs = target.Slides.Single().TextBoxes.Single().Paragraphs[0].Runs;
        Assert.Equal(new[] { "Plain ", "Bold" }, runs.Select(run => run.Text));
        Assert.True(runs[1].Bold);
    }

    [Fact]
    public void OdtToWordToleratesMissingStylesAndReportsTableCellImages() {
        OdtDocument template = OdtDocument.Create();
        template.AddParagraph("Minimal");
        template.AddTable(1, 1, "Media").Cell(0, 0).Paragraphs[0].AddImage(TinyPng, "cell.png",
            OdfLength.Centimeters(1), OdfLength.Centimeters(1));
        OdtDocument source = OdtDocument.Load(new MemoryStream(RemovePackageEntry(template.ToBytes(), "styles.xml")));

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
        ExcelSheet sheet = source.AddWorksheet("Data");
        source.AddWorksheet("Other, Sheet").CellAt(1, 1).SetValue(1);
        sheet.SetHyperlink(1, 1, "https://example.com", "42");
        sheet.CellAt(1, 1).SetValue(42);
        sheet.CellAt(1, 2).SetFormula("IF(A1=42,\"x,y\",\"other\")");
        sheet.CellAt(1, 3).SetFormula("SUM('Other, Sheet'!A1,A1)");

        OdfConversionResult<OdsDocument> conversion = source.ToOpenDocumentResult();
        OdsDocument target = conversion.Value;
        OdsSheet converted = target.GetSheet("Data")!;

        Assert.Equal(OdsCellValueKind.Number, converted.GetValue(0, 0).Kind);
        Assert.Equal(42D, converted.GetValue(0, 0).AsDouble());
        Assert.Equal("https://example.com", converted.RowRuns[0].CellRuns[0].HyperlinkHref);
        Assert.Equal("of:=IF([.A1]=42;\"x,y\";\"other\")", converted.GetFormula(0, 1));
        Assert.Equal("of:=SUM([$'Other, Sheet'.A1];[.A1])", converted.GetFormula(0, 2));
    }

    [Fact]
    public void OdsToExcelCreatesInternalLinksWithoutLosingTypedValues() {
        OdsDocument source = OdsDocument.Create();
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
    public void OdsToExcelPreservesNamedRangeHyperlinkFragments() {
        OdsDocument source = OdsDocument.Create();
        source.AddSheet("Target").Cell(0, 0).SetString("Destination");
        source.AddNamedRange("MyRange", "$'Target'.A1");
        OdsCell linked = source.AddSheet("Links").Cell(0, 0);
        linked.SetHyperlink("Go", "#MyRange");

        OdfConversionResult<ExcelDocument> conversion = source.ToExcelDocumentResult();
        using ExcelDocument target = conversion.Value;
        ExcelCellSnapshot cell = Assert.Single(target.CreateInspectionSnapshot().Worksheets
            .Single(sheet => sheet.Name == "Links").Cells);

        Assert.NotNull(cell.Hyperlink);
        Assert.False(cell.Hyperlink!.IsExternal);
        Assert.Equal("MyRange", cell.Hyperlink.Target);
        Assert.DoesNotContain(conversion.Report.Mappings, mapping => mapping.Feature == "hyperlinks"
            && mapping.Status == OdfConversionMappingStatus.Unsupported);
    }

    [Fact]
    public void OdsUnknownInternalHyperlinkFragmentIsOmittedAndReportedUnsupported() {
        OdsDocument source = OdsDocument.Create();
        OdsCell linked = source.AddSheet("Links").Cell(0, 0);
        linked.SetHyperlink("Missing", "#Missing");

        OdfConversionResult<ExcelDocument> conversion = source.ToExcelDocumentResult();
        using ExcelDocument target = conversion.Value;

        ExcelCellSnapshot cell = Assert.Single(target.CreateInspectionSnapshot().Worksheets.Single().Cells);
        Assert.Null(cell.Hyperlink);
        Assert.Contains(conversion.Report.Mappings, mapping => mapping.Feature == "hyperlinks"
            && mapping.Status == OdfConversionMappingStatus.Unsupported && mapping.Count == 1);
        Assert.Throws<OdfConversionLossException>(() => conversion.Report.RequireNoSkippedOrUnsupported());
    }

    [Fact]
    public void ExcelToOdsConvertsLowercaseFormulaReferences() {
        using ExcelDocument source = ExcelDocument.Create(new MemoryStream());
        ExcelSheet sheet = source.AddWorksheet("Data");
        sheet.CellAt(1, 1).SetValue(1);
        sheet.CellAt(1, 2).SetFormula("sum(a1,'Data'!a1)");
        source.SetNamedRange("A1_total", "'Data'!$A$1", save: false);
        sheet.CellAt(1, 3).SetFormula("SUM(A1_total,a1)");

        OdsDocument target = source.ToOpenDocument();

        Assert.Equal("of:=sum([.A1];[$'Data'.A1])", target.GetSheet("Data")!.GetFormula(0, 1));
        Assert.Equal("of:=SUM(A1_total;[.A1])", target.GetSheet("Data")!.GetFormula(0, 2));
    }

    [Fact]
    public void OdsToExcelPreservesRelativeExternalLinksAndMissingStyles() {
        OdsDocument template = OdsDocument.Create();
        OdsCell linked = template.AddSheet("Links").Cell(0, 0);
        linked.SetString("Docs");
        linked.SetHyperlink("Docs", "docs/page.html");
        OdsDocument source = OdsDocument.Load(new MemoryStream(RemovePackageEntry(template.ToBytes(), "styles.xml")));

        using ExcelDocument target = source.ToExcelDocument();
        ExcelCellSnapshot cell = Assert.Single(target.CreateInspectionSnapshot().Worksheets.Single().Cells);

        Assert.NotNull(cell.Hyperlink);
        Assert.True(cell.Hyperlink!.IsExternal);
        Assert.Equal("docs/page.html", cell.Hyperlink.Target);
    }

    [Fact]
    public void OdpToPowerPointPreservesStyledParagraphsNormalizedImagesAndMasterBackground() {
        OdpPresentation template = OdpPresentation.Create();
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
        OdpPresentation source = OdpPresentation.Load(new MemoryStream(package));

        using PowerPointPresentation target = source.ToPowerPointPresentation();
        PowerPointSlide converted = Assert.Single(target.Slides);

        Assert.Equal(new[] { "First", "Second" }, converted.TextBoxes.Single().Paragraphs.Select(paragraph => paragraph.Text));
        Assert.Single(converted.Pictures);
        Assert.Equal("445566", converted.GetBackground().Color);
    }

    [Fact]
    public void OdpToPowerPointToleratesMissingStylesPart() {
        OdpPresentation template = OdpPresentation.Create();
        template.AddSlide("Minimal").AddTextBox(OdfRect.FromCentimeters(1, 1, 8, 2), "Text");
        OdpPresentation source = OdpPresentation.Load(new MemoryStream(RemovePackageEntry(template.ToBytes(), "styles.xml")));

        using PowerPointPresentation target = source.ToPowerPointPresentation();

        Assert.Single(target.Slides);
        Assert.Contains("Text", target.Slides.Single().TextBoxes.Single().Text, StringComparison.Ordinal);
    }

    [Fact]
    public void OdsToExcelPrefersContentScopedDuplicateDataStyle() {
        OdsDocument template = OdsDocument.Create();
        template.AddNumberStyle("Amount", 2);
        template.AddSheet("Data").Cell(0, 0).SetNumber(12.5);
        byte[] package = RewriteXmlEntry(template.ToBytes(), "styles.xml", document => {
            XNamespace office = "urn:oasis:names:tc:opendocument:xmlns:office:1.0";
            XNamespace style = "urn:oasis:names:tc:opendocument:xmlns:style:1.0";
            XNamespace number = "urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0";
            document.Root!.Element(office + "styles")!.Add(
                new XElement(number + "percentage-style", new XAttribute(style + "name", "Amount")));
        });
        OdsDocument source = OdsDocument.Load(new MemoryStream(package));

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
        OdtDocument target = conversion.Value;
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

    private static string ReadFirstPackageEntry(byte[] packageBytes, Func<string, bool> predicate) {
        using var stream = new MemoryStream(packageBytes, writable: false);
        using var package = new System.IO.Compression.ZipArchive(stream, System.IO.Compression.ZipArchiveMode.Read);
        System.IO.Compression.ZipArchiveEntry entry = package.Entries
            .OrderBy(item => item.FullName, StringComparer.Ordinal)
            .FirstOrDefault(item => predicate(item.FullName))
            ?? throw new InvalidDataException("Matching package entry was not found.");
        using Stream input = entry.Open();
        using var reader = new StreamReader(input, Encoding.UTF8);
        return reader.ReadToEnd();
    }

    private static string Describe(OdfConversionReport report) => string.Join("; ", report.Mappings.Select(mapping =>
        mapping.Feature + "=" + mapping.Status + "(" + mapping.Count + "):" + mapping.Message));
}
