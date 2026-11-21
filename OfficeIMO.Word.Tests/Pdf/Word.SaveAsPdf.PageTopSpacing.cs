using W = DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using Xunit;

namespace OfficeIMO.Tests;

// Expected positions follow Microsoft Word 16 PDF exports of the same documents:
// space before is kept at the start of the document and of a new-page section,
// and dropped at the top of a page after a page break or natural flow.
public partial class Word {
    [Fact]
    public void SaveAsPdf_HeadingWithoutSpacingBeforeInheritsDocumentStyles() {
        double inherited = MeasureHeadingGap(headingBefore: null, defaultBefore: null);
        double explicitZero = MeasureHeadingGap(headingBefore: "0", defaultBefore: null);
        Assert.Equal(explicitZero, inherited, 1);

        // A document default before value reaches the heading through Normal.
        double fromDefaults = MeasureHeadingGap(headingBefore: null, defaultBefore: "240");
        double explicitTwelve = MeasureHeadingGap(headingBefore: "240", defaultBefore: null);
        Assert.Equal(explicitTwelve, fromDefaults, 1);
    }

    [Theory]
    [InlineData("pageBreakRun", "Heading1")]
    [InlineData("pageBreakBefore", "Heading1")]
    [InlineData("pageBreakBefore", "Heading2")]
    [InlineData("flow", "Heading1")]
    [InlineData("flow", "Normal")]
    public void SaveAsPdf_SpacingBeforeIsDroppedAtTopOfPage(string breakKind, string styleId) {
        (int page, double spaced) = MeasurePageTopTarget(breakKind, styleId, before: "360");
        (int zeroPage, double zero) = MeasurePageTopTarget(breakKind, styleId, before: "0");
        Assert.Equal(2, page);
        Assert.Equal(2, zeroPage);
        Assert.Equal(zero, spaced, 1);
    }

    [Theory]
    [InlineData("Heading1")]
    [InlineData("Normal")]
    public void SaveAsPdf_SpacingBeforeIsKeptAtDocumentStart(string styleId) {
        using WordDocument document = WordDocument.Create();
        SetPageTopStyleSpacing(document, styleId, "360");
        document.AddParagraph("Target").Style = ParsePageTopStyle(styleId);
        double spaced = FindPageTopBaseline(RenderPageTopPdf(document), "Target").Baseline;

        using WordDocument reference = WordDocument.Create();
        SetPageTopStyleSpacing(reference, styleId, "0");
        reference.AddParagraph("Target").Style = ParsePageTopStyle(styleId);
        double zero = FindPageTopBaseline(RenderPageTopPdf(reference), "Target").Baseline;
        Assert.Equal(18D, zero - spaced, 1);
    }

    [Theory]
    [InlineData(0, 18D)]
    [InlineData(160, 10D)]
    [InlineData(480, 0D)]
    public void SaveAsPdf_NextPageSectionKeepsSpacingBeforeCollapsedWithPreviousSection(int previousAfter, double expectedGap) {
        double spaced = MeasureSectionStartTarget(previousAfter, "360", "Normal");
        double zero = MeasureSectionStartTarget(previousAfter, "0", "Normal");
        Assert.Equal(expectedGap, zero - spaced, 1);
        double heading = MeasureSectionStartTarget(previousAfter, "360", "Heading1");
        double headingZero = MeasureSectionStartTarget(previousAfter, "0", "Heading1");
        Assert.Equal(expectedGap, headingZero - heading, 1);
    }

    [Fact]
    public void SaveAsPdf_NextPageSectionCollapsesWithEmptySectionMarkSpacing() {
        // The empty paragraph that holds the section break has no space after, so the
        // next section's first paragraph keeps its full space before.
        double spaced = MeasureEmptySectionMarkTarget("360");
        double zero = MeasureEmptySectionMarkTarget("0");
        Assert.Equal(18D, zero - spaced, 1);
    }

    private static double MeasureEmptySectionMarkTarget(string before) {
        using WordDocument document = WordDocument.Create();
        document.AddParagraph("First");
        WordSection second = document.AddSection(WordSectionBreakType.NextPage);
        var markProperties = (W.ParagraphProperties)document.Sections[0]._sectionProperties.Parent!;
        markProperties.AddChild(new W.SpacingBetweenLines { After = "0" }, true);
        ApplyPageTopSpacing(document, second.AddParagraph("Target"), "Normal", before);
        (int page, double baseline) = FindPageTopBaseline(RenderPageTopPdf(document), "Target");
        Assert.Equal(2, page);
        return baseline;
    }

    [Theory]
    [InlineData(WordCompatibilityMode.Word2013, 0D)]
    [InlineData(WordCompatibilityMode.Word2010, 10D)]
    [InlineData(WordCompatibilityMode.Word2007, 10D)]
    public void SaveAsPdf_PageBreakBeforeSpacingFollowsCompatibilityMode(WordCompatibilityMode mode, double expectedGap) {
        double spaced = MeasurePageBreakBeforeTarget(mode, "360");
        double zero = MeasurePageBreakBeforeTarget(mode, "0");
        Assert.Equal(expectedGap, zero - spaced, 1);
    }

    [Theory]
    // Without facing pages, a next-page section never adds a page, and odd/even
    // sections with a restart move the number instead of adding a page.
    [InlineData(WordSectionBreakType.NextPage, 1, 1, false, 2)]
    [InlineData(WordSectionBreakType.NextPage, 1, 3, false, 2)]
    [InlineData(WordSectionBreakType.NextPage, 2, 2, false, 3)]
    [InlineData(WordSectionBreakType.OddPage, 1, 2, false, 2)]
    [InlineData(WordSectionBreakType.OddPage, 1, 0, false, 3)]
    [InlineData(WordSectionBreakType.EvenPage, 2, 1, false, 3)]
    [InlineData(WordSectionBreakType.EvenPage, 2, 0, false, 4)]
    // Facing pages keep odd numbers on right-hand pages.
    [InlineData(WordSectionBreakType.NextPage, 1, 1, true, 3)]
    [InlineData(WordSectionBreakType.NextPage, 1, 2, true, 2)]
    [InlineData(WordSectionBreakType.OddPage, 1, 2, true, 3)]
    public void SaveAsPdf_SectionRestartAddsBlankPageOnlyForFacingPages(
        WordSectionBreakType breakType, int firstPages, int restart, bool mirrorMargins, int expectedPages) {
        using WordDocument document = WordDocument.Create();
        document.Settings.MirrorMargins = mirrorMargins;
        for (int page = 1; page <= firstPages; page++) document.AddParagraph("First" + page).PageBreakBeforeOverride = page > 1;
        WordSection second = document.AddSection(breakType);
        if (restart > 0) second.AddPageNumbering(restart);
        second.AddParagraph("SecondSection");
        using PdfPigDocument pdf = RenderPageTopPdf(document);
        Assert.Equal(expectedPages, pdf.NumberOfPages);
        Assert.Contains("SecondSection", pdf.GetPage(expectedPages).Text);
    }

    [Theory]
    [InlineData(0, 18D)]
    [InlineData(160, 18D)]
    [InlineData(480, 24D)]
    public void SaveAsPdf_ColumnBreakCarriesParagraphMarkIntoNextColumn(int breakAfter, double expectedSpacing) {
        // The paragraph mark after a column break takes an empty line at the top of
        // the next column; the following paragraph's spacing collapses against it.
        double spaced = MeasureColumnBreakTarget(breakAfter, 360);
        double markOnly = MeasureColumnBreakTarget(breakAfter, 0);
        double breakAtStart = MeasureColumnBreakTarget(breakAfter, 360, breakInsideTarget: true);
        Assert.Equal(expectedSpacing - breakAfter / 20D, markOnly - spaced, 1);
        Assert.True(breakAtStart - spaced > 10D, "The empty mark line must sit above the following paragraph.");
    }

    [Fact]
    public void SaveAsPdf_AutomaticParagraphSpacingUsesFourteenPoints() {
        using WordDocument document = WordDocument.Create();
        WordParagraph first = document.AddParagraph("First");
        SetPageTopAutomaticSpacing(first, before: true, after: true);
        WordParagraph second = document.AddParagraph("Second");
        SetPageTopAutomaticSpacing(second, before: true, after: false);
        second.LineSpacingAfter = global::OfficeIMO.OfficeLength.Points(0D);
        document.AddParagraph("Third").LineSpacingBefore = global::OfficeIMO.OfficeLength.Points(0D);

        using WordDocument reference = WordDocument.Create();
        WordParagraph referenceFirst = reference.AddParagraph("First");
        referenceFirst.LineSpacingBefore = global::OfficeIMO.OfficeLength.Points(0D);
        referenceFirst.LineSpacingAfter = global::OfficeIMO.OfficeLength.Points(14D);
        WordParagraph referenceSecond = reference.AddParagraph("Second");
        referenceSecond.LineSpacingBefore = global::OfficeIMO.OfficeLength.Points(14D);
        referenceSecond.LineSpacingAfter = global::OfficeIMO.OfficeLength.Points(0D);
        reference.AddParagraph("Third").LineSpacingBefore = global::OfficeIMO.OfficeLength.Points(0D);

        // Automatic spacing is 14 pt, omitted before the first paragraph, and
        // collapses with the neighbouring value like authored spacing.
        foreach (string token in new[] { "First", "Second", "Third" }) {
            Assert.Equal(FindPageTopBaseline(RenderPageTopPdf(reference), token).Baseline,
                FindPageTopBaseline(RenderPageTopPdf(document), token).Baseline, 1);
        }
    }

    private static void SetPageTopAutomaticSpacing(WordParagraph paragraph, bool before, bool after) {
        W.ParagraphProperties properties = paragraph._paragraph.ParagraphProperties ??= new W.ParagraphProperties();
        W.SpacingBetweenLines spacing = properties.SpacingBetweenLines ??= new W.SpacingBetweenLines();
        if (before) { spacing.Before = "100"; spacing.BeforeAutoSpacing = true; }
        if (after) { spacing.After = "100"; spacing.AfterAutoSpacing = true; }
    }

    private static double MeasureColumnBreakTarget(int breakAfter, int targetBefore, bool breakInsideTarget = false) {
        using WordDocument document = WordDocument.Create();
        document.Sections[0].ColumnCount = 2;
        WordParagraph first = document.AddParagraph("First");
        first.LineSpacingAfter = global::OfficeIMO.OfficeLength.Points(breakAfter / 20D);
        if (!breakInsideTarget) first.AddBreak(WordBreakType.Column);
        WordParagraph target = document.AddParagraph();
        if (breakInsideTarget) target.AddBreak(WordBreakType.Column);
        target.AddText("Target");
        target.LineSpacingBefore = global::OfficeIMO.OfficeLength.Points(targetBefore / 20D);
        document.AddParagraph("After");
        return FindPageTopBaseline(RenderPageTopPdf(document), "Target").Baseline;
    }

    private static double MeasureHeadingGap(string? headingBefore, string? defaultBefore) {
        using WordDocument document = WordDocument.Create();
        SetPageTopStyleSpacing(document, "Heading1", headingBefore);
        SetPageTopStyleSpacing(document, "Normal", null);
        W.SpacingBetweenLines defaults = GetPageTopDocumentDefaultSpacing(document);
        defaults.Before = defaultBefore;
        defaults.After = "160";
        document.AddParagraph("Body").Style = WordParagraphStyles.Normal;
        document.AddParagraph("Target").Style = WordParagraphStyles.Heading1;
        using PdfPigDocument pdf = RenderPageTopPdf(document);
        var words = pdf.GetPage(1).GetWords().ToList();
        // The gap from the body line keeps document-start spacing out of the comparison.
        return words.Single(word => word.Text == "Body").Letters[0].StartBaseLine.Y -
            words.Single(word => word.Text == "Target").Letters[0].StartBaseLine.Y;
    }

    private static (int Page, double Baseline) MeasurePageTopTarget(string breakKind, string styleId, string before) {
        using WordDocument document = WordDocument.Create();
        int fillers = breakKind == "flow" ? CountFirstPageFillers() : 1;
        for (int i = 0; i < fillers; i++) document.AddParagraph("Filler" + i);
        if (breakKind == "pageBreakRun") document.AddParagraph().AddBreak(WordBreakType.Page);
        WordParagraph target = document.AddParagraph("Target");
        ApplyPageTopSpacing(document, target, styleId, before);
        if (breakKind == "pageBreakBefore") target.PageBreakBeforeOverride = true;
        document.AddParagraph("After");
        return FindPageTopBaseline(RenderPageTopPdf(document), "Target");
    }

    private static int CountFirstPageFillers() {
        using WordDocument document = WordDocument.Create();
        for (int i = 0; i < 80; i++) document.AddParagraph("Filler" + i);
        using PdfPigDocument pdf = RenderPageTopPdf(document);
        return pdf.GetPage(1).GetWords().Count(word => word.Text.StartsWith("Filler", StringComparison.Ordinal));
    }

    private static double MeasureSectionStartTarget(int previousAfter, string before, string styleId) {
        using WordDocument document = WordDocument.Create();
        WordParagraph last = document.AddParagraph("First");
        last.LineSpacingAfter = global::OfficeIMO.OfficeLength.Points(previousAfter / 20D);
        WordSection second = document.AddSection(WordSectionBreakType.NextPage);
        ApplyPageTopSpacing(document, second.AddParagraph("Target"), styleId, before);
        (int page, double baseline) = FindPageTopBaseline(RenderPageTopPdf(document), "Target");
        Assert.Equal(2, page);
        return baseline;
    }

    private static double MeasurePageBreakBeforeTarget(WordCompatibilityMode mode, string before) {
        using WordDocument document = WordDocument.Create();
        document.CompatibilitySettings.CompatibilityMode = mode;
        document.AddParagraph("First");
        WordParagraph target = document.AddParagraph("Target");
        target.PageBreakBeforeOverride = true;
        target.LineSpacingBefore = global::OfficeIMO.OfficeLength.Points(int.Parse(before, System.Globalization.CultureInfo.InvariantCulture) / 20D);
        (int page, double baseline) = FindPageTopBaseline(RenderPageTopPdf(document), "Target");
        Assert.Equal(2, page);
        return baseline;
    }

    // Body cases use direct formatting so the surrounding Normal paragraphs keep their layout.
    private static void ApplyPageTopSpacing(WordDocument document, WordParagraph target, string styleId, string before) {
        target.Style = ParsePageTopStyle(styleId);
        if (styleId == "Normal") {
            target.LineSpacingBefore = global::OfficeIMO.OfficeLength.Points(int.Parse(before, System.Globalization.CultureInfo.InvariantCulture) / 20D);
        } else {
            SetPageTopStyleSpacing(document, styleId, before);
        }
    }

    private static WordParagraphStyles ParsePageTopStyle(string styleId) => styleId switch {
        "Heading1" => WordParagraphStyles.Heading1,
        "Heading2" => WordParagraphStyles.Heading2,
        _ => WordParagraphStyles.Normal
    };

    private static void SetPageTopStyleSpacing(WordDocument document, string styleId, string? before) {
        W.Styles styles = document._wordprocessingDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!;
        W.Style style = styles.Elements<W.Style>().Single(item => item.StyleId?.Value == styleId);
        W.StyleParagraphProperties properties = style.StyleParagraphProperties ??= new W.StyleParagraphProperties();
        W.SpacingBetweenLines spacing = properties.SpacingBetweenLines ??= new W.SpacingBetweenLines();
        spacing.Before = before;
        spacing.BeforeLines = null;
        spacing.BeforeAutoSpacing = null;
    }

    private static W.SpacingBetweenLines GetPageTopDocumentDefaultSpacing(WordDocument document) {
        W.Styles styles = document._wordprocessingDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!;
        W.DocDefaults defaults = styles.DocDefaults ??= new W.DocDefaults();
        W.ParagraphPropertiesDefault paragraphDefaults = defaults.ParagraphPropertiesDefault ??= new W.ParagraphPropertiesDefault();
        W.ParagraphPropertiesBaseStyle baseStyle = paragraphDefaults.ParagraphPropertiesBaseStyle ??= new W.ParagraphPropertiesBaseStyle();
        return baseStyle.SpacingBetweenLines ??= new W.SpacingBetweenLines();
    }

    private static PdfPigDocument RenderPageTopPdf(WordDocument document) {
        using WordDocument loaded = WordDocument.Load(new MemoryStream(document.ToBytes(WordFileFormat.Docx)));
        return PdfPigDocument.Open(loaded.ToPdf(new WordPdfExportOptions { IncludePageNumbers = false })
            .ToBytes(new OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true }));
    }

    private static (int Page, double Baseline) FindPageTopBaseline(PdfPigDocument pdf, string token) {
        using (pdf) {
            foreach (var page in pdf.GetPages()) {
                var word = page.GetWords().FirstOrDefault(item => item.Text == token);
                if (word != null) return (page.Number, word.Letters[0].StartBaseLine.Y);
            }
        }

        throw new InvalidOperationException("Could not find '" + token + "' in the rendered PDF.");
    }
}
