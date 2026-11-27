using W = DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using Xunit;

namespace OfficeIMO.Tests;

// Expected positions follow Microsoft Word 16 PDF exports of the same frames
// (Letter page, one-inch margins, footer and header 0.5 inch from the edge).
public partial class Word {
    private const double FramePageWidth = 612D;
    private const double FrameMargin = 72D;

    [Fact]
    public void SaveAsPdf_FooterPageNumberFrameSitsAtRightMarginBesideItsAnchor() {
        using WordDocument document = CreateFrameDocument();
        WordFooter footer = document.PageSetup.Footer;
        AddFrameParagraph(footer.AddParagraph("FRAMED"), "around", "text", "margin", xAlign: "right", y: 1);
        footer.AddParagraph("ANCHOR");
        using PdfPigDocument pdf = RenderFramePdf(document);
        var framed = FindFrameWord(pdf, "FRAMED");
        var anchor = FindFrameWord(pdf, "ANCHOR");
        Assert.Equal(FramePageWidth - FrameMargin, framed.BoundingBox.Right, 0);
        Assert.Equal(anchor.Letters[0].StartBaseLine.Y, framed.Letters[0].StartBaseLine.Y, 0);
        Assert.Equal(FrameMargin, anchor.BoundingBox.Left, 0);
    }

    [Fact]
    public void SaveAsPdf_FooterFrameOffsetsFromItsAnchorAndDoesNotGrowTheFooter() {
        double near = MeasureFooterFrame(y: 1, frameLast: false).Frame;
        double lower = MeasureFooterFrame(y: 200, frameLast: false).Frame;
        Assert.Equal(10D, near - lower, 0);

        // A frame after the last paragraph floats below it; the footer keeps its own height.
        (double lastFrame, double anchor) = MeasureFooterFrame(y: 1, frameLast: true);
        double withoutFrame = MeasureFooterWithoutFrame();
        Assert.Equal(withoutFrame, anchor, 1);
        Assert.True(anchor - lastFrame > 10D, "The trailing frame starts below the footer text.");
    }

    [Theory]
    [InlineData("footer")]
    [InlineData("header")]
    public void SaveAsPdf_PageAnchoredRunningFrameUsesPageCoordinates(string story) {
        using WordDocument document = CreateFrameDocument();
        WordHeaderFooter target = story == "footer" ? document.PageSetup.Footer : document.PageSetup.Header;
        AddFrameParagraph(target.AddParagraph("FRAMED"), "around", "page", "page", x: 5760, y: 1440);
        target.AddParagraph("ANCHOR");
        using PdfPigDocument pdf = RenderFramePdf(document);
        var framed = FindFrameWord(pdf, "FRAMED");
        Assert.Equal(288D, framed.BoundingBox.Left, 0);
        // The frame's first line starts one inch below the top edge of the page.
        Assert.InRange(792D - 72D - framed.Letters[0].StartBaseLine.Y, 8D, 14D);
    }

    [Theory]
    [InlineData(72D, "X", 540D)]
    [InlineData(144D, "X", 468D)]
    [InlineData(72D, "LONGFRAMETEXT", -1D)]
    public void SaveAsPdf_PageRelativeRightFrameSitsInTheRightMargin(double rightMargin, string text, double expectedLeft) {
        // Word places the frame against the text area inside the margin, and moves it
        // inward only when it would otherwise leave the page (-1: right edge at the page edge).
        using WordDocument document = CreateFrameDocument();
        document.Sections[0].PageSetup.Margins.Right = global::OfficeIMO.OfficeLength.Points(rightMargin);
        AddFrameParagraph(document.PageSetup.Footer.AddParagraph(text), "around", "text", "page", xAlign: "right", y: 1);
        document.PageSetup.Footer.AddParagraph("ANCHOR");
        using PdfPigDocument pdf = RenderFramePdf(document);
        var framed = FindFrameWord(pdf, text);
        if (expectedLeft >= 0D) Assert.Equal(expectedLeft, framed.BoundingBox.Left, 0);
        else Assert.Equal(FramePageWidth, framed.BoundingBox.Right, 0);
    }

    [Fact]
    public void SaveAsPdf_BodyFrameWrapsFollowingTextBesideIt() {
        using WordDocument document = CreateFrameDocument();
        AddFrameParagraph(document.AddParagraph("FRAMED"), "around", "text", "margin", x: 1, width: 2880, horizontalSpace: 180);
        document.AddParagraph(string.Join(" ", Enumerable.Range(1, 90).Select(index => "w" + index.ToString("00"))));
        using PdfPigDocument pdf = RenderFramePdf(document);
        var lines = pdf.GetPage(1).GetWords().Where(word => word.Text.StartsWith("w", StringComparison.Ordinal))
            .GroupBy(word => Math.Round(word.Letters[0].StartBaseLine.Y, 1)).OrderByDescending(group => group.Key)
            .Select(group => group.Min(word => word.BoundingBox.Left)).ToArray();
        // Two lines sit beside the one-line frame (144 pt wide plus 9 pt clearance); the next returns to the margin.
        Assert.Equal(FrameMargin + 144D + 9D, lines[0], 0);
        Assert.Equal(FrameMargin + 144D + 9D, lines[1], 0);
        Assert.Equal(FrameMargin, lines[2], 0);
    }

    [Fact]
    public void SaveAsPdf_FieldMarkupInsideUnknownMarkupDoesNotStopABodyFrame() {
        using WordDocument document = CreateFrameDocument();
        WordParagraph framedParagraph = document.AddParagraph("FRAMED");
        AddFrameParagraph(framedParagraph, "around", "text", "margin", x: 1, width: 2880, horizontalSpace: 180);
        // An unbalanced field begin the model doesn't read: the frame is still a frame.
        var wrapper = new DocumentFormat.OpenXml.OpenXmlUnknownElement("x", "wrapper", "urn:officeimo:tests:frame-wrapper");
        wrapper.InnerXml = "<x:group xmlns:x=\"urn:officeimo:tests:frame-wrapper\" " +
            "xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\">" +
            "<w:r><w:fldChar w:fldCharType=\"begin\"/></w:r><w:r><w:instrText> PAGE </w:instrText></w:r></x:group>";
        framedParagraph._paragraph.AppendChild(wrapper);
        document.AddParagraph(string.Join(" ", Enumerable.Range(1, 90).Select(index => "w" + index.ToString("00"))));
        using PdfPigDocument pdf = RenderFramePdf(document);
        double firstLine = pdf.GetPage(1).GetWords().Where(word => word.Text.StartsWith("w", StringComparison.Ordinal))
            .GroupBy(word => Math.Round(word.Letters[0].StartBaseLine.Y, 1)).OrderByDescending(group => group.Key)
            .Select(group => group.Min(word => word.BoundingBox.Left)).First();
        Assert.Equal(FrameMargin + 144D + 9D, firstLine, 0);
    }

    [Theory]
    [InlineData("notBeside")]
    [InlineData("none")]
    public void SaveAsPdf_BodyFrameWrapKindControlsFollowingText(string wrap) {
        using WordDocument document = CreateFrameDocument();
        // The frame's text sits at its right edge, so text running under the frame stays readable.
        WordParagraph framedParagraph = document.AddParagraph("FRAMED");
        framedParagraph.ParagraphAlignment = WordParagraphAlignment.Right;
        AddFrameParagraph(framedParagraph, wrap, "text", "margin", x: 1, width: 2880);
        document.AddParagraph("FOLLOWING text");
        using PdfPigDocument pdf = RenderFramePdf(document);
        var framed = FindFrameWord(pdf, "FRAMED");
        var following = FindFrameWord(pdf, "FOLLOWING");
        Assert.Equal(FrameMargin, following.BoundingBox.Left, 0);
        if (wrap == "notBeside") Assert.True(framed.Letters[0].StartBaseLine.Y - following.Letters[0].StartBaseLine.Y > 20D);
        else Assert.Equal(framed.Letters[0].StartBaseLine.Y, following.Letters[0].StartBaseLine.Y, 0);
    }

    [Theory]
    [InlineData(false, 0, 10D)]
    [InlineData(true, 0, 18D)]
    [InlineData(true, 160, 10D)]
    public void SaveAsPdf_NotBesideFrameSpacingFollowsWord(bool sectionStart, int previousAfter, double expectedGap) {
        // Mid-page, the next paragraph collapses its 18 pt space before with the frame's
        // 8 pt space after. At a new-page section start the frame passes the carried
        // section spacing on instead, and its own space after does not collapse.
        double spaced = MeasureNotBesideFrameTarget(sectionStart, previousAfter, 18D);
        double zero = MeasureNotBesideFrameTarget(sectionStart, previousAfter, 0D);
        Assert.Equal(expectedGap, zero - spaced, 1);
    }

    private static double MeasureNotBesideFrameTarget(bool sectionStart, int previousAfter, double before) {
        using WordDocument document = CreateFrameDocument();
        WordParagraph first = document.AddParagraph("First");
        first.LineSpacingAfter = global::OfficeIMO.OfficeLength.Points(previousAfter / 20D);
        WordSection section = sectionStart ? document.AddSection(WordSectionBreakType.NextPage) : document.Sections[0];
        WordParagraph framed = section.AddParagraph("FRAMED");
        framed.LineSpacingAfter = global::OfficeIMO.OfficeLength.Points(8D);
        AddFrameParagraph(framed, "notBeside", "text", "margin", x: 1, width: 2880);
        WordParagraph target = section.AddParagraph("TARGET");
        target.LineSpacingBefore = global::OfficeIMO.OfficeLength.Points(before);
        using PdfPigDocument pdf = RenderFramePdf(document);
        return FindFrameWord(pdf, "TARGET").Letters[0].StartBaseLine.Y;
    }

    [Fact]
    public void SaveAsPdf_ConsecutiveParagraphsWithOneFrameShareIt() {
        using WordDocument document = CreateFrameDocument();
        AddFrameParagraph(document.AddParagraph("FIRSTFRAME"), "around", "text", "margin", x: 1, width: 2880);
        AddFrameParagraph(document.AddParagraph("SECONDFRAME"), "around", "text", "margin", x: 1, width: 2880);
        document.AddParagraph("BESIDE text");
        using PdfPigDocument pdf = RenderFramePdf(document);
        Assert.Equal(FrameMargin, FindFrameWord(pdf, "SECONDFRAME").BoundingBox.Left, 0);
        Assert.Equal(FrameMargin + 144D, FindFrameWord(pdf, "BESIDE").BoundingBox.Left, 0);
    }

    private static (double Frame, double Anchor) MeasureFooterFrame(int y, bool frameLast) {
        using WordDocument document = CreateFrameDocument();
        WordFooter footer = document.PageSetup.Footer;
        if (frameLast) footer.AddParagraph("ANCHOR");
        AddFrameParagraph(footer.AddParagraph("FRAMED"), "around", "text", "margin", xAlign: "right", y: y);
        if (!frameLast) footer.AddParagraph("ANCHOR");
        using PdfPigDocument pdf = RenderFramePdf(document);
        return (FindFrameWord(pdf, "FRAMED").Letters[0].StartBaseLine.Y, FindFrameWord(pdf, "ANCHOR").Letters[0].StartBaseLine.Y);
    }

    private static double MeasureFooterWithoutFrame() {
        using WordDocument document = CreateFrameDocument();
        document.PageSetup.Footer.AddParagraph("ANCHOR");
        using PdfPigDocument pdf = RenderFramePdf(document);
        return FindFrameWord(pdf, "ANCHOR").Letters[0].StartBaseLine.Y;
    }

    private static WordDocument CreateFrameDocument() {
        WordDocument document = WordDocument.Create();
        document.Sections[0].PageSetup.Size.PageSize = WordPageSize.Letter;
        document.Sections[0].PageSetup.Margins.Top = document.Sections[0].PageSetup.Margins.Bottom = global::OfficeIMO.OfficeLength.Points(72D);
        document.Sections[0].PageSetup.Margins.Left = document.Sections[0].PageSetup.Margins.Right = global::OfficeIMO.OfficeLength.Points(72D);
        document.Sections[0].PageSetup.Margins.HeaderDistance = document.Sections[0].PageSetup.Margins.FooterDistance = global::OfficeIMO.OfficeLength.Points(36D);
        return document;
    }

    private static void AddFrameParagraph(WordParagraph paragraph, string wrap, string verticalAnchor, string horizontalAnchor,
        string? xAlign = null, int? x = null, int? y = null, int? width = null, int? horizontalSpace = null) {
        var frame = new W.FrameProperties {
            Wrap = wrap switch {
                "notBeside" => W.TextWrappingValues.NotBeside,
                "none" => W.TextWrappingValues.None,
                _ => W.TextWrappingValues.Around
            },
            VerticalPosition = verticalAnchor == "page" ? W.VerticalAnchorValues.Page : verticalAnchor == "margin" ? W.VerticalAnchorValues.Margin : W.VerticalAnchorValues.Text,
            HorizontalPosition = horizontalAnchor == "page" ? W.HorizontalAnchorValues.Page : W.HorizontalAnchorValues.Margin
        };
        if (xAlign == "right") frame.XAlign = W.HorizontalAlignmentValues.Right;
        if (x.HasValue) frame.X = x.Value.ToString(System.Globalization.CultureInfo.InvariantCulture);
        if (y.HasValue) frame.Y = y.Value.ToString(System.Globalization.CultureInfo.InvariantCulture);
        if (width.HasValue) frame.Width = width.Value.ToString(System.Globalization.CultureInfo.InvariantCulture);
        if (horizontalSpace.HasValue) frame.HorizontalSpace = horizontalSpace.Value.ToString(System.Globalization.CultureInfo.InvariantCulture);
        W.ParagraphProperties properties = paragraph._paragraph.ParagraphProperties ??= new W.ParagraphProperties();
        properties.AddChild(frame, true);
    }

    private static PdfPigDocument RenderFramePdf(WordDocument document) {
        using WordDocument loaded = WordDocument.Load(new MemoryStream(document.ToBytes(WordFileFormat.Docx)));
        return PdfPigDocument.Open(loaded.ToPdf(new WordPdfExportOptions { IncludePageNumbers = false })
            .ToBytes(new OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true }));
    }

    private static UglyToad.PdfPig.Content.Word FindFrameWord(PdfPigDocument pdf, string token) =>
        pdf.GetPages().SelectMany(page => page.GetWords()).First(word => word.Text == token);
}
