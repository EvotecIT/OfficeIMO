using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using Xunit;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Tests;

public partial class Word {
    // These three controls were independently exported by native Word: a
    // continuous header change keeps the shared page's original header and
    // takes effect only when later body content starts another physical page.
    [Theory]
    [InlineData("inherited", 1)]
    [InlineData("changed-shared-page", 1)]
    [InlineData("changed-later-body-break", 2)]
    public void SaveAsPdf_ContinuousDefaultHeaderFollowsPhysicalPageOwnership(string scenario, int expectedPages) {
        bool changedHeader = scenario != "inherited";
        bool bodyBreak = scenario == "changed-later-body-break";
        using WordDocument source = WordDocument.Load(new MemoryStream(
            CreateContinuousHeaderSource(changedHeader, bodyBreak)));
        string before = source.MainDocumentPartRoot.Document.OuterXml;

        var result = source.ToPdfDocumentResult(EmptyRunningOptions());
        Assert.DoesNotContain(result.Report.Warnings, warning => warning.Code == "NativeContinuousSectionPaginationApproximated");
        using var pdf = PdfPigDocument.Open(result.Value.ToBytes());

        Assert.Equal(expectedPages, pdf.NumberOfPages);
        AssertContinuousHeaderPage(pdf, 1, "HEADER A");
        if (bodyBreak) {
            AssertContinuousHeaderPage(pdf, 2, "HEADER B");
            Assert.Contains("S2 BEFORE BODY BREAK", ContinuousHeaderPageText(pdf, 1));
            Assert.DoesNotContain("S2 FOLLOWING BODY MARKER", ContinuousHeaderPageText(pdf, 1));
            Assert.Contains("S2 FOLLOWING BODY MARKER", ContinuousHeaderPageText(pdf, 2));
            Assert.Contains("S2 END", ContinuousHeaderPageText(pdf, 2));
        }
        AssertContinuousHeaderBody(pdf, ContinuousHeaderBodyMarkers);
        Assert.Equal(before, source.MainDocumentPartRoot.Document.OuterXml);
    }

    [Fact]
    public void SaveAsPdf_ContinuousHeaderFallbackReportsPaginationAndRetainsSupportedTableContent() {
        using WordDocument source = WordDocument.Load(new MemoryStream(
            CreateContinuousHeaderSource(changedHeader: true, bodyBreak: false)));
        WordTable table = source.Sections[1].AddTable(1, 1);
        table.Rows[0].Cells[0].Paragraphs[0].Text = "TABLE CELL";

        var result = source.ToPdfDocumentResult(EmptyRunningOptions());
        var warning = Assert.Single(result.Report.Warnings, item => item.Code == "NativeContinuousSectionPaginationApproximated");
        Assert.Equal(OfficeIMO.OfficeConversionLossKind.Approximation, warning.LossKind);
        using var pdf = PdfPigDocument.Open(result.Value.ToBytes());

        Assert.Equal(2, pdf.NumberOfPages);
        AssertContinuousHeaderPage(pdf, 1, "HEADER A");
        AssertContinuousHeaderPage(pdf, 2, "HEADER B");
        Assert.Contains("TABLE CELL", ContinuousHeaderPageText(pdf, 2));
        AssertContinuousHeaderBody(pdf, ContinuousHeaderBodyMarkers.Concat(new[] { "TABLE CELL" }).ToArray());
    }

    [Fact]
    public void SaveAsPdf_ContinuousInheritedSuccessorRetainsLatestHeaderForNextPhysicalPage() {
        using WordDocument source = WordDocument.Load(new MemoryStream(
            CreateContinuousHeaderSource(changedHeader: true, bodyBreak: false, successor: "inherited")));

        using var pdf = PdfPigDocument.Open(source.ToPdfBytes(EmptyRunningOptions()));

        Assert.Equal(2, pdf.NumberOfPages);
        AssertContinuousHeaderPage(pdf, 1, "HEADER A");
        AssertContinuousHeaderPage(pdf, 2, "HEADER B");
        Assert.Contains("S3 BEFORE BODY BREAK", ContinuousHeaderPageText(pdf, 1));
        Assert.Contains("S3 FOLLOWING BODY MARKER", ContinuousHeaderPageText(pdf, 2));
        AssertContinuousHeaderBody(pdf, ContinuousHeaderBodyMarkers.Concat(new[] {
            "S3 BEFORE BODY BREAK", "S3 FOLLOWING BODY MARKER"
        }).ToArray());
    }

    [Fact]
    public void SaveAsPdf_ExplicitSectionOwnsPageAfterUnconsumedContinuousHeader() {
        using WordDocument source = WordDocument.Load(new MemoryStream(
            CreateContinuousHeaderSource(changedHeader: true, bodyBreak: false, successor: "explicit")));

        using var pdf = PdfPigDocument.Open(source.ToPdfBytes(EmptyRunningOptions()));

        Assert.Equal(2, pdf.NumberOfPages);
        AssertContinuousHeaderPage(pdf, 1, "HEADER A");
        AssertContinuousHeaderPage(pdf, 2, "HEADER C");
        Assert.DoesNotContain("HEADER B", string.Join(" ", pdf.GetPages().Select(page => page.Text)));
        Assert.Contains("S3 AFTER EXPLICIT SECTION", ContinuousHeaderPageText(pdf, 2));
        AssertContinuousHeaderBody(pdf, ContinuousHeaderBodyMarkers.Concat(new[] { "S3 AFTER EXPLICIT SECTION" }).ToArray());
    }

    [Fact]
    public void SaveAsPdf_ContinuousDefaultHeaderPersistsAcrossNaturalOverflow() {
        const int flowParagraphs = 40;
        using WordDocument source = WordDocument.Load(new MemoryStream(
            CreateContinuousHeaderSource(changedHeader: true, bodyBreak: false, overflowParagraphs: flowParagraphs)));

        using var pdf = PdfPigDocument.Open(source.ToPdfBytes(EmptyRunningOptions()));

        Assert.True(pdf.NumberOfPages >= 3);
        AssertContinuousHeaderPage(pdf, 1, "HEADER A");
        for (int page = 2; page <= pdf.NumberOfPages; page++) AssertContinuousHeaderPage(pdf, page, "HEADER B");
        AssertContinuousHeaderBody(pdf, ContinuousHeaderBodyMarkers.Concat(
            Enumerable.Range(0, flowParagraphs).Select(index => "FLOW" + index.ToString("D2"))).ToArray());
    }

    [Fact]
    public void SaveAsPdf_ContinuousTallerHeaderMeasuresTheNextPhysicalBodyFrame() {
        using WordDocument source = WordDocument.Load(new MemoryStream(
            CreateContinuousHeaderSource(changedHeader: true, bodyBreak: true, discriminator: "tall")));
        using WordDocument control = WordDocument.Load(new MemoryStream(
            CreateContinuousHeaderSource(changedHeader: true, bodyBreak: true)));

        using var pdf = PdfPigDocument.Open(source.ToPdfBytes(EmptyRunningOptions()));
        using var ordinary = PdfPigDocument.Open(control.ToPdfBytes(EmptyRunningOptions()));

        Assert.Equal(2, pdf.NumberOfPages);
        AssertContinuousHeaderPage(pdf, 1, "HEADER A");
        AssertContinuousHeaderPage(pdf, 2, ContinuousLongHeaderB);
        AssertContinuousHeaderBody(pdf, ContinuousHeaderBodyMarkers);
        var headerTokens = new HashSet<string>(ContinuousLongHeaderB.Split(' '));
        var headerLetters = pdf.GetPage(2).GetWords().Where(word => headerTokens.Contains(word.Text))
            .SelectMany(word => word.Letters).ToArray();
        Assert.True(headerLetters.Select(letter => letter.StartBaseLine.Y).Distinct().Count() > 1);
        var following = Assert.Single(pdf.GetPage(2).GetWords(), word => word.Text == "FOLLOWING");
        double bodyTop = following.Letters.Max(letter => letter.BoundingBox.Top);
        double headerBottom = headerLetters.Min(letter => letter.BoundingBox.Bottom);
        Assert.True(bodyTop < headerBottom, "The first body line overlaps the taller running header.");
        var ordinaryFollowing = Assert.Single(ordinary.GetPage(2).GetWords(), word => word.Text == "FOLLOWING");
        Assert.True(bodyTop < ordinaryFollowing.Letters.Max(letter => letter.BoundingBox.Top));
    }

    [Fact]
    public void SaveAsPdf_ContinuousEmptyFirstSectionRetainsPageHeaderAndBlankParagraphSpacing() {
        using WordDocument source = WordDocument.Load(new MemoryStream(
            CreateContinuousHeaderSource(changedHeader: true, bodyBreak: false, discriminator: "empty-first")));
        using WordDocument control = WordDocument.Load(new MemoryStream(
            CreateContinuousHeaderSource(changedHeader: true, bodyBreak: false)));

        using var pdf = PdfPigDocument.Open(source.ToPdfBytes(EmptyRunningOptions()));
        using var ordinary = PdfPigDocument.Open(control.ToPdfBytes(EmptyRunningOptions()));

        Assert.Equal(1, pdf.NumberOfPages);
        AssertContinuousHeaderPage(pdf, 1, "HEADER A");
        AssertContinuousHeaderBody(pdf, new[] { "S2 AFTER EMPTY FIRST SECTION", "S2 END" });
        double bodyBaseline = pdf.GetPage(1).GetWords().First(word => word.Text == "S2").Letters[0].StartBaseLine.Y;
        double ordinaryBaseline = ordinary.GetPage(1).GetWords().First(word => word.Text == "S1").Letters[0].StartBaseLine.Y;
        // The source authors a 12-point exact paragraph, independently of the
        // native application's font-dependent glyph top or baseline offset.
        Assert.Equal(12D, ordinaryBaseline - bodyBaseline, 3);
    }

    [Fact]
    public void SaveAsPdf_SuccessiveContinuousHeadersSelectTheLastStoryAtTheNextPhysicalPage() {
        using WordDocument source = WordDocument.Load(new MemoryStream(
            CreateContinuousHeaderSource(changedHeader: true, bodyBreak: true, discriminator: "successive")));

        using var pdf = PdfPigDocument.Open(source.ToPdfBytes(EmptyRunningOptions()));

        Assert.Equal(2, pdf.NumberOfPages);
        AssertContinuousHeaderPage(pdf, 1, "HEADER A");
        AssertContinuousHeaderPage(pdf, 2, "HEADER C");
        Assert.Contains("S2 SECTION END", ContinuousHeaderPageText(pdf, 1));
        Assert.Contains("S3 BEFORE BODY BREAK", ContinuousHeaderPageText(pdf, 1));
        Assert.DoesNotContain("S3 FOLLOWING BODY MARKER", ContinuousHeaderPageText(pdf, 1));
        Assert.Contains("S3 FOLLOWING BODY MARKER", ContinuousHeaderPageText(pdf, 2));
        AssertContinuousHeaderBody(pdf, new[] {
            "S1 BEFORE CONTINUOUS", "S1 SECTION END", "S2 AFTER CONTINUOUS", "S2 SECTION END",
            "S3 AFTER CONTINUOUS", "S3 BEFORE BODY BREAK", "S3 FOLLOWING BODY MARKER", "S3 END"
        });
    }

    private const string ContinuousLongHeaderB =
        "HEADER B LONG ALPHA BRAVO CHARLIE DELTA ECHO FOXTROT GOLF HOTEL INDIA " +
        "JULIET KILO LIMA MIKE NOVEMBER OSCAR PAPA QUEBEC ROMEO SIERRA TANGO " +
        "UNIFORM VICTOR WHISKEY XRAY YANKEE ZULU";

    private static readonly string[] ContinuousHeaderBodyMarkers = {
        "S1 BEFORE CONTINUOUS", "S1 SECTION END", "S2 AFTER CONTINUOUS",
        "S2 BEFORE BODY BREAK", "S2 FOLLOWING BODY MARKER", "S2 END"
    };

    private static string ContinuousHeaderPageText(PdfPigDocument pdf, int page) =>
        string.Join(" ", pdf.GetPage(page).GetWords().Select(word => word.Text));

    private static void AssertContinuousHeaderPage(PdfPigDocument pdf, int number, string header) {
        string text = ContinuousHeaderPageText(pdf, number);
        Assert.Contains(header, text);
        string[] identities = { "HEADER A", "HEADER B", "HEADER C" };
        string selected = Assert.Single(identities, identity => header == identity || header.StartsWith(identity + " ", StringComparison.Ordinal));
        foreach (string other in identities.Where(other => other != selected))
            Assert.DoesNotContain(other, text);
        Assert.Contains("FOOTER FIXED", text);
        Assert.True(text.Contains("S1 ") || text.Contains("S2 ") || text.Contains("S3 ") || text.Contains("FLOW"));
        var page = pdf.GetPage(number);
        Assert.Single(page.GetWords(), word => word.Text == "HEADER");
        Assert.Single(page.GetWords(), word => word.Text == "FOOTER");
        Assert.All(page.Letters, letter => {
            Assert.InRange(letter.BoundingBox.Left, 0D, page.Width);
            Assert.InRange(letter.BoundingBox.Right, 0D, page.Width);
            Assert.InRange(letter.BoundingBox.Bottom, 0D, page.Height);
            Assert.InRange(letter.BoundingBox.Top, 0D, page.Height);
        });
    }

    private static void AssertContinuousHeaderBody(PdfPigDocument pdf, IReadOnlyList<string> markers) {
        string allText = string.Join(" ", Enumerable.Range(1, pdf.NumberOfPages).Select(page => ContinuousHeaderPageText(pdf, page)));
        int previous = -1;
        foreach (string marker in markers) {
            int position = allText.IndexOf(marker, StringComparison.Ordinal);
            Assert.True(position > previous, "Body marker was missing or reordered: " + marker);
            Assert.Equal(-1, allText.IndexOf(marker, position + marker.Length, StringComparison.Ordinal));
            previous = position;
        }
    }

    private static byte[] CreateContinuousHeaderSource(bool changedHeader, bool bodyBreak,
        string? successor = null, int overflowParagraphs = 0, string? discriminator = null) {
        using var stream = new MemoryStream();
        using (WordprocessingDocument package = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document)) {
            MainDocumentPart main = package.AddMainDocumentPart();
            HeaderPart first = main.AddNewPart<HeaderPart>("rIdHeaderA");
            first.Header = new W.Header(ContinuousHeaderParagraph("HEADER A"));
            HeaderPart second = main.AddNewPart<HeaderPart>("rIdHeaderB");
            second.Header = new W.Header(ContinuousHeaderParagraph(discriminator == "tall" ? ContinuousLongHeaderB : "HEADER B"));
            FooterPart footer = main.AddNewPart<FooterPart>("rIdFooterFixed");
            footer.Footer = new W.Footer(ContinuousHeaderParagraph("FOOTER FIXED"));
            StyleDefinitionsPart styles = main.AddNewPart<StyleDefinitionsPart>();
            styles.Styles = new W.Styles(new W.DocDefaults(new W.RunPropertiesDefault(
                new W.RunPropertiesBaseStyle(new W.RunFonts { Ascii = "Arial", HighAnsi = "Arial", EastAsia = "Arial", ComplexScript = "Arial" },
                    new W.FontSize { Val = "20" }, new W.FontSizeComplexScript { Val = "20" }))));
            DocumentSettingsPart settings = main.AddNewPart<DocumentSettingsPart>();
            settings.Settings = new W.Settings(new W.Compatibility(new W.CompatibilitySetting {
                Name = W.CompatSettingNameValues.CompatibilityMode,
                Uri = "http://schemas.microsoft.com/office/word", Val = "15"
            }));

            bool emptyFirst = discriminator == "empty-first";
            W.Paragraph firstEnd = ContinuousHeaderParagraph(emptyFirst ? null : ContinuousHeaderBodyMarkers[1]);
            firstEnd.ParagraphProperties!.Append(ContinuousHeaderSectionProperties("rIdHeaderA", "rIdFooterFixed", continuous: false));
            var body = new W.Body();
            if (!emptyFirst) body.Append(ContinuousHeaderParagraph(ContinuousHeaderBodyMarkers[0]));
            body.Append(firstEnd);
            if (discriminator == "successive") {
                body.Append(ContinuousHeaderParagraph("S2 AFTER CONTINUOUS"));
                W.Paragraph secondEnd = ContinuousHeaderParagraph("S2 SECTION END");
                secondEnd.ParagraphProperties!.Append(ContinuousHeaderSectionProperties("rIdHeaderB", null, continuous: true));
                body.Append(secondEnd, ContinuousHeaderParagraph("S3 AFTER CONTINUOUS"),
                    ContinuousHeaderParagraph("S3 BEFORE BODY BREAK"),
                    ContinuousHeaderParagraph("S3 FOLLOWING BODY MARKER", pageBreakBefore: true), ContinuousHeaderParagraph("S3 END"));
                HeaderPart third = main.AddNewPart<HeaderPart>("rIdHeaderC");
                third.Header = new W.Header(ContinuousHeaderParagraph("HEADER C"));
                body.Append(ContinuousHeaderSectionProperties("rIdHeaderC", null, continuous: true));
            } else if (emptyFirst) {
                body.Append(ContinuousHeaderParagraph("S2 AFTER EMPTY FIRST SECTION"), ContinuousHeaderParagraph("S2 END"),
                    ContinuousHeaderSectionProperties("rIdHeaderB", null, continuous: true));
            } else {
                for (int index = 2; index < ContinuousHeaderBodyMarkers.Length; index++)
                    body.Append(ContinuousHeaderParagraph(ContinuousHeaderBodyMarkers[index], pageBreakBefore: bodyBreak && index == 4));
                for (int index = 0; index < overflowParagraphs; index++)
                    body.Append(ContinuousHeaderParagraph("FLOW" + index.ToString("D2")));

                string? secondHeader = changedHeader ? "rIdHeaderB" : null;
                if (successor != null) {
                    body.Elements<W.Paragraph>().Last().ParagraphProperties!.Append(
                        ContinuousHeaderSectionProperties(secondHeader, null, continuous: true));
                    if (successor == "inherited") {
                        body.Append(ContinuousHeaderParagraph("S3 BEFORE BODY BREAK"));
                        body.Append(ContinuousHeaderParagraph("S3 FOLLOWING BODY MARKER", pageBreakBefore: true));
                        body.Append(ContinuousHeaderSectionProperties(null, null, continuous: true));
                    } else {
                        HeaderPart third = main.AddNewPart<HeaderPart>("rIdHeaderC");
                        third.Header = new W.Header(ContinuousHeaderParagraph("HEADER C"));
                        body.Append(ContinuousHeaderParagraph("S3 AFTER EXPLICIT SECTION"));
                        body.Append(ContinuousHeaderSectionProperties("rIdHeaderC", null, continuous: false));
                    }
                } else body.Append(ContinuousHeaderSectionProperties(secondHeader, null, continuous: true));
            }
            main.Document = new W.Document(body);
        }
        return stream.ToArray();
    }

    private static W.Paragraph ContinuousHeaderParagraph(string? text, bool pageBreakBefore = false) {
        var properties = new W.ParagraphProperties(new W.KeepNext { Val = false }, new W.KeepLines { Val = false });
        if (pageBreakBefore) properties.Append(new W.PageBreakBefore());
        properties.Append(new W.WidowControl { Val = false },
            new W.SpacingBetweenLines { Before = "0", After = "0", Line = "240", LineRule = W.LineSpacingRuleValues.Exact },
            new W.Justification { Val = W.JustificationValues.Left });
        var paragraph = new W.Paragraph(properties);
        if (text != null) paragraph.Append(new W.Run(new W.Text(text) { Space = SpaceProcessingModeValues.Preserve }));
        return paragraph;
    }

    private static W.SectionProperties ContinuousHeaderSectionProperties(string? header, string? footer, bool continuous) {
        var properties = new W.SectionProperties();
        if (header != null) properties.Append(new W.HeaderReference { Id = header, Type = W.HeaderFooterValues.Default });
        if (footer != null) properties.Append(new W.FooterReference { Id = footer, Type = W.HeaderFooterValues.Default });
        if (continuous) properties.Append(new W.SectionType { Val = W.SectionMarkValues.Continuous });
        properties.Append(new W.PageSize { Width = 4800U, Height = 6400U }, new W.PageMargin {
            Top = 640, Bottom = 640, Left = 480U, Right = 480U, Header = 360U, Footer = 360U, Gutter = 0U
        });
        return properties;
    }
}
