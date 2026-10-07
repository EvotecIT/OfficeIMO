using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using OfficeIMO.Pdf;
using W = DocumentFormat.OpenXml.Wordprocessing;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using Xunit;

namespace OfficeIMO.Tests;

public partial class Word {
    [Theory]
    [InlineData(false, false, false)]
    [InlineData(false, false, true)]
    [InlineData(false, true, false)]
    [InlineData(false, true, true)]
    [InlineData(true, false, false)]
    [InlineData(true, false, true)]
    [InlineData(true, true, false)]
    [InlineData(true, true, true)]
    public void SaveAsPdf_HiddenHeaderFooterMarkJoinsTextAndPreservesWhitespace(bool footer, bool nativeDoc, bool whitespace) {
        using WordDocument source = CreateJoinedParagraphDocument();
        source.AddParagraph("BODY"); source.AddHeadersAndFooters();
        WordHeaderFooter story = footer ? source.Footer.Default : source.Header.Default;
        WordParagraph alpha = story.AddParagraph(whitespace ? "ALPHA " : "ALPHA");
        story.AddParagraph("BETA"); HideJoinMark(alpha, true);
        using WordDocument document = WordDocument.Load(new MemoryStream(source.ToBytes(nativeDoc ? WordFileFormat.Doc : WordFileFormat.Docx)));
        string before = GetHeaderJoinStoryXml(document, footer);
        using var pdf = OpenJoinedParagraphPdf(document);
        var page = pdf.GetPage(1);
        if (whitespace) {
            var words = page.GetWords().ToArray();
            var first = Assert.Single(words, word => word.Text == "ALPHA");
            var last = Assert.Single(words, word => word.Text == "BETA");
            Assert.Equal(first.BoundingBox.Bottom, last.BoundingBox.Bottom, 3);
            Assert.True(last.BoundingBox.Left > first.BoundingBox.Right);
        } else Assert.Single(page.GetWords(), word => word.Text == "ALPHABETA");
        Assert.Equal(before, GetHeaderJoinStoryXml(document, footer));
    }

    [Theory]
    [InlineData(false, false)]
    [InlineData(false, true)]
    [InlineData(true, false)]
    [InlineData(true, true)]
    public void SaveAsPdf_VisibleHeaderFooterMarkRetainsSeparateLines(bool footer, bool nativeDoc) {
        using WordDocument source = CreateJoinedParagraphDocument();
        source.AddParagraph("BODY"); source.AddHeadersAndFooters();
        WordHeaderFooter story = footer ? source.Footer.Default : source.Header.Default;
        WordParagraph alpha = story.AddParagraph("ALPHA"); HideJoinMark(alpha, false);
        story.AddParagraph("BETA");
        using WordDocument document = WordDocument.Load(new MemoryStream(source.ToBytes(nativeDoc ? WordFileFormat.Doc : WordFileFormat.Docx)));
        using var pdf = OpenJoinedParagraphPdf(document);
        var words = pdf.GetPage(1).GetWords().ToArray();
        var first = Assert.Single(words, word => word.Text == "ALPHA");
        var last = Assert.Single(words, word => word.Text == "BETA");
        Assert.True(first.BoundingBox.Bottom > last.BoundingBox.Bottom);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void SaveAsPdf_HiddenHeaderFooterChainSkipsHiddenTextAndUsesFirstAlignment(bool footer) {
        using WordDocument document = CreateJoinedParagraphDocument();
        document.AddParagraph("BODY"); document.AddHeadersAndFooters();
        WordHeaderFooter story = footer ? document.Footer.Default : document.Header.Default;
        WordParagraph alpha = story.AddParagraph("ALPHA "); alpha.ParagraphAlignment = WordParagraphAlignment.Center; HideJoinMark(alpha, true);
        WordParagraph hidden = story.AddParagraph("SECRET"); hidden.Hidden = true; HideJoinMark(hidden, true);
        WordParagraph beta = story.AddParagraph("BETA"); beta.ParagraphAlignment = WordParagraphAlignment.Right;
        story.AddParagraph("SECOND");
        using var pdf = OpenJoinedParagraphPdf(document);
        var words = pdf.GetPage(1).GetWords().ToArray();
        var first = Assert.Single(words, word => word.Text == "ALPHA");
        var last = Assert.Single(words, word => word.Text == "BETA");
        Assert.Equal(first.BoundingBox.Bottom, last.BoundingBox.Bottom, 3);
        Assert.True(last.BoundingBox.Right < 450);
        Assert.DoesNotContain("SECRET", pdf.GetPage(1).Text);
        Assert.Contains("SECOND", pdf.GetPage(1).Text);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void SaveAsPdf_MixedHeaderFooterJoinTypographyRetainsContentAndReportsLimitation(bool footer) {
        using WordDocument document = CreateJoinedParagraphDocument();
        document.AddParagraph("BODY"); document.AddHeadersAndFooters();
        WordHeaderFooter story = footer ? document.Footer.Default : document.Header.Default;
        WordParagraph alpha = story.AddParagraph("ALPHA"); HideJoinMark(alpha, true);
        story.AddParagraph("BETA").Bold = true;
        var result = document.ToPdfDocumentResult(new WordToPdfOptions { IncludePageNumbers = false,
            ResourcePolicy = PdfResourcePolicy.CreatePortableDeterministic() });
        using var pdf = PdfPigDocument.Open(result.Value.ToBytes());
        Assert.Contains("ALPHA", pdf.GetPage(1).Text); Assert.Contains("BETA", pdf.GetPage(1).Text);
        Assert.Contains(result.Report.Warnings, warning => warning.Code == "NativeHiddenParagraphJoinUnsupported" && warning.Source == "header/footer");
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void SaveAsPdf_HiddenHeaderFooterTableCellMarkJoinsWithinCell(bool footer) {
        using WordDocument source = CreateJoinedParagraphDocument();
        source.AddParagraph("BODY"); source.AddHeadersAndFooters();
        WordHeaderFooter story = footer ? source.Footer.Default : source.Header.Default;
        WordTable table = story.AddTable(1, 1);
        WordParagraph alpha = table.Rows[0].Cells[0].AddParagraph("ALPHA ", removeExistingParagraphs: true); HideJoinMark(alpha, true);
        table.Rows[0].Cells[0].AddParagraph("BETA");
        using WordDocument document = WordDocument.Load(new MemoryStream(source.ToBytes()));
        using var pdf = OpenJoinedParagraphPdf(document);
        var words = pdf.GetPage(1).GetWords().ToArray();
        Assert.Equal(Assert.Single(words, word => word.Text == "ALPHA").BoundingBox.Bottom,
            Assert.Single(words, word => word.Text == "BETA").BoundingBox.Bottom, 3);
    }

    private static string GetHeaderJoinStoryXml(WordDocument document, bool footer) => footer
        ? document._wordprocessingDocument.MainDocumentPart!.FooterParts.First().Footer.OuterXml
        : document._wordprocessingDocument.MainDocumentPart!.HeaderParts.First().Header.OuterXml;

    [Theory]
    [InlineData(false, true)]
    [InlineData(true, true)]
    [InlineData(false, false)]
    [InlineData(true, false)]
    public void SaveAsPdf_HeaderFooterJoinComparesEffectiveDocumentDefaultSize(bool footer, bool sameSize) {
        using WordDocument document = CreateJoinedParagraphDocument();
        W.Styles styles = document._wordprocessingDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!;
        styles.DocDefaults = new W.DocDefaults(new W.RunPropertiesDefault(new W.RunPropertiesBaseStyle(
            new W.RunFonts { Ascii = "Arial", HighAnsi = "Arial", EastAsia = "Arial", ComplexScript = "Arial" },
            new W.FontSize { Val = "24" }, new W.FontSizeComplexScript { Val = "24" })));
        W.Style normal = styles.Elements<W.Style>().Single(style => style.StyleId?.Value == "Normal");
        normal.StyleRunProperties!.RemoveAllChildren<W.FontSize>();
        normal.StyleRunProperties.RemoveAllChildren<W.FontSizeComplexScript>();
        document.AddParagraph("BODY"); document.AddHeadersAndFooters();
        WordHeaderFooter story = footer ? document.Footer.Default : document.Header.Default;
        WordParagraph alpha = story.AddParagraph("ALPHA"); HideJoinMark(alpha, true);
        story.AddParagraph("BETA").FontSize = sameSize ? 12 : 14;
        var result = document.ToPdfDocumentResult(new WordToPdfOptions { IncludePageNumbers = false,
            ResourcePolicy = PdfResourcePolicy.CreatePortableDeterministic() });
        using var pdf = PdfPigDocument.Open(result.Value.ToBytes());
        if (sameSize) {
            Assert.Single(pdf.GetPage(1).GetWords(), word => word.Text == "ALPHABETA");
            Assert.DoesNotContain(result.Report.Warnings, warning => warning.Code == "NativeHiddenParagraphJoinUnsupported");
        } else {
            var words = pdf.GetPage(1).GetWords().ToArray();
            Assert.True(Assert.Single(words, word => word.Text == "ALPHA").BoundingBox.Bottom >
                Assert.Single(words, word => word.Text == "BETA").BoundingBox.Bottom);
            Assert.Contains(result.Report.Warnings, warning => warning.Code == "NativeHiddenParagraphJoinUnsupported");
        }
    }
}
