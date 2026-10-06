using OfficeIMO.Pdf;
using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using Xunit;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Tests;

public partial class Word {
    [Theory]
    [InlineData(false, false)]
    [InlineData(false, true)]
    [InlineData(true, false)]
    [InlineData(true, true)]
    public void SaveAsPdf_HiddenParagraphMarkJoinsTextWithoutMutatingSource(bool nativeDoc, bool columns) {
        using WordDocument source = CreateJoinedParagraphDocument();
        WordParagraph alpha = source.AddParagraph("ALPHA");
        HideJoinMark(alpha, true);
        WordParagraph hiddenBlank = source.AddParagraph("SECRET"); hiddenBlank.Hidden = true; HideJoinMark(hiddenBlank, true);
        source.AddParagraph("BETA");
        source.AddParagraph("SECOND");
        if (columns) source.Sections[0].ColumnCount = 2;
        using WordDocument document = WordDocument.Load(new MemoryStream(source.ToBytes(nativeDoc ? WordFileFormat.Doc : WordFileFormat.Docx)));
        string before = document._wordprocessingDocument.MainDocumentPart!.Document.OuterXml;
        using var pdf = OpenJoinedParagraphPdf(document);
        var letters = pdf.GetPage(1).Letters;
        Assert.Contains("ALPHABETA", pdf.GetPage(1).Text);
        Assert.DoesNotContain("SECRET", pdf.GetPage(1).Text);
        Assert.Equal(letters.First(l => l.Value == "A").StartBaseLine.Y,
            Assert.Single(letters, l => l.Value == "B").StartBaseLine.Y, 3);
        Assert.Equal(before, document._wordprocessingDocument.MainDocumentPart!.Document.OuterXml);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void SaveAsPdf_VisibleParagraphMarkOverridesInheritedHiddenMark(bool nativeDoc) {
        using WordDocument source = CreateJoinedParagraphDocument();
        W.Styles styles = source._wordprocessingDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!;
        styles.Append(new W.Style { StyleId = "HiddenJoin", Type = W.StyleValues.Paragraph,
            StyleName = new W.StyleName { Val = "Hidden Join" }, BasedOn = new W.BasedOn { Val = "Normal" },
            StyleRunProperties = new W.StyleRunProperties(new W.Vanish()) });
        WordParagraph alpha = source.AddParagraph("ALPHA"); alpha.SetStyleId("HiddenJoin"); alpha.Hidden = false;
        HideJoinMark(alpha, false);
        source.AddParagraph("BETA");
        using WordDocument document = WordDocument.Load(new MemoryStream(source.ToBytes(nativeDoc ? WordFileFormat.Doc : WordFileFormat.Docx)));
        using var pdf = OpenJoinedParagraphPdf(document);
        Assert.True(pdf.GetPage(1).Letters.First(l => l.Value == "A").StartBaseLine.Y >
            Assert.Single(pdf.GetPage(1).Letters, l => l.Value == "B").StartBaseLine.Y);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void SaveAsPdf_JoinedParagraphRetainsEachSourceStyleFont(bool nativeDoc) {
        using WordDocument source = CreateJoinedParagraphDocument();
        W.Styles styles = source._wordprocessingDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!;
        foreach (var item in new[] { ("LargeJoin", "32"), ("SmallJoin", "16") })
            styles.Append(new W.Style { StyleId = item.Item1, Type = W.StyleValues.Paragraph,
                StyleName = new W.StyleName { Val = item.Item1 }, BasedOn = new W.BasedOn { Val = "Normal" },
                StyleRunProperties = new W.StyleRunProperties(new W.FontSize { Val = item.Item2 }, new W.FontSizeComplexScript { Val = item.Item2 }) });
        WordParagraph alpha = source.AddParagraph("ALPHA"); alpha.SetStyleId("LargeJoin"); HideJoinMark(alpha, true);
        source.AddParagraph("BETA").SetStyleId("SmallJoin");
        using WordDocument document = WordDocument.Load(new MemoryStream(source.ToBytes(nativeDoc ? WordFileFormat.Doc : WordFileFormat.Docx)));
        using var pdf = OpenJoinedParagraphPdf(document);
        var a = pdf.GetPage(1).Letters.First(l => l.Value == "A");
        var b = Assert.Single(pdf.GetPage(1).Letters, l => l.Value == "B");
        Assert.Equal(a.StartBaseLine.Y, b.StartBaseLine.Y, 3);
        Assert.Equal(16D, a.FontSize, 3); Assert.Equal(8D, b.FontSize, 3);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void SaveAsPdf_HiddenMarkChainRetainsFirstAlignmentAndFinalSpacing(bool nativeDoc) {
        using WordDocument source = CreateJoinedParagraphDocument();
        WordParagraph alpha = source.AddParagraph("ALPHA"); alpha.ParagraphAlignment = WordParagraphAlignment.Right;
        alpha.LineSpacingBeforePoints = 20; alpha.LineSpacingAfterPoints = 80; HideJoinMark(alpha, true);
        WordParagraph beta = source.AddParagraph("BETA"); beta.LineSpacingBeforePoints = 80; beta.LineSpacingAfterPoints = 20;
        source.AddParagraph("SECOND");
        using WordDocument document = WordDocument.Load(new MemoryStream(source.ToBytes(nativeDoc ? WordFileFormat.Doc : WordFileFormat.Docx)));
        using var pdf = OpenJoinedParagraphPdf(document);
        var letters = pdf.GetPage(1).Letters;
        var a = letters.First(l => l.Value == "A"); var b = Assert.Single(letters, l => l.Value == "B");
        Assert.Equal(a.StartBaseLine.Y, b.StartBaseLine.Y, 3);
        Assert.True(a.StartBaseLine.X > 400);
        Assert.InRange(a.StartBaseLine.Y - letters.Last(l => l.Value == "S").StartBaseLine.Y, 33D, 35D);
    }

    private static WordDocument CreateJoinedParagraphDocument() {
        WordDocument document = WordDocument.Create();
        W.Style normal = document._wordprocessingDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!
            .Elements<W.Style>().Single(style => style.StyleId?.Value == "Normal");
        normal.StyleRunProperties = new W.StyleRunProperties(new W.RunFonts { Ascii = "Arial", HighAnsi = "Arial", EastAsia = "Arial", ComplexScript = "Arial" },
            new W.FontSize { Val = "24" }, new W.FontSizeComplexScript { Val = "24" });
        normal.StyleParagraphProperties = new W.StyleParagraphProperties(new W.SpacingBetweenLines {
            Before = "0", After = "0", Line = "240", LineRule = W.LineSpacingRuleValues.Auto });
        return document;
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void SaveAsPdf_HiddenParagraphJoinUsesInheritedMarkVisibility(bool nativeDoc) {
        using WordDocument source = CreateJoinedParagraphDocument();
        W.Styles styles = source._wordprocessingDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!;
        styles.Append(new W.Style { StyleId = "InheritedHiddenJoin", Type = W.StyleValues.Paragraph,
            StyleName = new W.StyleName { Val = "Inherited Hidden Join" }, BasedOn = new W.BasedOn { Val = "Normal" },
            StyleRunProperties = new W.StyleRunProperties(new W.Vanish()) });
        WordParagraph alpha = source.AddParagraph("ALPHA"); alpha.SetStyleId("InheritedHiddenJoin"); alpha.Hidden = false;
        source.AddParagraph("BETA");
        using WordDocument document = WordDocument.Load(new MemoryStream(source.ToBytes(nativeDoc ? WordFileFormat.Doc : WordFileFormat.Docx)));
        using var pdf = OpenJoinedParagraphPdf(document);
        Assert.Contains("ALPHABETA", pdf.GetPage(1).Text);
        Assert.Equal(pdf.GetPage(1).Letters.First(l => l.Value == "A").StartBaseLine.Y,
            Assert.Single(pdf.GetPage(1).Letters, l => l.Value == "B").StartBaseLine.Y, 3);
    }

    [Fact]
    public void SaveAsPdf_HiddenParagraphJoinDoesNotRemoveFollowingPageBreakBefore() {
        using WordDocument document = CreateJoinedParagraphDocument();
        WordParagraph alpha = document.AddParagraph("ALPHA"); HideJoinMark(alpha, true);
        document.AddParagraph("BETA").PageBreakBeforeOverride = true;
        var options = new WordToPdfOptions { IncludePageNumbers = false,
            ResourcePolicy = PdfResourcePolicy.CreatePortableDeterministic() };
        using var pdf = PdfPigDocument.Open(document.ToPdfBytes(options));
        Assert.Equal(2, pdf.NumberOfPages);
        Assert.Contains("ALPHA", pdf.GetPage(1).Text); Assert.Contains("BETA", pdf.GetPage(2).Text);
    }

    [Fact]
    public void SaveAsPdf_HiddenParagraphJoinReportsUnsupportedDecorationWithoutLosingText() {
        using WordDocument document = CreateJoinedParagraphDocument();
        WordParagraph alpha = document.AddParagraph("ALPHA"); HideJoinMark(alpha, true);
        alpha._paragraph.ParagraphProperties!.AddChild(new W.Shading { Fill = "FFCC00", Val = W.ShadingPatternValues.Clear }, true);
        document.AddParagraph("BETA");
        var options = new WordToPdfOptions { IncludePageNumbers = false,
            ResourcePolicy = PdfResourcePolicy.CreatePortableDeterministic() };
        var result = document.ToPdfDocumentResult(options);
        using var pdf = PdfPigDocument.Open(result.Value.ToBytes());
        Assert.Contains("ALPHA", pdf.GetPage(1).Text); Assert.Contains("BETA", pdf.GetPage(1).Text);
        Assert.True(pdf.GetPage(1).Letters.First(l => l.Value == "A").StartBaseLine.Y >
            Assert.Single(pdf.GetPage(1).Letters, l => l.Value == "B").StartBaseLine.Y);
        Assert.Contains(result.Report.Warnings, warning => warning.Code == "NativeHiddenParagraphJoinUnsupported");
    }

    private static void HideJoinMark(WordParagraph paragraph, bool hidden) {
        paragraph._paragraph.ParagraphProperties ??= new W.ParagraphProperties();
        paragraph._paragraph.ParagraphProperties.ParagraphMarkRunProperties = new W.ParagraphMarkRunProperties(new W.Vanish { Val = hidden });
    }

    private static PdfPigDocument OpenJoinedParagraphPdf(WordDocument document) => PdfPigDocument.Open(document.ToPdfBytes(
        new WordToPdfOptions { IncludePageNumbers = false, ResourcePolicy = PdfResourcePolicy.CreatePortableDeterministic() }));
}
