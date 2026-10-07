using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Pdf;
using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using UglyToad.PdfPig.Content;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class WordPdfRunSpacingReviewTests {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void LiteralTokenTextKeepsItsIdentityBesideAGenuinePageField(bool spaced) {
        using WordDocument document = WordDocument.Create();
        document.AddParagraph("Main");
        foreach (string literal in new[] { "{page}", "{pages}", "{documentpages}" }) {
            WordParagraph paragraph = document.HeaderDefaultOrCreate.AddParagraph(literal);
            if (spaced) paragraph.Spacing = 20;
        }
        WordParagraph mixed = document.HeaderDefaultOrCreate.AddParagraph("{page}");
        mixed.Spacing = 20;
        mixed._paragraph.Append(new SimpleField(new Run(new Text("888"))) { Instruction = " PAGE " });
        using var pdf = PdfPigDocument.Open(document.ToPdfBytes(Options()));
        string text = pdf.GetPage(1).Text;
        Assert.Contains("{page}", text, StringComparison.Ordinal);
        Assert.Contains("{pages}", text, StringComparison.Ordinal);
        Assert.Contains("{documentpages}", text, StringComparison.Ordinal);
        Assert.Equal(1, pdf.GetPage(1).Letters.Count(letter => letter.Value == "1"));
        Assert.Empty(document.ValidateDocument());
    }

    [Fact]
    public void NaturalDocumentPageFieldKeepsItsValueInAStyledHeaderZone() {
        using WordDocument document = WordDocument.Create();
        document.AddParagraph("Main");
        document.HeaderDefaultOrCreate.AddParagraph("Styled label").Spacing = 20;
        WordParagraph count = document.HeaderDefaultOrCreate.AddParagraph();
        count._paragraph.Append(new SimpleField(new Run(new Text("888"))) { Instruction = " NUMPAGES " });
        using var pdf = PdfPigDocument.Open(document.ToPdfBytes(Options()));
        Assert.DoesNotContain("{documentpages}", pdf.GetPage(1).Text, StringComparison.Ordinal);
        Assert.Equal(1, pdf.GetPage(1).Letters.Count(letter => letter.Value == "1"));
        Assert.Empty(document.ValidateDocument());
    }

    [Theory]
    [InlineData("body", 50)]
    [InlineData("body", 200)]
    [InlineData("inline", 50)]
    [InlineData("inline", 200)]
    [InlineData("cell", 50)]
    [InlineData("cell", 200)]
    [InlineData("header", 50)]
    [InlineData("header", 200)]
    [InlineData("footer", 50)]
    [InlineData("footer", 200)]
    public void InheritedListMarkerSpacingReachesRenderingAndColumnReservation(string route, int scale) {
        Letter[] natural = RenderList(route, 100, 0);
        Letter[] styled = RenderList(route, scale, 20);
        double naturalAdvance = natural.Single(letter => letter.Value == "2").StartBaseLine.X
            - natural.Single(letter => letter.Value == "1").StartBaseLine.X;
        double styledAdvance = styled.Single(letter => letter.Value == "2").StartBaseLine.X
            - styled.Single(letter => letter.Value == "1").StartBaseLine.X;
        Assert.InRange(Math.Abs(styledAdvance - naturalAdvance * scale / 100D - 1D), 0D, 0.03D);
        // A right-aligned marker fits inside the same hanging-indent column.
        Assert.InRange(Math.Abs(styled.Single(letter => letter.Value == "B").StartBaseLine.X
            - natural.Single(letter => letter.Value == "B").StartBaseLine.X), 0D, 2D);
    }

    [Theory]
    [InlineData(false, "PAGE")]
    [InlineData(true, "PAGE")]
    [InlineData(false, "NUMPAGES")]
    [InlineData(true, "NUMPAGES")]
    [InlineData(false, "SECTIONPAGES")]
    [InlineData(true, "SECTIONPAGES")]
    public void PageFieldResultDirectSpacingStylesTheGeneratedToken(bool complex, string instruction) {
        double natural = FieldFollowingTextOrigin(complex, instruction, false);
        double styled = FieldFollowingTextOrigin(complex, instruction, true);
        // Helvetica's single page-number digit advances 6.672 points at 12pt.
        Assert.InRange(Math.Abs(styled - natural - (6.672D + 1D)), 0D, 0.03D);
    }

    private static Letter[] RenderList(string route, int scale, int spacing) {
        using WordDocument document = WordDocument.Create();
        WordList list = document.AddList(WordListStyle.Numbered);
        WordListLevel level = list.Numbering.Levels[0];
        level.StartNumberingValue = 12;
        level.LevelText = "%1.";
        level.LevelJustification = WordListLevelAlignment.Right;
        level.LevelSuffix = WordListLevelSuffix.Tab;
        level.IndentationLeft = 1080;
        level.IndentationHanging = 720;
        WordParagraph item = list.AddItem("BODY");
        WordParagraph paragraph = item;
        if (route == "cell") paragraph = document.AddTable(1, 1).Rows[0].Cells[0].Paragraphs[0];
        else if (route == "header") paragraph = document.HeaderDefaultOrCreate.AddParagraph();
        else if (route == "footer") paragraph = document.FooterDefaultOrCreate.AddParagraph();
        if (!ReferenceEquals(item, paragraph)) {
            paragraph._paragraph.RemoveAllChildren();
            foreach (var child in item._paragraph.ChildElements) paragraph._paragraph.Append(child.CloneNode(true));
            item._paragraph.Remove();
            document.AddParagraph("Main");
        }
        Styles styles = document._wordprocessingDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!;
        styles.Append(new Style(new StyleRunProperties(new RunFonts { Ascii = "Arial", HighAnsi = "Arial" }, new Spacing { Val = spacing },
            new CharacterScale { Val = scale }, new FontSize { Val = "24" })) {
            StyleId = "MarkerSpacing", Type = StyleValues.Paragraph
        });
        paragraph._paragraph.ParagraphProperties!.ParagraphStyleId = new ParagraphStyleId { Val = "MarkerSpacing" };
        if (route == "inline") paragraph.PageBreakBefore = true;
        Assert.Empty(document.ValidateDocument());
        using var pdf = PdfPigDocument.Open(document.ToPdfBytes(Options()));
        return pdf.GetPages().SelectMany(page => page.Letters).ToArray();
    }

    private static double FieldFollowingTextOrigin(bool complex, string instruction, bool styled) {
        using WordDocument document = WordDocument.Create();
        document.AddParagraph("Main");
        WordParagraph paragraph = document.HeaderDefaultOrCreate.AddParagraph();
        var result = new Run(new RunProperties(new RunFonts { Ascii = "Arial", HighAnsi = "Arial" }, new Caps(),
            new Spacing { Val = styled ? 20 : 0 }, new CharacterScale { Val = styled ? 200 : 100 },
            new FontSize { Val = "24" }), new Text("888"));
        if (complex) paragraph._paragraph.Append(new Run(new FieldChar { FieldCharType = FieldCharValues.Begin }),
            new Run(new FieldCode(" " + instruction + " ")), new Run(new FieldChar { FieldCharType = FieldCharValues.Separate }),
            result, new Run(new FieldChar { FieldCharType = FieldCharValues.End }));
        else paragraph._paragraph.Append(new SimpleField(result) { Instruction = " " + instruction + " " });
        paragraph._paragraph.Append(new Run(new RunProperties(new RunFonts { Ascii = "Arial", HighAnsi = "Arial" },
            new FontSize { Val = "24" }), new Text("X")));
        using var pdf = PdfPigDocument.Open(document.ToPdfBytes(Options()));
        var letters = pdf.GetPage(1).Letters;
        Assert.Equal(1, letters.Count(letter => letter.Value == "1"));
        Assert.DoesNotContain(letters, letter => letter.Value == "8");
        return letters.Single(letter => letter.Value == "X").StartBaseLine.X;
    }

    private static WordToPdfOptions Options() => new() {
        IncludePageNumbers = false, ResourcePolicy = PdfResourcePolicy.CreatePortableDeterministic()
    };
}
