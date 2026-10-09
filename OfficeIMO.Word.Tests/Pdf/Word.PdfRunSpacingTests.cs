using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Pdf;
using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using UglyToad.PdfPig.Content;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class WordPdfRunSpacingTests {
    [Theory]
    [InlineData("body")]
    [InlineData("cell")]
    [InlineData("header")]
    [InlineData("footer")]
    [InlineData("link")]
    public void DirectRunMetricsReachEveryTextStory(string route) {
        Letter[] natural = Render(route, "direct", false);
        Letter[] styled = Render(route, "direct", true);
        AssertGeometry(natural, styled, 50D, 1D);
    }

    [Theory]
    [InlineData("body")]
    [InlineData("cell")]
    [InlineData("header")]
    [InlineData("footer")]
    public void AdjacentNormalRunRestoresItsWidthAndTracking(string route) {
        Letter[] natural = Render(route, "direct", false);
        Letter[] mixed = Render(route, "direct", true, mixed: true);
        double firstTwoAdvance = natural[2].StartBaseLine.X - natural[0].StartBaseLine.X;
        Assert.InRange(Math.Abs((mixed[2].StartBaseLine.X - mixed[0].StartBaseLine.X) - (firstTwoAdvance / 2D + 2D)), 0D, 0.02D);
        Assert.InRange(Math.Abs((mixed[3].StartBaseLine.X - mixed[2].StartBaseLine.X) - (natural[3].StartBaseLine.X - natural[2].StartBaseLine.X)), 0D, 0.02D);
    }

    [Theory]
    [InlineData("document", 50D, 1D)]
    [InlineData("paragraph", 200D, -0.5D)]
    [InlineData("character", 100D, 0D)]
    [InlineData("direct", 50D, 1D)]
    [InlineData("conditional", 50D, 1D)]
    public void InheritedRunMetricsRespectOverridesAndExplicitResets(string level, double width, double spacing) {
        string route = level == "conditional" ? "cell" : "body";
        AssertGeometry(Render(route, level, false), Render(route, level, true), width, spacing);
    }

    [Fact]
    public void CharacterScaleRoundTripsAndNullRestoresInheritance() {
        using WordDocument document = WordDocument.Create();
        WordParagraph run = document.AddParagraph("Authored width");
        run.CharacterScale = 50;
        run.Spacing = 20;
        using WordDocument loaded = WordDocument.Load(new MemoryStream(document.ToBytes()));
        Assert.Equal(50, loaded.Paragraphs[0].CharacterScale);
        Assert.Equal(20, loaded.Paragraphs[0].Spacing);
        loaded.Paragraphs[0].CharacterScale = 100;
        Assert.Equal(100, loaded.Paragraphs[0].CharacterScale);
        loaded.Paragraphs[0].CharacterScale = null;
        Assert.Null(loaded.Paragraphs[0].CharacterScale);
        Assert.Equal(20, loaded.Paragraphs[0].Spacing);
        Assert.Throws<ArgumentOutOfRangeException>(() => run.CharacterScale = 0);
        Assert.Throws<ArgumentOutOfRangeException>(() => run.CharacterScale = 601);
    }

    private static void AssertGeometry(Letter[] natural, Letter[] styled, double width, double spacing) {
        Assert.Equal("ABCD", string.Concat(styled.Select(letter => letter.Value)));
        for (int i = 1; i < styled.Length; i++) {
            double expected = (natural[i].StartBaseLine.X - natural[0].StartBaseLine.X) * width / 100D + i * spacing;
            double actual = styled[i].StartBaseLine.X - styled[0].StartBaseLine.X;
            Assert.InRange(Math.Abs(expected - actual), 0D, 0.02D);
        }
        Assert.InRange(Math.Abs(natural[0].BoundingBox.Height - styled[0].BoundingBox.Height), 0D, 0.02D);
    }

    private static Letter[] Render(string route, string level, bool styled, bool mixed = false) {
        using WordDocument document = WordDocument.Create();
        Styles styles = document._wordprocessingDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!;
        styles.DocDefaults = new DocDefaults(new RunPropertiesDefault(new RunPropertiesBaseStyle(
            new RunFonts { Ascii = "Arial", HighAnsi = "Arial" }, new FontSize { Val = "24" })));
        WordParagraph paragraph;
        WordTable? table = null;
        if (route == "cell") {
            table = document.AddTable(1, 1);
            table._tableProperties!.TableStyle = null;
            paragraph = table.Rows[0].Cells[0].Paragraphs[0];
            paragraph.Text = "ABCD";
        } else if (route is "header" or "footer") {
            document.AddParagraph("BODY");
            document.AddHeadersAndFooters();
            paragraph = (route == "header" ? (WordHeaderFooter)document.Header.Default : document.Footer.Default).AddParagraph("ABCD");
        } else {
            paragraph = document.AddParagraph(route == "link" ? string.Empty : "ABCD");
            if (route == "link") paragraph.AddHyperLink("ABCD", new Uri("https://example.com/"));
        }
        paragraph.FontSize = 12;
        paragraph.FontFamily = "Arial";
        if (styled) {
            RunPropertiesBaseStyle defaults = styles.DocDefaults!.RunPropertiesDefault!.RunPropertiesBaseStyle!;
            if (level is "document" or "paragraph" or "character" or "direct") {
                defaults.Spacing = new Spacing { Val = 20 };
                defaults.CharacterScale = new CharacterScale { Val = 50 };
            }
            if (level is "paragraph" or "character" or "direct") {
                styles.Append(new Style(new StyleRunProperties(new Spacing { Val = -10 }, new CharacterScale { Val = 200 })) {
                    Type = StyleValues.Paragraph, StyleId = "WideText"
                });
                styles.Append(new Style(new BasedOn { Val = "WideText" }) { Type = StyleValues.Paragraph, StyleId = "DerivedWide" });
                paragraph._paragraph.ParagraphProperties ??= new ParagraphProperties();
                paragraph._paragraph.ParagraphProperties.ParagraphStyleId = new ParagraphStyleId { Val = "DerivedWide" };
            }
            if (level is "character" or "direct") {
                styles.Append(new Style(new StyleRunProperties(new Spacing { Val = 0 }, new CharacterScale { Val = 100 })) {
                    Type = StyleValues.Character, StyleId = "NaturalText"
                });
                paragraph._paragraph.Descendants<Run>().Last().RunProperties!.RunStyle = new RunStyle { Val = "NaturalText" };
            }
            if (level == "conditional") {
                styles.Append(new Style(new StyleRunProperties(new Spacing { Val = -10 }, new CharacterScale { Val = 200 }),
                    new TableStyleProperties(new RunPropertiesBaseStyle(new Spacing { Val = 20 }, new CharacterScale { Val = 50 })) {
                        Type = TableStyleOverrideValues.FirstRow
                    }) { Type = StyleValues.Table, StyleId = "FittedTable" });
                table!._tableProperties!.TableStyle = new TableStyle { Val = "FittedTable" };
                table.ConditionalFormattingFirstRow = true;
            }
            if (level == "direct") {
                Run run = paragraph._paragraph.Descendants<Run>().Last();
                run.RunProperties ??= new RunProperties();
                run.RunProperties.Spacing = new Spacing { Val = 20 };
                run.RunProperties.CharacterScale = new CharacterScale { Val = 50 };
            }
        }
        if (mixed) {
            paragraph.Text = "AB";
            WordParagraph normal = paragraph.AddText("CD");
            normal.FontSize = 12;
            normal.FontFamily = "Arial";
            normal.CharacterScale = 100;
            normal.Spacing = 0;
        }
        byte[] bytes = document.ToPdfBytes(new WordToPdfOptions {
            IncludePageNumbers = false, ResourcePolicy = PdfResourcePolicy.CreatePortableDeterministic()
        });
        using PdfPigDocument pdf = PdfPigDocument.Open(bytes);
        return pdf.GetPage(1).Letters.Where(letter => "ABCD".IndexOf(letter.Value, StringComparison.Ordinal) >= 0)
            .GroupBy(letter => Math.Round(letter.StartBaseLine.Y, 1))
            .Select(line => line.OrderBy(letter => letter.StartBaseLine.X).ToArray())
            .Single(line => string.Concat(line.Select(letter => letter.Value)) == "ABCD");
    }
}
