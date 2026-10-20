using OfficeIMO.Pdf;
using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using W = DocumentFormat.OpenXml.Wordprocessing;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using Xunit;

namespace OfficeIMO.Tests;

public partial class Word {
    [Theory]
    [InlineData("body")]
    [InlineData("cell")]
    [InlineData("header")]
    [InlineData("footer")]
    public void SaveAsPdf_UndeclaredLineSpacingGivesEmptyAndVisibleParagraphsTheSameNaturalHeight(string story) {
        using WordDocument empty = CreateNaturalEmptyLineDocument(story, visibleMiddle: false);
        using WordDocument visible = CreateNaturalEmptyLineDocument(story, visibleMiddle: true);
        using var emptyPdf = PdfPigDocument.Open(ExportNaturalEmptyLineDocument(empty));
        using var visiblePdf = PdfPigDocument.Open(ExportNaturalEmptyLineDocument(visible));
        var emptyPage = Assert.Single(emptyPdf.GetPages());
        var visiblePage = Assert.Single(visiblePdf.GetPages());
        var emptyA = Assert.Single(emptyPage.Letters, letter => letter.Value == "A");
        var emptyB = Assert.Single(emptyPage.Letters, letter => letter.Value == "B");
        var visibleA = Assert.Single(visiblePage.Letters, letter => letter.Value == "A");
        var visibleB = Assert.Single(visiblePage.Letters, letter => letter.Value == "B");
        // The embedded test font has a 1.2-em natural line box. A 12pt
        // paragraph mark consumes the same 14.4pt as a visible 12pt line;
        // empty source paragraphs cannot use a separate generic multiplier.
        Assert.Equal(28.8D, emptyA.StartBaseLine.Y - emptyB.StartBaseLine.Y, 3);
        Assert.Equal(visibleA.StartBaseLine.Y, emptyA.StartBaseLine.Y, 3);
        Assert.Equal(visibleB.StartBaseLine.Y, emptyB.StartBaseLine.Y, 3);
        if (story == "header") {
            var emptyBody = Assert.Single(emptyPage.Letters, letter => letter.Value == "C");
            var visibleBody = Assert.Single(visiblePage.Letters, letter => letter.Value == "C");
            Assert.Equal(visibleBody.StartBaseLine.Y, emptyBody.StartBaseLine.Y, 3);
        }
    }

    private static WordDocument CreateNaturalEmptyLineDocument(string story, bool visibleMiddle) {
        WordDocument document = WordDocument.Create();
        W.Styles styles = document._wordprocessingDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!;
        // An explicitly empty paragraph-defaults element removes dependence on
        // application-supplied template defaults, without declaring line units.
        styles.DocDefaults = new W.DocDefaults(
            new W.RunPropertiesDefault(new W.RunPropertiesBaseStyle(
                new W.RunFonts { Ascii = "Arial", HighAnsi = "Arial" }, new W.FontSize { Val = "24" })),
            new W.ParagraphPropertiesDefault());
        foreach (W.Style style in styles.Elements<W.Style>()) style.RemoveAllChildren<W.StyleParagraphProperties>();
        WordTableCell? cell = story == "cell" ? document.AddTable(1, 1).Rows[0].Cells[0] : null;
        WordHeaderFooter? running = story switch {
            "header" => document.HeaderDefaultOrCreate,
            "footer" => document.FooterDefaultOrCreate,
            _ => null
        };
        if (running != null) foreach (var element in running.ChildElements.ToArray()) element.Remove();
        WordParagraph Add(string text) => running != null ? running.AddParagraph(text) : cell != null
            ? cell.AddParagraph(text, removeExistingParagraphs: cell.Paragraphs.Count == 1 && cell.Paragraphs[0].Text.Length == 0)
            : document.AddParagraph(text);
        foreach (string text in new[] { "A", visibleMiddle ? "M" : string.Empty, "B" }) {
            WordParagraph paragraph = Add(text);
            paragraph._paragraph.ParagraphProperties = new W.ParagraphProperties(
                new W.SpacingBetweenLines { Before = "0", After = "0" },
                new W.ParagraphMarkRunProperties(new W.RunFonts { Ascii = "Arial", HighAnsi = "Arial" }, new W.FontSize { Val = "24" }));
            paragraph.FontFamily = "Arial";
            paragraph.FontSize = 12;
        }
        if (running != null) document.AddParagraph("C");
        return document;
    }

    private static byte[] ExportNaturalEmptyLineDocument(WordDocument document) {
        var options = new PdfOptions();
        options.RegisterNamedFontFamily(new PdfEmbeddedFontFamily("Arial", CreateBaselineMetricFont()));
        return document.ToPdfBytes(new WordToPdfOptions {
            IncludePageNumbers = false,
            ResourcePolicy = PdfResourcePolicy.CreatePortableDeterministic(),
            PdfOptions = options,
            Margins = PageMargins.Uniform(30D),
            PageSize = new PageSize(300D, 300D)
        });
    }
}
