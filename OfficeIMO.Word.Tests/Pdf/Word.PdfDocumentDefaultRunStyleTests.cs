using DocumentFormat.OpenXml;
using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using W = DocumentFormat.OpenXml.Wordprocessing;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class WordPdfDocumentDefaultRunStyleTests {
    [Theory]
    [InlineData("body")]
    [InlineData("cell")]
    [InlineData("header")]
    [InlineData("footer")]
    public void DocumentDefaultFormattingMatchesExplicitFormattingInEveryStory(string route) {
        using WordDocument inherited = CreateDocument(route, inheritFormatting: true, overrideFormatting: false);
        using WordDocument explicitFormatting = CreateDocument(route, inheritFormatting: false, overrideFormatting: false);
        AssertEquivalentPdf(explicitFormatting, inherited, "DEFAULT");
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void ExplicitRunAndStyleOverridesStopDocumentDefaultFormatting(bool characterStyle) {
        using WordDocument inherited = CreateDocument("body", inheritFormatting: true, overrideFormatting: true, characterStyle);
        using WordDocument control = CreateDocument("body", inheritFormatting: false, overrideFormatting: true, characterStyle);
        AssertEquivalentPdf(control, inherited, "Default");
    }

    private static WordDocument CreateDocument(string route, bool inheritFormatting, bool overrideFormatting, bool characterStyle = false) {
        WordDocument document = WordDocument.Create();
        W.Styles styles = document._wordprocessingDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!;
        var defaults = new W.RunPropertiesBaseStyle(
            new W.RunFonts { Ascii = "Arial", HighAnsi = "Arial" }, new W.FontSize { Val = "24" });
        if (inheritFormatting) defaults.Append(Formatting(enabled: true));
        styles.DocDefaults = new W.DocDefaults(new W.RunPropertiesDefault(defaults));
        W.Style normal = styles.Elements<W.Style>().Single(style => style.StyleId?.Value == "Normal");
        normal.StyleRunProperties = new W.StyleRunProperties();
        WordParagraph paragraph = route switch {
            "cell" => document.AddTable(1, 1).Rows[0].Cells[0].Paragraphs[0],
            "header" => document.HeaderDefaultOrCreate.AddParagraph(),
            "footer" => document.FooterDefaultOrCreate.AddParagraph(),
            _ => document.AddParagraph()
        };
        var properties = new W.RunProperties();
        if (!inheritFormatting || overrideFormatting) {
            if (characterStyle) {
                styles.Append(new W.Style(new W.StyleRunProperties(Formatting(!overrideFormatting))) {
                    Type = W.StyleValues.Character, StyleId = "ExplicitRunStyle"
                });
                properties.RunStyle = new W.RunStyle { Val = "ExplicitRunStyle" };
            } else {
                properties.Append(Formatting(!overrideFormatting));
            }
        }
        paragraph._paragraph.Append(new W.Run(properties, new W.Text("Default")));
        if (route is "header" or "footer") document.AddParagraph();
        return document;
    }

    private static OpenXmlElement[] Formatting(bool enabled) => new OpenXmlElement[] {
        new W.Bold { Val = enabled }, new W.Italic { Val = enabled },
        new W.Caps { Val = enabled }, new W.Strike { Val = enabled },
        new W.Underline { Val = enabled ? W.UnderlineValues.Single : W.UnderlineValues.None },
        new W.Color { Val = enabled ? "336699" : "auto" },
        new W.VerticalTextAlignment { Val = enabled ? W.VerticalPositionValues.Superscript : W.VerticalPositionValues.Baseline }
    };

    private static void AssertEquivalentPdf(WordDocument expected, WordDocument actual, string expectedText) {
        using var reference = PdfPigDocument.Open(expected.ToPdfBytes(new WordToPdfOptions { IncludePageNumbers = false }));
        using var candidate = PdfPigDocument.Open(actual.ToPdfBytes(new WordToPdfOptions { IncludePageNumbers = false }));
        Assert.Equal(1, reference.NumberOfPages);
        Assert.Equal(reference.NumberOfPages, candidate.NumberOfPages);
        var referencePage = reference.GetPage(1);
        var candidatePage = candidate.GetPage(1);
        Assert.Equal(expectedText, string.Concat(candidatePage.Letters.Select(letter => letter.Value)));
        Assert.Equal(referencePage.Letters.Select(letter =>
            (letter.Value, letter.FontName, letter.FontSize, letter.StartBaseLine, letter.Color.ToRGBValues())),
            candidatePage.Letters.Select(letter =>
            (letter.Value, letter.FontName, letter.FontSize, letter.StartBaseLine, letter.Color.ToRGBValues())));
        Assert.Equal(referencePage.Paths.Count, candidatePage.Paths.Count);
    }
}
