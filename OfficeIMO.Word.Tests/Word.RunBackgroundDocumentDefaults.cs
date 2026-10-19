using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Pdf;
using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using Xunit;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;

namespace OfficeIMO.Tests;

public sealed class WordRunBackgroundDocumentDefaultsTests {
    [Theory]
    [InlineData("document", "FF0000", "yellow")]
    [InlineData("paragraph-base", "00FF00", "green")]
    [InlineData("paragraph", "0000FF", "blue")]
    [InlineData("character-base", "00FFFF", "cyan")]
    [InlineData("character", "FF00FF", "magenta")]
    [InlineData("direct", "FFFFFF", "white")]
    [InlineData("off", null, "none")]
    [InlineData("clear", null, "none")]
    public void RunBackgroundSourcesRespectPriorityAndExplicitResets(string level, string? fill, string highlight) {
        using WordDocument document = WordDocument.Create();
        Styles styles = document.MainDocumentPartRoot.StyleDefinitionsPart!.Styles!;
        styles.DocDefaults = new DocDefaults(new RunPropertiesDefault(new RunPropertiesBaseStyle(
            Shade("FF0000"), new Highlight { Val = HighlightColorValues.Yellow })));
        WordParagraph run = document.AddParagraph("Background");
        string[] levels = { "document", "paragraph-base", "paragraph", "character-base", "character", "direct", "off", "clear" };
        int depth = Array.IndexOf(levels, level);
        if (depth >= 1) {
            styles.Append(new Style(new StyleRunProperties(Shade("00FF00"), new Highlight { Val = HighlightColorValues.Green })) {
                Type = StyleValues.Paragraph, StyleId = "BackgroundParagraphBase"
            });
            styles.Append(new Style(new BasedOn { Val = "BackgroundParagraphBase" }) {
                Type = StyleValues.Paragraph, StyleId = "BackgroundParagraph"
            });
            run.SetStyleId("BackgroundParagraph");
        }
        if (depth >= 2) styles.Elements<Style>().Single(style => style.StyleId == "BackgroundParagraph")
            .StyleRunProperties = new StyleRunProperties(Shade("0000FF"), new Highlight { Val = HighlightColorValues.Blue });
        if (depth >= 3) {
            styles.Append(new Style(new StyleRunProperties(Shade("00FFFF"), new Highlight { Val = HighlightColorValues.Cyan })) {
                Type = StyleValues.Character, StyleId = "BackgroundCharacterBase"
            });
            styles.Append(new Style(new BasedOn { Val = "BackgroundCharacterBase" }) {
                Type = StyleValues.Character, StyleId = "BackgroundCharacter"
            });
            run.SetCharacterStyleId("BackgroundCharacter");
        }
        if (depth >= 4) styles.Elements<Style>().Single(style => style.StyleId == "BackgroundCharacter")
            .StyleRunProperties = new StyleRunProperties(Shade("FF00FF"), new Highlight { Val = HighlightColorValues.Magenta });
        if (depth >= 5) {
            run.RunShadingFillColorHex = "FFFFFF";
            run.Highlight = WordHighlightColor.White;
        }
        if (level is "off" or "clear") {
            run._runProperties!.Highlight!.Val = HighlightColorValues.None;
            run._runProperties.Shading = level == "off"
                ? new Shading { Val = ShadingPatternValues.Nil, Fill = "FFFFFF" }
                : Shade("auto");
        }

        using WordDocument loaded = WordDocument.Load(new MemoryStream(document.ToBytes()));
        WordParagraph loadedRun = Assert.Single(loaded.Paragraphs);
        string before = loaded.MainDocumentPartRoot.StyleDefinitionsPart!.Styles!.OuterXml;
        Assert.Equal(fill, WordRunFormatting.ResolveRunShadingFillColorHex(loadedRun));
        Assert.Equal(new HighlightColorValues(highlight), WordRunFormatting.ResolveRunHighlight(loadedRun));
        Assert.Equal(before, loaded.MainDocumentPartRoot.StyleDefinitionsPart.Styles.OuterXml);
    }

    [Theory]
    [InlineData("body", false)]
    [InlineData("link", false)]
    [InlineData("header", false)]
    [InlineData("cell", false)]
    [InlineData("body", true)]
    [InlineData("link", true)]
    [InlineData("header", true)]
    [InlineData("cell", true)]
    public void NativePdfPaintsDocumentDefaultRunBackground(string story, bool highlight) {
        using WordDocument document = WordDocument.Create();
        var defaults = new RunPropertiesBaseStyle(Shade("FF0000"));
        if (highlight) defaults.Append(new Highlight { Val = HighlightColorValues.Yellow });
        document.MainDocumentPartRoot.StyleDefinitionsPart!.Styles!.DocDefaults =
            new DocDefaults(new RunPropertiesDefault(defaults));
        WordParagraph paragraph;
        if (story == "header") {
            document.AddHeadersAndFooters();
            document.AddParagraph("BODY");
            paragraph = document.Header.Default.AddParagraph("BACKGROUND");
        } else if (story == "cell") {
            paragraph = document.AddTable(1, 1).Rows[0].Cells[0].Paragraphs[0];
            paragraph.Text = "BACKGROUND";
        } else {
            paragraph = document.AddParagraph(story == "link" ? string.Empty : "BACKGROUND");
            if (story == "link") paragraph.AddHyperLink("BACKGROUND", new Uri("https://example.com/"));
        }
        paragraph.FontFamily = "Arial";
        paragraph.FontSize = 12;

        using WordDocument loaded = WordDocument.Load(new MemoryStream(document.ToBytes()));
        using PdfPigDocument pdf = PdfPigDocument.Open(loaded.ToPdfBytes(new WordToPdfOptions {
            IncludePageNumbers = false, ResourcePolicy = PdfResourcePolicy.CreatePortableDeterministic()
        }));
        var page = pdf.GetPage(1);
        var letters = page.Letters.Where(letter => letter.Value.Length == 1 && "BACKGROUND".Contains(letter.Value))
            .GroupBy(letter => Math.Round(letter.StartBaseLine.Y, 1))
            .Select(line => line.OrderBy(letter => letter.StartBaseLine.X).ToArray())
            .Single(line => string.Concat(line.Select(letter => letter.Value)) == "BACKGROUND");
        Assert.Contains(page.Paths, path => path.IsFilled && path.FillColor.ToRGBValues() == (highlight ? (1D, 1D, 0D) : (1D, 0D, 0D)) &&
            path.GetBoundingRectangle() is { } bounds &&
            bounds.Left <= letters[0].StartBaseLine.X && bounds.Right >= letters[0].StartBaseLine.X &&
            bounds.Bottom <= letters[0].StartBaseLine.Y && bounds.Top >= letters[0].StartBaseLine.Y);
    }


    [Theory]
    [InlineData(false, false)]
    [InlineData(true, false)]
    [InlineData(false, true)]
    [InlineData(true, true)]
    public void NativeTableHighlightsOverrideDocumentDefaultsAfterReopen(bool conditional, bool reset) {
        using WordDocument document = WordDocument.Create();
        Styles styles = document.MainDocumentPartRoot.StyleDefinitionsPart!.Styles!;
        styles.DocDefaults = new DocDefaults(new RunPropertiesDefault(new RunPropertiesBaseStyle(
            new Highlight { Val = HighlightColorValues.Yellow })));
        var style = new Style { StyleId = "BlueTableHighlight", Type = StyleValues.Table };
        if (conditional) style.Append(new TableStyleProperties(new RunPropertiesBaseStyle(
            new Highlight { Val = HighlightColorValues.Blue })) { Type = TableStyleOverrideValues.FirstRow });
        else style.StyleRunProperties = new StyleRunProperties(new Highlight { Val = HighlightColorValues.Blue });
        styles.Append(style);
        WordTable table = document.AddTable(1, 1);
        table._tableProperties!.TableStyle = new TableStyle { Val = "BlueTableHighlight" };
        table.ConditionalFormattingFirstRow = true;
        WordParagraph run = table.Rows[0].Cells[0].Paragraphs[0];
        run.Text = "Table highlight";
        if (reset) {
            styles.Append(new Style(new StyleRunProperties(new Highlight { Val = HighlightColorValues.None })) {
                StyleId = "NoHighlight", Type = StyleValues.Paragraph
            });
            run.SetStyleId("NoHighlight");
        }
        using WordDocument loaded = WordDocument.Load(new MemoryStream(document.ToBytes()));
        using PdfPigDocument pdf = PdfPigDocument.Open(loaded.ToPdfBytes(new WordToPdfOptions {
            IncludePageNumbers = false, ResourcePolicy = PdfResourcePolicy.CreatePortableDeterministic()
        }));
        Assert.Equal(!reset, pdf.GetPage(1).Paths.Any(path => path.IsFilled && path.FillColor.ToRGBValues() == (0D, 0D, 1D)));
        Assert.DoesNotContain(pdf.GetPage(1).Paths, path => path.IsFilled && path.FillColor.ToRGBValues() == (1D, 1D, 0D));
    }

    private static Shading Shade(string fill) => new Shading { Val = ShadingPatternValues.Clear, Fill = fill };

    [Theory]
    [InlineData(false, false)]
    [InlineData(false, true)]
    [InlineData(true, false)]
    [InlineData(true, true)]
    public void StyleBackgroundResetsBlockDocumentDefaults(bool characterStyle, bool nil) {
        using WordDocument document = WordDocument.Create();
        Styles styles = document.MainDocumentPartRoot.StyleDefinitionsPart!.Styles!;
        styles.DocDefaults = new DocDefaults(new RunPropertiesDefault(new RunPropertiesBaseStyle(
            Shade("FF0000"), new Highlight { Val = HighlightColorValues.Yellow })));
        styles.Append(new Style(new StyleRunProperties(
            nil ? new Shading { Val = ShadingPatternValues.Nil, Fill = "FF0000" } : Shade("auto"),
            new Highlight { Val = HighlightColorValues.None })) {
            Type = characterStyle ? StyleValues.Character : StyleValues.Paragraph, StyleId = "ResetBackground"
        });
        WordParagraph paragraph = document.AddParagraph("Reset");
        if (characterStyle) paragraph.SetCharacterStyleId("ResetBackground");
        else paragraph.SetStyleId("ResetBackground");

        Assert.Null(WordRunFormatting.ResolveRunShadingFillColorHex(paragraph));
        Assert.Equal(HighlightColorValues.None, WordRunFormatting.ResolveRunHighlight(paragraph));
    }
}
