using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests;

public partial class Word {
    [Theory]
    [InlineData("auto", true)]
    [InlineData("336699", false)]
    public void ParagraphShadingExplicitFillOverridesLoadedThemeSelectors(string fill, bool tint) {
        using WordDocument source = WordDocument.Create();
        WordParagraph paragraph = source.AddParagraph("Themed");
        paragraph.ShadingFillColor = OfficeColor.Parse("112233");
        Shading themed = paragraph._paragraph.ParagraphProperties!.Shading!;
        themed.ThemeFill = ThemeColorValues.Accent2;
        if (tint) themed.ThemeFillTint = "40";
        else themed.ThemeFillShade = "80";
        using WordDocument loaded = WordDocument.Load(new MemoryStream(source.ToBytes()));
        WordParagraph actual = loaded.Paragraphs.Single();
        Assert.NotNull(actual._paragraph.ParagraphProperties!.Shading!.ThemeFill);

        if (fill == "auto") actual.ShadingFillValue = fill;
        else actual.ShadingFillColor = OfficeColor.Parse(fill);

        Shading shading = actual._paragraph.ParagraphProperties.Shading;
        Assert.Null(shading.ThemeFill);
        Assert.Null(shading.ThemeFillTint);
        Assert.Null(shading.ThemeFillShade);
        Assert.Equal(fill, shading.Fill?.Value);
        Assert.Empty(loaded.DocumentValidationErrors);
    }

    [Theory]
    [InlineData(WordFileFormat.Docx)]
    [InlineData(WordFileFormat.Doc)]
    public void LegacyDoc_ExplicitAutomaticParagraphShadingOverridesRgbStyle(WordFileFormat format) {
        using WordDocument source = WordDocument.Create();
        Styles styles = source._wordprocessingDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!;
        Style normal = styles.Elements<Style>().Single(style => style.StyleId?.Value == "Normal");
        normal.StyleParagraphProperties ??= new StyleParagraphProperties();
        normal.StyleParagraphProperties.Shading = new Shading {
            Val = ShadingPatternValues.Clear, Color = "auto", Fill = "FF0000"
        };
        source.AddParagraph("InheritsRed");
        source.AddParagraph("ClearsRed").ShadingFillValue = "auto";
        Assert.Empty(source.DocumentValidationErrors);

        using WordDocument loaded = WordDocument.Load(new MemoryStream(source.ToBytes(new WordSaveOptions { Format = format })));
        WordParagraph clear = Assert.Single(loaded.Paragraphs, paragraph => paragraph.Text == "ClearsRed");
        Assert.Equal("auto", clear._paragraph.ParagraphProperties!.Shading!.Fill!.Value);
        Assert.Null(clear.ShadingFillColor);
        Assert.Empty(loaded.DocumentValidationErrors);
    }

    [Theory]
    [InlineData("body")]
    [InlineData("table-paragraph")]
    [InlineData("header")]
    [InlineData("footnote")]
    public void LegacyDoc_ParagraphRgbShadingSurvivesNativeSaveAndReopen(string route) {
        using WordDocument source = WordDocument.Create();
        source.AddParagraph("Anchor");
        WordParagraph target;
        if (route == "table-paragraph") {
            target = source.AddTable(1, 1).Rows[0].Cells[0].Paragraphs[0];
            target.Text = "RgbShaded";
        } else if (route == "header") {
            target = source.PageSetup.Header.AddParagraph("RgbShaded");
        } else if (route == "footnote") {
            target = source.AddParagraph("Note").AddFootNote("RgbShaded").Paragraphs!.Last();
        } else target = source.AddParagraph("RgbShaded");
        target.ShadingFillColor = OfficeColor.Parse("336699");

        // An exact RGB record carries the fill, so no palette approximation is reported.
        using WordDocument loaded = WordDocument.Load(new MemoryStream(source.ToBytes(new WordSaveOptions { Format = WordFileFormat.Doc })));
        WordParagraph actual = route switch {
            "table-paragraph" => loaded.Tables.Single().Rows[0].Cells[0].Paragraphs.Single(),
            "header" => Assert.Single(loaded.PageSetup.GetHeader(WordHeaderFooterType.Default).Paragraphs, paragraph => paragraph.Text == "RgbShaded"),
            "footnote" => Assert.Single(loaded.FootNotes.SelectMany(note => note.Paragraphs!), paragraph => paragraph.Text.Contains("RgbShaded")),
            _ => Assert.Single(loaded.Paragraphs, paragraph => paragraph.Text == "RgbShaded")
        };
        Assert.Equal(OfficeColor.Parse("336699"), actual.ShadingFillColor);
        Assert.Equal(WordShadingPattern.Clear, actual.ShadingPattern);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void LegacyDoc_ParagraphStyleRgbShadingSurvivesNativeSaveAndReopen(bool customStyle) {
        using WordDocument source = WordDocument.Create();
        Styles styles = source._wordprocessingDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!;
        Style style;
        if (customStyle) {
            style = new Style(new StyleName { Val = "Rgb paragraph shading" }) {
                StyleId = "RgbParagraphShading", Type = StyleValues.Paragraph, CustomStyle = true
            };
            styles.Append(style);
        } else style = styles.Elements<Style>().Single(candidate => candidate.StyleId?.Value == "Normal");
        style.StyleParagraphProperties ??= new StyleParagraphProperties();
        style.StyleParagraphProperties.Shading = new Shading {
            Val = ShadingPatternValues.Clear, Color = "auto", Fill = "E5EFF7"
        };
        source.AddParagraph("StyleShaded").SetStyleId(style.StyleId!.Value!);

        using WordDocument loaded = WordDocument.Load(new MemoryStream(source.ToBytes(new WordSaveOptions { Format = WordFileFormat.Doc })));
        Styles loadedStyles = loaded._wordprocessingDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!;
        Style actual = loadedStyles.Elements<Style>().Single(candidate => customStyle
            ? candidate.StyleName?.Val?.Value == "Rgb paragraph shading"
            : candidate.StyleId?.Value == "Normal");
        Assert.Equal("E5EFF7", actual.StyleParagraphProperties?.Shading?.Fill?.Value);
        Assert.Equal(ShadingPatternValues.Clear, actual.StyleParagraphProperties?.Shading?.Val?.Value);
    }

    [Fact]
    public void LegacyDoc_ImportsRgbParagraphShadingFromMicrosoftWord() {
        using WordDocument source = WordDocument.Load(GetFixtureDoc(Path.Combine("ParagraphShading", "word-rgb-paragraph.doc")));
        WordParagraph shaded = Assert.Single(source.Paragraphs, paragraph => paragraph.Text.StartsWith("Nested001", StringComparison.Ordinal));
        Assert.Equal(OfficeColor.Parse("E5EFF7"), shaded.ShadingFillColor);
        Assert.Equal(WordShadingPattern.Clear, shaded.ShadingPattern);

        using WordDocument loaded = WordDocument.Load(new MemoryStream(source.ToBytes(new WordSaveOptions { Format = WordFileFormat.Doc })));
        Assert.Equal(OfficeColor.Parse("E5EFF7"), Assert.Single(loaded.Paragraphs,
            paragraph => paragraph.Text.StartsWith("Nested001", StringComparison.Ordinal)).ShadingFillColor);
    }
}
