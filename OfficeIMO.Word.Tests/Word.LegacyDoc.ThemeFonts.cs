using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests;

public partial class Word {
    [Theory]
    [InlineData("body")]
    [InlineData("header")]
    [InlineData("footer")]
    [InlineData("footnote")]
    [InlineData("endnote")]
    [InlineData("comment")]
    public void LegacyDoc_ThemeFontsPreserveParagraphStylesAcrossStoriesAndRepeatedSaves(string story) {
        using WordDocument source = WordDocument.Create();
        MainDocumentPart main = ConfigureNativeThemeFonts(source);
        source.AddParagraph("Body control");
        WordParagraph paragraph = story switch {
            "body" => source.AddParagraph("Story heading"),
            "header" => source.HeaderDefaultOrCreate.AddParagraph("Story heading"),
            "footer" => source.FooterDefaultOrCreate.AddParagraph("Story heading"),
            "footnote" => source.AddParagraph("Reference").AddFootNote("Story heading").FootNote!.Paragraphs!.Single(item => item.Text == "Story heading"),
            "endnote" => source.AddParagraph("Reference").AddEndNote("Story heading").EndNote!.Paragraphs!.Single(item => item.Text == "Story heading"),
            _ => AddCharacterScaleComment(source)
        };
        paragraph.SetStyleId("Heading1");
        Style heading = main.StyleDefinitionsPart!.Styles!.Elements<Style>().Single(style => style.StyleId == "Heading1");
        heading.StyleRunProperties = new StyleRunProperties(NativeMajorThemeFonts(), new FontSize { Val = "32" });
        string before = main.StyleDefinitionsPart.Styles.OuterXml;
        byte[] bytes = source.ToBytes(WordFileFormat.Doc);
        for (int save = 0; save < 2; save++) {
            using WordDocument loaded = WordDocument.Load(new MemoryStream(bytes));
            Style actual = loaded._wordprocessingDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!
                .Elements<Style>().Single(style => style.StyleId == "Heading1");
            Assert.Equal("Courier New", actual.StyleRunProperties?.RunFonts?.Ascii?.Value);
            Assert.Equal("32", actual.StyleRunProperties?.FontSize?.Val?.Value);
            Assert.Empty(loaded.ValidateDocument());
            if (save == 0) bytes = loaded.ToBytes(WordFileFormat.Doc);
        }
        Assert.Equal(before, main.StyleDefinitionsPart.Styles.OuterXml);
        Assert.Empty(source.ValidateDocument());
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void LegacyDoc_ThemeFontsResolveDirectRunsAheadOfLiteralFallbacks(bool minor) {
        using WordDocument source = WordDocument.Create();
        ConfigureNativeThemeFonts(source);
        WordParagraph paragraph = source.AddParagraph("Direct theme font");
        RunFonts fonts = NativeMajorThemeFonts();
        if (minor) { fonts.AsciiTheme = ThemeFontValues.MinorAscii; fonts.HighAnsiTheme = ThemeFontValues.MinorHighAnsi; }
        paragraph._run!.RunProperties = new RunProperties(fonts);
        string before = paragraph._paragraph.OuterXml;
        using WordDocument loaded = WordDocument.Load(new MemoryStream(source.ToBytes(WordFileFormat.Doc)));
        Assert.Equal(minor ? "Arial" : "Courier New", loaded.Paragraphs[0]._run!.RunProperties?.RunFonts?.Ascii?.Value);
        Assert.Equal(before, paragraph._paragraph.OuterXml);
        Assert.Empty(loaded.ValidateDocument());
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void LegacyDoc_ThemeFontsResolveTableStyleAndConditionalRunFormatting(bool conditional) {
        using WordDocument source = WordDocument.Create();
        MainDocumentPart main = ConfigureNativeThemeFonts(source);
        Style style = new() { Type = StyleValues.Table, StyleId = "ThemeTable", CustomStyle = true };
        style.Append(new StyleName { Val = "Theme table" }, new BasedOn { Val = "TableNormal" });
        if (conditional) style.Append(new TableStyleProperties(new RunPropertiesBaseStyle(NativeMajorThemeFonts())) { Type = TableStyleOverrideValues.FirstRow });
        else style.StyleRunProperties = new StyleRunProperties(NativeMajorThemeFonts());
        main.StyleDefinitionsPart!.Styles!.Append(style);
        WordTable table = source.AddTable(1, 1, WordTableStyle.TableNormal);
        table._tableProperties!.TableStyle = new TableStyle { Val = "ThemeTable" };
        table._tableProperties.TableLook = new TableLook { FirstRow = true, NoHorizontalBand = true, NoVerticalBand = true };
        table.Rows[0].Cells[0].AddParagraph("Theme cell", removeExistingParagraphs: true);
        using WordDocument loaded = WordDocument.Load(new MemoryStream(source.ToBytes(WordFileFormat.Doc)));
        Run run = loaded.Tables[0]._table.Descendants<Run>().Single(item => item.InnerText == "Theme cell");
        Assert.Equal("Courier New", run.RunProperties?.RunFonts?.Ascii?.Value);
        Assert.Empty(source.ValidateDocument());
        Assert.Empty(loaded.ValidateDocument());
    }

    [Fact]
    public void LegacyDoc_ThemeFontsSurviveDocumentDefaultStyleMaterialization() {
        using WordDocument source = WordDocument.Create();
        MainDocumentPart main = ConfigureNativeThemeFonts(source);
        Styles styles = main.StyleDefinitionsPart!.Styles!;
        styles.DocDefaults!.RunPropertiesDefault!.RunPropertiesBaseStyle!.AddChild(new CharacterScale { Val = 150L }, true);
        styles.DocDefaults.ParagraphPropertiesDefault!.ParagraphPropertiesBaseStyle!.SpacingBetweenLines = new SpacingBetweenLines { After = "160" };
        styles.Append(new Style(new StyleName { Val = "Theme root" }, new StyleRunProperties(NativeMajorThemeFonts())) {
            Type = StyleValues.Paragraph, StyleId = "ThemeRoot", CustomStyle = true
        });
        source.AddParagraph("Root theme font").SetStyleId("ThemeRoot");
        using WordDocument loaded = WordDocument.Load(new MemoryStream(source.ToBytes(WordFileFormat.Doc)));
        Style actual = loaded._wordprocessingDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!.Elements<Style>()
            .Single(style => style.StyleName?.Val?.Value == "Theme root");
        Assert.Equal("Courier New", actual.StyleRunProperties?.RunFonts?.Ascii?.Value);
        Assert.Equal(150L, actual.StyleRunProperties?.CharacterScale?.Val?.Value);
        Assert.Equal("160", actual.StyleParagraphProperties?.SpacingBetweenLines?.After?.Value);
    }

    [Fact]
    public void LegacyDoc_ThemeFontsRetainTheSingleFamilyConstraint() {
        using WordDocument source = WordDocument.Create();
        ConfigureNativeThemeFonts(source);
        WordParagraph paragraph = source.AddParagraph("Different script families");
        RunFonts fonts = NativeMajorThemeFonts();
        fonts.HighAnsiTheme = ThemeFontValues.MinorHighAnsi;
        paragraph._run!.RunProperties = new RunProperties(fonts);
        Assert.Throws<NotSupportedException>(() => source.ToBytes(WordFileFormat.Doc));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void LegacyDoc_ThemeFontsPreserveListLevelAndInstanceOverrideFonts(bool instanceOverride) {
        using WordDocument source = CreateNativeListDefinitionControl();
        ConfigureNativeThemeFonts(source);
        Numbering numbering = source._wordprocessingDocument.MainDocumentPart!.NumberingDefinitionsPart!.Numbering!;
        Level level = numbering.Elements<AbstractNum>().Single().Elements<Level>().First();
        if (instanceOverride) {
            level = (Level)level.CloneNode(true);
            numbering.Elements<NumberingInstance>().Single().Append(new LevelOverride(level) { LevelIndex = 0 });
        }
        level.NumberingSymbolRunProperties = new NumberingSymbolRunProperties(NativeMajorThemeFonts());
        string before = numbering.OuterXml;
        using WordDocument loaded = WordDocument.Load(new MemoryStream(source.ToBytes(WordFileFormat.Doc)));
        Numbering actual = loaded._wordprocessingDocument.MainDocumentPart!.NumberingDefinitionsPart!.Numbering!;
        Level actualLevel = instanceOverride ? actual.Elements<NumberingInstance>().Single().Elements<LevelOverride>().Single().Level!
            : actual.Elements<AbstractNum>().Single().Elements<Level>().First();
        Assert.Equal("Courier New", actualLevel.NumberingSymbolRunProperties?.RunFonts?.Ascii?.Value);
        Assert.Equal(before, numbering.OuterXml);
        Assert.Empty(loaded.ValidateDocument());
    }

    private static MainDocumentPart ConfigureNativeThemeFonts(WordDocument source) {
        MainDocumentPart main = source._wordprocessingDocument.MainDocumentPart!;
        var scheme = main.ThemePart!.Theme!.ThemeElements!.FontScheme!;
        scheme.MajorFont!.LatinFont!.Typeface = "Courier New";
        scheme.MinorFont!.LatinFont!.Typeface = "Arial";
        return main;
    }

    private static RunFonts NativeMajorThemeFonts() => new() {
        Ascii = "Times New Roman", HighAnsi = "Times New Roman",
        AsciiTheme = ThemeFontValues.MajorAscii, HighAnsiTheme = ThemeFontValues.MajorHighAnsi
    };
}
