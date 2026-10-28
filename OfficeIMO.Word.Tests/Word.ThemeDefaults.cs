using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests;

public partial class Word {
    private static Style GetStyleDefinition(WordDocument document, string styleId) =>
        document._wordprocessingDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!
            .Elements<Style>().Single(style => style.StyleId?.Value == styleId);

    private static string HalfPoints(OfficeLength size) =>
        ((int)Math.Round(size.ToPoints() * 2D, MidpointRounding.AwayFromZero)).ToString(System.Globalization.CultureInfo.InvariantCulture);

    [Theory]
    [InlineData(1)]
    [InlineData(2)]
    [InlineData(3)]
    [InlineData(4)]
    [InlineData(5)]
    [InlineData(6)]
    public void Create_HeadingStylesFollowTheDefaultThemeTypeScale(int level) {
        using WordDocument document = WordDocument.Create();
        OfficeTheme theme = OfficeTheme.Default;
        Style heading = GetStyleDefinition(document, "Heading" + level);
        OfficeLength expected = level switch {
            1 => theme.TypeScale.Heading1, 2 => theme.TypeScale.Heading2, 3 => theme.TypeScale.Heading3,
            4 => theme.TypeScale.Heading4, 5 => theme.TypeScale.Heading5, _ => theme.TypeScale.Heading6
        };

        Assert.Equal(HalfPoints(expected), heading.StyleRunProperties!.FontSize!.Val!.Value);
        Assert.NotNull(heading.StyleRunProperties.Bold);
        Assert.Equal(theme.TextStyle.HeadingColor.Resolve(theme).ToRgbHex(), heading.StyleRunProperties.Color!.Val!.Value);
        Assert.NotNull(heading.StyleParagraphProperties!.KeepNext);
        Assert.NotEqual("0", heading.StyleParagraphProperties.SpacingBetweenLines!.After!.Value);
        Assert.Equal((theme.TextStyle.GetHeadingSpacing(level).Before.ToPoints() * 20D).ToString(System.Globalization.CultureInfo.InvariantCulture),
            heading.StyleParagraphProperties.SpacingBetweenLines.Before!.Value);
    }

    [Theory]
    [InlineData(1)]
    [InlineData(2)]
    [InlineData(3)]
    [InlineData(4)]
    [InlineData(5)]
    [InlineData(6)]
    public void SavedHeadingCharacterStylesAndThemeFontsMatchTheirParagraphStyles(int level) {
        using WordDocument source = WordDocument.Create();
        source.AddParagraph("Heading text").SetCharacterStyle(
            (WordCharacterStyles)Enum.Parse(typeof(WordCharacterStyles), "Heading" + level + "Char"));
        using WordDocument document = WordDocument.Load(new MemoryStream(source.ToBytes()));
        Style heading = GetStyleDefinition(document, "Heading" + level);
        Style character = GetStyleDefinition(document, "Heading" + level + "Char");

        Assert.Equal(heading.StyleRunProperties!.FontSize!.Val!.Value, character.StyleRunProperties!.FontSize!.Val!.Value);
        Assert.Equal(heading.StyleRunProperties.Color!.Val!.Value, character.StyleRunProperties.Color!.Val!.Value);
        Assert.Equal(heading.StyleRunProperties.Bold!.Val?.Value, character.StyleRunProperties.Bold!.Val?.Value);
        Assert.Equal(heading.StyleRunProperties.RunFonts!.AsciiTheme!.Value, character.StyleRunProperties.RunFonts!.AsciiTheme!.Value);
        var fonts = document._wordprocessingDocument.MainDocumentPart!.ThemePart!.Theme.ThemeElements!.FontScheme!;
        Assert.Equal(OfficeTheme.Default.Fonts.Heading.Family, fonts.MajorFont!.LatinFont!.Typeface!.Value);
        Assert.Equal(OfficeTheme.Default.Fonts.Body.Family, fonts.MinorFont!.LatinFont!.Typeface!.Value);
        Assert.Empty(document.ValidateDocument());
    }

    [Fact]
    public void Create_DefinesTitleAndSubtitleStyles() {
        using WordDocument document = WordDocument.Create();
        Style title = GetStyleDefinition(document, "Title");
        Assert.Equal(HalfPoints(OfficeTheme.Default.TypeScale.Title), title.StyleRunProperties!.FontSize!.Val!.Value);
        Assert.NotNull(GetStyleDefinition(document, "Subtitle"));

        WordParagraph paragraph = document.AddParagraph("Report").SetStyleId("Title");
        Assert.Equal("Title", paragraph.StyleId);
    }

    [Fact]
    public void AddTable_DefaultStyleUsesThemeHeaderAndHorizontalRules() {
        using WordDocument document = WordDocument.Create();
        WordTable table = document.AddTable(3, 3);
        Assert.Equal(WordTableStyle.OfficeIMOModern, table.Style);

        Style style = GetStyleDefinition(document, "OfficeIMOModern");
        OfficeThemeTableStyle theme = OfficeTheme.Default.TableStyle;
        TableBorders borders = style.StyleTableProperties!.TableBorders!;
        Assert.Null(borders.InsideVerticalBorder);
        Assert.Null(borders.LeftBorder);
        Assert.NotNull(borders.InsideHorizontalBorder);

        TableStyleProperties header = style.Elements<TableStyleProperties>().Single(part => part.Type?.Value == TableStyleOverrideValues.FirstRow);
        Assert.Equal(theme.HeaderFill.Resolve(OfficeTheme.Default).ToRgbHex(),
            header.TableStyleConditionalFormattingTableCellProperties!.Shading!.Fill!.Value);
        Assert.NotNull(header.RunPropertiesBaseStyle!.Bold);
        Assert.NotNull(header.TableStyleConditionalFormattingTableRowProperties!.GetFirstChild<TableHeader>());
    }

    [Theory]
    [InlineData("body")]
    [InlineData("section")]
    [InlineData("header")]
    [InlineData("footer")]
    [InlineData("cell")]
    [InlineData("before")]
    [InlineData("after")]
    public void NewTableOwnersSaveTheSharedModernStyle(string owner) {
        using WordDocument document = WordDocument.Create();
        document.AddHeadersAndFooters();
        WordTable table = owner switch {
            "body" => document.AddTable(2, 2),
            "section" => document.Sections[0].AddTable(2, 2),
            "header" => document.Header.Default!.AddTable(2, 2),
            "footer" => document.Footer.Default!.AddTable(2, 2),
            "cell" => document.AddTable(1, 1).Rows[0].Cells[0].AddTable(2, 2),
            "before" => document.AddParagraph("Anchor").AddTableBefore(2, 2),
            "after" => document.AddParagraph("Anchor").AddTableAfter(2, 2),
            _ => throw new ArgumentOutOfRangeException(nameof(owner))
        };
        table.Rows[0].Cells[0].Paragraphs[0].Text = "Header";
        Assert.Equal(WordTableStyle.OfficeIMOModern, table.Style);
        using var stream = new MemoryStream(document.ToBytes());
        using var package = DocumentFormat.OpenXml.Packaging.WordprocessingDocument.Open(stream, false);
        IEnumerable<Table> saved = owner switch {
            "header" => package.MainDocumentPart!.HeaderParts.SelectMany(part => part.Header.Descendants<Table>()),
            "footer" => package.MainDocumentPart!.FooterParts.SelectMany(part => part.Footer.Descendants<Table>()),
            _ => package.MainDocumentPart!.Document.Descendants<Table>()
        };
        Table[] savedTables = saved.ToArray();
        Assert.Contains(savedTables, item => item.InnerText == "Header");
        Assert.All(savedTables, item => Assert.Equal("OfficeIMOModern", item.GetFirstChild<TableProperties>()!.TableStyle!.Val!.Value));
        Assert.Empty(new DocumentFormat.OpenXml.Validation.OpenXmlValidator().Validate(package));
    }

    [Fact]
    public void Load_KeepsAnAuthoredTableStyle() {
        using WordDocument source = WordDocument.Create();
        source.AddTable(2, 2, WordTableStyle.TableGrid);
        using WordDocument loaded = WordDocument.Load(new MemoryStream(source.ToBytes()));
        Assert.Equal(WordTableStyle.TableGrid, Assert.Single(loaded.Tables).Style);
    }

    [Fact]
    public void Load_KeepsTheDocumentsOwnHeadingStyle() {
        using var stream = new MemoryStream();
        using (WordDocument document = WordDocument.Create()) {
            document.AddParagraph("Heading").Style = WordParagraphStyles.Heading1;
            Style heading = GetStyleDefinition(document, "Heading1");
            heading.StyleRunProperties!.FontSize!.Val = "50";
            document.Save(stream);
        }

        stream.Position = 0;
        using (WordDocument loaded = WordDocument.Load(stream)) {
            loaded.AddParagraph("More");
            using var saved = new MemoryStream();
            loaded.Save(saved);
            saved.Position = 0;
            using WordDocument reloaded = WordDocument.Load(saved);
            Assert.Equal("50", GetStyleDefinition(reloaded, "Heading1").StyleRunProperties!.FontSize!.Val!.Value);
        }
    }
}
