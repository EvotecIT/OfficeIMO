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
    public void Create_HeadingStylesFollowTheDefaultThemeTypeScale(int level) {
        using WordDocument document = WordDocument.Create();
        OfficeTheme theme = OfficeTheme.Default;
        Style heading = GetStyleDefinition(document, "Heading" + level);
        OfficeLength expected = level == 1 ? theme.TypeScale.Heading1 : level == 2 ? theme.TypeScale.Heading2 : theme.TypeScale.Heading3;

        Assert.Equal(HalfPoints(expected), heading.StyleRunProperties!.FontSize!.Val!.Value);
        Assert.NotNull(heading.StyleRunProperties.Bold);
        Assert.Equal(theme.Palette.Text.ToRgbHex(), heading.StyleRunProperties.Color!.Val!.Value);
        Assert.NotNull(heading.StyleParagraphProperties!.KeepNext);
        Assert.NotEqual("0", heading.StyleParagraphProperties.SpacingBetweenLines!.After!.Value);
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
        WordTable table = document.AddTable(3, 3, WordTableStyle.OfficeIMOModern);
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
