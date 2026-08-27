using OfficeIMO.Drawing;
using OfficeIMO.Excel;
using OfficeIMO.Markup;
using OfficeIMO.Markup.Excel;
using OfficeIMO.Markup.PowerPoint;
using OfficeIMO.Markup.Word;
using OfficeIMO.PowerPoint;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests.Markup;

public class OfficeMarkupTextFormattingTests {
    [Fact]
    public void StyleResolver_ParsesCompleteTypographyAttributes() {
        var block = new OfficeMarkupTextBoxBlock("Mixed case");
        block.Attributes["font-family"] = "Aptos";
        block.Attributes["font-size"] = "18pt";
        block.Attributes["bold"] = "true";
        block.Attributes["italic"] = "true";
        block.Attributes["underline"] = "wavy";
        block.Attributes["strikethrough"] = "double";
        block.Attributes["baseline"] = "superscript";
        block.Attributes["text-case"] = "toggle";
        block.Attributes["small-caps"] = "true";
        block.Attributes["color"] = "336699";
        block.Attributes["highlight"] = "#FFF2CC";

        var document = new OfficeMarkupDocument(OfficeMarkupProfile.Presentation);
        OfficeMarkupResolvedStyle style = Assert.IsType<OfficeMarkupResolvedStyle>(
            OfficeMarkupStyleResolver.Create(document).Resolve(block));

        Assert.Equal("Aptos", style.FontName);
        Assert.Equal(18, style.FontSize);
        Assert.True(style.Bold);
        Assert.True(style.Italic);
        Assert.Equal(OfficeTextDecorationStyle.Wavy, style.UnderlineStyle);
        Assert.Equal(OfficeTextDecorationStyle.Double, style.StrikethroughStyle);
        Assert.Equal(OfficeTextBaseline.Superscript, style.Baseline);
        Assert.Equal(OfficeTextCase.ToggleCase, style.TextCase);
        Assert.True(style.SmallCaps);
        Assert.Equal("#336699", style.TextColor);
        Assert.Equal("#FFF2CC", style.HighlightColor);
    }

    [Fact]
    public void WordExporter_AppliesRootAndNestedListTypographyInsideAndOutsideSections() {
        static OfficeMarkupListBlock CreateList() {
            var root = new OfficeMarkupListBlock(ordered: false);
            root.Attributes["font-family"] = "Consolas";
            root.Attributes["underline"] = "double";
            var rootItem = new OfficeMarkupListItem("Root item");
            var nested = new OfficeMarkupListBlock(ordered: false);
            nested.Attributes["italic"] = "true";
            nested.Attributes["color"] = "336699";
            nested.Items.Add(new OfficeMarkupListItem("Nested item"));
            rootItem.Blocks.Add(nested);
            root.Items.Add(rootItem);
            return root;
        }

        var source = new OfficeMarkupDocument(OfficeMarkupProfile.Document);
        source.Blocks.Add(CreateList());
        var section = new OfficeMarkupSectionBlock("Section");
        section.Blocks.Add(CreateList());
        source.Blocks.Add(section);

        using WordDocument word = source.ToWordDocumentResult().RequireValue();
        WordParagraphSnapshot[] paragraphs = word.CreateInspectionSnapshot().Sections
            .SelectMany(item => item.Elements).OfType<WordParagraphSnapshot>()
            .Where(item => item.Text.Contains("item", StringComparison.Ordinal))
            .ToArray();

        Assert.Equal(4, paragraphs.Length);
        Assert.All(paragraphs.Where(item => item.Text.Contains("Root", StringComparison.Ordinal)), paragraph => {
            WordRunSnapshot run = Assert.Single(paragraph.Runs);
            Assert.Equal("Consolas", run.FontFamily);
            Assert.Equal(WordUnderlineStyle.Double, run.UnderlineStyle);
        });
        Assert.All(paragraphs.Where(item => item.Text.Contains("Nested", StringComparison.Ordinal)), paragraph => {
            WordRunSnapshot run = Assert.Single(paragraph.Runs);
            Assert.True(run.Italic);
            Assert.Equal("336699", run.ColorHex);
        });
    }

    [Fact]
    public void WordExporter_AppliesBlockTypographyAndCase() {
        var paragraph = new OfficeMarkupParagraphBlock("Mixed case");
        AddCommonStyle(paragraph.Attributes, "double", "superscript", "uppercase");
        var source = new OfficeMarkupDocument(OfficeMarkupProfile.Document);
        source.Blocks.Add(paragraph);

        using WordDocument word = source.ToWordDocumentResult().RequireValue();
        WordParagraphSnapshot paragraphSnapshot = Assert.IsType<WordParagraphSnapshot>(
            Assert.Single(Assert.Single(word.CreateInspectionSnapshot().Sections).Elements));
        WordRunSnapshot run = Assert.Single(paragraphSnapshot.Runs);

        Assert.Equal("MIXED CASE", run.Text);
        Assert.True(run.Bold);
        Assert.True(run.Italic);
        Assert.Equal(WordUnderlineStyle.Double, run.UnderlineStyle);
        Assert.True(run.DoubleStrike);
        Assert.Equal("Superscript", run.VerticalTextAlignment);
        Assert.Equal("SmallCaps", run.CapsStyle);
        Assert.Equal("336699", run.ColorHex);
        Assert.Equal("FFF2CC", run.RunShadingFillColorHex);
    }

    [Fact]
    public void ExcelExporter_AppliesNativeAccountingUnderlineScriptsAndCase() {
        var source = new OfficeMarkupDocument(OfficeMarkupProfile.Workbook);
        source.Blocks.Add(new OfficeMarkupSheetBlock("Data"));
        var range = new OfficeMarkupRangeBlock("A1") { Sheet = "Data" };
        range.Values.Add(new[] { "Mixed case" });
        source.Blocks.Add(range);
        var formatting = new OfficeMarkupFormattingBlock("A1");
        formatting.Attributes["sheet"] = "Data";
        formatting.Attributes["font"] = "Aptos";
        formatting.Attributes["font-size"] = "17pt";
        formatting.Attributes["bold"] = "true";
        formatting.Attributes["italic"] = "true";
        formatting.Attributes["underline"] = "double-accounting";
        formatting.Attributes["strikethrough"] = "true";
        formatting.Attributes["baseline"] = "subscript";
        formatting.Attributes["text-case"] = "toggle";
        formatting.Attributes["color"] = "336699";
        formatting.Attributes["highlight"] = "FFF2CC";
        source.Blocks.Add(formatting);

        using ExcelDocument workbook = source.ToExcelDocumentResult().RequireValue();
        ExcelSheet sheet = Assert.Single(workbook.Sheets);
        ExcelCellStyleSnapshot style = sheet.GetCellStyle(1, 1);

        Assert.Equal("mIXED CASE", sheet.CellAt(1, 1).GetValue().Value?.ToString());
        Assert.Equal("Aptos", style.FontName);
        Assert.Equal(17D, style.FontSize);
        Assert.True(style.Bold);
        Assert.True(style.Italic);
        Assert.Equal(ExcelUnderlineStyle.DoubleAccounting, style.UnderlineStyle);
        Assert.True(style.Strikethrough);
        Assert.Equal(ExcelVerticalTextAlignment.Subscript, style.VerticalTextAlignment);
        Assert.Equal("336699", style.FontColorHex);
        Assert.Equal("FFF2CC", style.FillColorHex);
    }

    [Fact]
    public void PowerPointExporter_AppliesRunTypographyAndCase() {
        var textBox = new OfficeMarkupTextBoxBlock("Mixed case");
        AddCommonStyle(textBox.Attributes, "wavy", "subscript", "lowercase");
        var slide = new OfficeMarkupSlideBlock();
        slide.Blocks.Add(textBox);
        var source = new OfficeMarkupDocument(OfficeMarkupProfile.Presentation);
        source.Blocks.Add(slide);

        using PowerPointPresentation presentation = source.ToPowerPointPresentationResult().RequireValue();
        PowerPointTextBox output = Assert.Single(Assert.Single(presentation.Slides).Shapes.OfType<PowerPointTextBox>(),
            box => box.Text.Contains("mixed case", StringComparison.Ordinal));
        PowerPointTextRun run = Assert.Single(Assert.Single(output.Paragraphs).Runs);

        Assert.Equal("mixed case", run.Text);
        Assert.True(run.Bold);
        Assert.True(run.Italic);
        Assert.Equal(PowerPointUnderlineStyle.Wavy, run.UnderlineStyle);
        Assert.Equal(PowerPointStrikeStyle.Double, run.StrikeStyle);
        Assert.True(run.BaselinePercent < 0D);
        Assert.Equal(PowerPointCapitalization.SmallCaps, run.Capitalization);
        Assert.Equal("336699", run.Color);
        Assert.Equal("FFF2CC", run.HighlightColor);
    }

    private static void AddCommonStyle(IDictionary<string, string> attributes, string underline, string baseline, string textCase) {
        attributes["font"] = "Aptos";
        attributes["font-size"] = "18";
        attributes["bold"] = "true";
        attributes["italic"] = "true";
        attributes["underline"] = underline;
        attributes["strikethrough"] = "double";
        attributes["baseline"] = baseline;
        attributes["text-case"] = textCase;
        attributes["small-caps"] = "true";
        attributes["color"] = "336699";
        attributes["highlight"] = "FFF2CC";
    }
}
