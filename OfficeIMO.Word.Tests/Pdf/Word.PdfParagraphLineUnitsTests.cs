using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests;

public partial class Word {
    [Theory]
    [InlineData(true, WordLineSpacingRule.Auto, 240, false, 36D)]
    [InlineData(true, WordLineSpacingRule.Auto, 240, true, 36D)]
    [InlineData(true, WordLineSpacingRule.Auto, 360, false, 45D)]
    [InlineData(true, WordLineSpacingRule.AtLeast, 600, false, 48D)]
    [InlineData(true, WordLineSpacingRule.Exact, 240, false, 30D)]
    [InlineData(false, WordLineSpacingRule.Auto, 240, false, 28.5D)]
    [InlineData(false, WordLineSpacingRule.Auto, 360, false, 36.75D)]
    [InlineData(false, WordLineSpacingRule.Exact, 240, false, 24D)]
    public void SaveAsPdf_ParagraphSpacingInLinesUsesSectionUnitsRatherThanTheLineAdvance(
        bool grid, WordLineSpacingRule rule, int line, bool after, double expected) {
        using WordDocument source = CreateLineUnitsDocument(grid, rule, line);
        foreach (WordParagraph paragraph in source.Paragraphs) {
            SpacingBetweenLines spacing = paragraph._paragraph.ParagraphProperties!.SpacingBetweenLines!;
            if (after) spacing.AfterLines = 100; else spacing.BeforeLines = 100;
        }
        double[] baselines = ExportGridDocumentBaselines(source);
        Assert.Equal(expected, baselines[0] - baselines[1], 3);
        Assert.Equal(expected, baselines[1] - baselines[2], 3);
    }

    [Theory]
    [InlineData("style", 36D)]
    [InlineData("document", 36D)]
    [InlineData("snap-disabled", 34.5D)]
    public void SaveAsPdf_InheritedLineUnitsTakePrecedenceOverDirectPointSpacing(string route, double expected) {
        using WordDocument source = CreateLineUnitsDocument(true, WordLineSpacingRule.Auto, 240);
        Styles styles = source._wordprocessingDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!;
        if (route == "document") {
            ParagraphPropertiesBaseStyle defaults = styles.DocDefaults!.ParagraphPropertiesDefault!.ParagraphPropertiesBaseStyle!;
            (defaults.SpacingBetweenLines ??= new SpacingBetweenLines()).BeforeLines = 100;
        } else {
            Style normal = styles.Elements<Style>().Single(style => style.StyleId?.Value == "Normal");
            StyleParagraphProperties properties = normal.StyleParagraphProperties ??= new StyleParagraphProperties();
            (properties.SpacingBetweenLines ??= new SpacingBetweenLines()).BeforeLines = 100;
        }
        if (route == "snap-disabled") foreach (WordParagraph paragraph in source.Paragraphs)
            paragraph._paragraph.ParagraphProperties!.Append(new SnapToGrid { Val = false });
        double[] baselines = ExportGridDocumentBaselines(source);
        Assert.Equal(expected, baselines[0] - baselines[1], 3);
        Assert.Equal(expected, baselines[1] - baselines[2], 3);
    }

    [Fact]
    public void SaveAsPdf_TableParagraphLineUnitsUseTheGridPitchWithoutEnablingTableLineHeightAdjustment() {
        using WordDocument source = CreateLineUnitsDocument(true, WordLineSpacingRule.Auto, 240);
        WordParagraph[] paragraphs = source.Paragraphs.ToArray();
        WordTable table = source.AddTable(1, 1);
        TableCell cell = table.Rows[0].Cells[0]._tableCell;
        cell.RemoveAllChildren<Paragraph>();
        foreach (WordParagraph paragraph in paragraphs) {
            paragraph._paragraph.ParagraphProperties!.SpacingBetweenLines!.BeforeLines = 100;
            paragraph._paragraph.Remove();
            cell.Append(paragraph._paragraph);
        }
        double[] baselines = ExportGridDocumentBaselines(source);
        Assert.Equal(34.5D, baselines[0] - baselines[1], 3);
        Assert.Equal(34.5D, baselines[1] - baselines[2], 3);
    }

    [Theory]
    [InlineData("points", 30D)]
    [InlineData("line-units", 18D)]
    [InlineData("explicit-break", 30D)]
    [InlineData("automatic-break", 30D)]
    [InlineData("columns", 18D)]
    [InlineData("list", 18D)]
    public void SaveAsPdf_InitialParagraphSpacingIsPreservedAndSuppressedAtLaterPageStarts(string route, double expected) {
        bool grid = route is "line-units" or "columns" or "list";
        using WordDocument plain = CreateLineUnitsDocument(grid, WordLineSpacingRule.Auto, 240, route);
        using WordDocument spaced = CreateLineUnitsDocument(grid, WordLineSpacingRule.Auto, 240, route);
        if (route == "columns") { plain.Sections[0].ColumnCount = 2; spaced.Sections[0].ColumnCount = 2; }
        if (route == "automatic-break") {
            plain.Sections[0]._sectionProperties.GetFirstChild<PageSize>()!.Height = 5000;
            spaced.Sections[0]._sectionProperties.GetFirstChild<PageSize>()!.Height = 5000;
        }
        WordParagraph[] paragraphs = spaced.Paragraphs.ToArray();
        for (int index = 0; index < paragraphs.Length; index++) {
            if (grid) paragraphs[index]._paragraph.ParagraphProperties!.SpacingBetweenLines!.BeforeLines = 100;
            else paragraphs[index].LineSpacingBefore = global::OfficeIMO.OfficeLength.Points(30D);
            if (route == "explicit-break" && index > 0) {
                paragraphs[index]._paragraph.ParagraphProperties!.Append(new PageBreakBefore());
                plain.Paragraphs[index]._paragraph.ParagraphProperties!.Append(new PageBreakBefore());
            }
        }
        double[] original = ExportGridDocumentBaselines(plain);
        double[] result = ExportGridDocumentBaselines(spaced);
        Assert.Equal(expected, original[0] - result[0], 3);
        if (route == "explicit-break") {
            Assert.Equal(original[1], result[1], 3);
            Assert.Equal(original[2], result[2], 3);
        }
        if (route == "automatic-break") Assert.Equal(original[0], result[2], 3);
    }

    [Fact]
    public void SaveAsPdf_ListItemsRetainIndividualGridSnappingWhenTheirMinimumAdvancesMatch() {
        using WordDocument snapped = CreateLineUnitsDocument(true, WordLineSpacingRule.AtLeast, 600, "list");
        using WordDocument mixed = CreateLineUnitsDocument(true, WordLineSpacingRule.AtLeast, 600, "list");
        mixed.Paragraphs[1]._paragraph.ParagraphProperties!.Append(new SnapToGrid { Val = false });
        double[] original = ExportGridDocumentBaselines(snapped);
        double[] result = ExportGridDocumentBaselines(mixed);
        Assert.Equal(original[0], result[0], 3);
        Assert.Equal(.75D, original[1] - result[1], 3);
        Assert.Equal(original[2], result[2], 3);
    }

    [Theory]
    [InlineData(true, 100, 0, false, 36D)]
    [InlineData(true, 100, 50, false, 36D)]
    [InlineData(true, 50, 100, false, 36D)]
    [InlineData(false, 100, 0, false, 28.5D)]
    [InlineData(true, 100, 50, true, 18D)]
    [InlineData(false, 100, 50, true, 16.5D)]
    public void SaveAsPdf_ListParagraphLineUnitsPreserveCollapsedItemSpacing(
        bool grid, int before, int after, bool contextual, double expected) {
        using WordDocument source = CreateLineUnitsDocument(grid, WordLineSpacingRule.Auto, 240, "list");
        foreach (WordParagraph paragraph in source.Paragraphs) {
            SpacingBetweenLines spacing = paragraph._paragraph.ParagraphProperties!.SpacingBetweenLines!;
            spacing.BeforeLines = before;
            spacing.AfterLines = after;
            paragraph._paragraph.ParagraphProperties.Append(new ContextualSpacing { Val = contextual });
        }
        double[] baselines = ExportGridDocumentBaselines(source);
        Assert.Equal(expected, baselines[0] - baselines[1], 3);
        Assert.Equal(expected, baselines[1] - baselines[2], 3);
    }

    [Theory]
    [InlineData("list", "010", 27D, 27D)]
    [InlineData("list", "101", 27D, 27D)]
    [InlineData("inline", "010", 27D, 27D)]
    [InlineData("inline", "101", 27D, 27D)]
    [InlineData("body", "010", 27D, 27D)]
    [InlineData("body", "101", 27D, 27D)]
    [InlineData("table", "010", 25.5D, 25.5D)]
    [InlineData("table", "101", 25.5D, 25.5D)]
    public void SaveAsPdf_ParagraphLineUnitsHonorEachItemsContextualSpacing(
        string route, string flags, double firstAdvance, double secondAdvance) {
        using WordDocument source = CreateLineUnitsDocument(true, WordLineSpacingRule.Auto, 240, route);
        WordParagraph[] paragraphs = source.Paragraphs.ToArray();
        for (int i = 0; i < paragraphs.Length; i++) {
            ParagraphProperties properties = paragraphs[i]._paragraph.ParagraphProperties!;
            properties.SpacingBetweenLines!.BeforeLines = 100;
            properties.SpacingBetweenLines.AfterLines = 50;
            properties.Append(new ContextualSpacing { Val = flags[i] == '1' });
        }
        if (route == "table") {
            TableCell cell = source.AddTable(1, 1).Rows[0].Cells[0]._tableCell;
            cell.RemoveAllChildren<Paragraph>();
            foreach (WordParagraph paragraph in paragraphs) {
                paragraph._paragraph.Remove();
                cell.Append(paragraph._paragraph);
            }
        }
        double[] baselines = ExportGridDocumentBaselines(source);
        Assert.Equal(firstAdvance, baselines[0] - baselines[1], 3);
        Assert.Equal(secondAdvance, baselines[1] - baselines[2], 3);
    }

    private static WordDocument CreateLineUnitsDocument(bool grid, WordLineSpacingRule rule, int line, string route = "body") {
        WordDocument source = WordDocument.Create();
        if (grid) source.Sections[0]._sectionProperties.Append(new DocGrid { Type = DocGridValues.Lines, LinePitch = 360 });
        WordList? list = route is "list" or "inline" ? source.AddCustomList() : null;
        list?.Numbering.AddLevel(new WordListLevel(WordListLevelKind.DecimalDot) {
            LevelSuffix = route == "inline" ? WordListLevelSuffix.Space : WordListLevelSuffix.Tab
        });
        foreach (char text in "ABC") {
            WordParagraph paragraph = list?.AddItem(text.ToString()) ?? source.AddParagraph(text.ToString());
            paragraph.FontSize = 11;
            paragraph.FontFamily = "Arial";
            paragraph.LineSpacing = line; paragraph.LineSpacingRule = rule;
            paragraph.LineSpacingBefore = global::OfficeIMO.OfficeLength.Zero; paragraph.LineSpacingAfter = global::OfficeIMO.OfficeLength.Zero;
        }
        return source;
    }
}
