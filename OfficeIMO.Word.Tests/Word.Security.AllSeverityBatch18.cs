using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests;

public class WordAllSeverityBatch18SecurityTests {
    [Fact]
    public void TextBoxAlignmentAcceptsInsideAndDefaultsMalformedValues() {
        Assert.Equal(WordHorizontalAlignmentValues.Inside, HorizontalAlignmentHelper.FromString(" inside "));
        Assert.Equal("inside", HorizontalAlignmentHelper.ToString(WordHorizontalAlignmentValues.Inside));
        Assert.Equal(WordHorizontalAlignmentValues.Center, HorizontalAlignmentHelper.FromString(null));
        Assert.Equal(WordHorizontalAlignmentValues.Center, HorizontalAlignmentHelper.FromString(string.Empty));
        Assert.Equal(WordHorizontalAlignmentValues.Center, HorizontalAlignmentHelper.FromString("future-value"));
        Assert.Equal(3, (int)WordHorizontalAlignmentValues.Outside);
        Assert.Equal(4, (int)WordHorizontalAlignmentValues.Inside);
    }

    [Fact]
    public void ParagraphFormattingCreatesMissingRunAndParagraphProperties() {
        using WordDocument document = WordDocument.Create();
        var paragraph = new WordParagraph(document, newParagraph: true, newRun: false);

        paragraph.VerticalTextAlignment = WordVerticalTextPosition.Superscript;
        paragraph.Borders.LeftStyle = WordBorderStyle.Single;
        WordParagraph emptyText = paragraph.AddText(null);

        Assert.NotNull(paragraph._run);
        Assert.Equal(VerticalPositionValues.Superscript,
            paragraph._run!.RunProperties?.VerticalTextAlignment?.Val?.Value);
        Assert.Equal(BorderValues.Single,
            paragraph._paragraph.ParagraphProperties?.ParagraphBorders?.LeftBorder?.Val?.Value);
        Assert.Equal(string.Empty, emptyText.Text);
    }

    [Fact]
    public void VerticalTextAlignmentFormatsTheHyperlinkRunInPlace() {
        using WordDocument document = WordDocument.Create();
        WordParagraph paragraph = document.AddParagraph("Before ");
        paragraph.AddHyperLink("linked", new Uri("https://example.test"));
        WordParagraph hyperlinkParagraph = Assert.Single(
            document.Paragraphs,
            candidate => candidate.IsHyperLink);
        int directRunCount = paragraph._paragraph.Elements<Run>().Count();

        hyperlinkParagraph.VerticalTextAlignment = WordVerticalTextPosition.Superscript;

        Assert.Equal(WordVerticalTextPosition.Superscript, hyperlinkParagraph.VerticalTextAlignment);
        Assert.Equal(
            VerticalPositionValues.Superscript,
            hyperlinkParagraph.Hyperlink!._runProperties.VerticalTextAlignment?.Val?.Value);
        Assert.Equal(directRunCount, paragraph._paragraph.Elements<Run>().Count());

        hyperlinkParagraph.VerticalTextAlignment = null;

        Assert.Null(hyperlinkParagraph.VerticalTextAlignment);
        Assert.Null(hyperlinkParagraph.Hyperlink!._runProperties.VerticalTextAlignment);
    }

    [Fact]
    public void EmptyTableHeaderAndRowApisRemainUsable() {
        using WordDocument document = WordDocument.Create();
        WordTable table = document.AddTable(0, 3);

        Assert.False(table.RepeatAsHeaderRowAtTheTopOfEachPage);
        Assert.False(table.RepeatHeaderRowAtTheTopOfEachPage);
        table.RepeatAsHeaderRowAtTheTopOfEachPage = true;
        table.RepeatHeaderRowAtTheTopOfEachPage = true;
        Exception? commentException = Record.Exception(
            () => table.AddComment("Reviewer", "R", "No range exists"));
        WordTableRow row = table.AddRow();

        Assert.Null(commentException);
        Assert.Equal(3, row.CellsCount);
    }

    [Fact]
    public void TableCommentRepairsCellsWithoutParagraphs() {
        using WordDocument document = WordDocument.Create();
        WordTable table = document.AddTable(1, 1);
        table.Rows[0].Cells[0]._tableCell.RemoveAllChildren<Paragraph>();

        Exception? exception = Record.Exception(
            () => table.AddComment("Reviewer", "R", "Comment"));

        Assert.Null(exception);
        Paragraph paragraph = Assert.Single(
            table.Rows[0].Cells[0]._tableCell.Elements<Paragraph>());
        Assert.NotNull(paragraph.GetFirstChild<CommentRangeStart>());
        Assert.NotNull(paragraph.GetFirstChild<CommentRangeEnd>());
        Assert.NotNull(paragraph.Descendants<CommentReference>().SingleOrDefault());
    }

    [Fact]
    public void TableCommentKeepsRangeOutsideNestedTables() {
        using WordDocument document = WordDocument.Create();
        WordTable table = document.AddTable(1, 2);
        WordTable nested = table.Rows[0].Cells[1].AddTable(1, 1);

        table.AddComment("Reviewer", "R", "Comment");

        TableCell outerLastCell = table.Rows[0].Cells[1]._tableCell;
        Paragraph outerLastParagraph = outerLastCell.Elements<Paragraph>().Last();
        Assert.NotNull(outerLastParagraph.GetFirstChild<CommentRangeEnd>());
        Assert.NotNull(outerLastParagraph.Descendants<CommentReference>().SingleOrDefault());
        Assert.Empty(nested._table.Descendants<CommentRangeEnd>());
        Assert.Empty(nested._table.Descendants<CommentReference>());
    }

    [Fact]
    public void CleanupDoesNotMergeRunsContainingNonTextContent() {
        using WordDocument document = WordDocument.Create();
        WordParagraph paragraph = document.AddParagraph();
        paragraph._paragraph.RemoveAllChildren<Run>();
        paragraph._paragraph.Append(
            new Run(new Text("first")),
            new Run(new Text("second"), new Break(), new Text("third")));

        document.CleanupDocument(DocumentCleanupOptions.MergeIdenticalRuns);

        Run[] runs = paragraph._paragraph.Elements<Run>().ToArray();
        Assert.Equal(2, runs.Length);
        Assert.NotNull(runs[1].GetFirstChild<Break>());
        Assert.Equal(new[] { "second", "third" },
            runs[1].Elements<Text>().Select(static text => text.Text).ToArray());
    }

    [Fact]
    public void VerticalMergeUsesCellIndexWhenRowsHaveProperties() {
        using WordDocument document = WordDocument.Create();
        WordTable table = document.AddTable(2, 2);
        table.Rows[0]._tableRow.TableRowProperties = new TableRowProperties();
        table.Rows[1]._tableRow.TableRowProperties = new TableRowProperties();

        table.Rows[0].Cells[0].MergeVertically(1);

        Assert.Equal(MergedCellValues.Restart,
            table.Rows[0].Cells[0]._tableCell.TableCellProperties?.VerticalMerge?.Val?.Value);
        Assert.Equal(MergedCellValues.Continue,
            table.Rows[1].Cells[0]._tableCell.TableCellProperties?.VerticalMerge?.Val?.Value);
        Assert.Null(table.Rows[1].Cells[1]._tableCell.TableCellProperties?.VerticalMerge);
    }
}
