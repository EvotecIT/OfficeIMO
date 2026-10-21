using DocumentFormat.OpenXml;
using OfficeIMO.Word;
using OfficeIMO.Word.Fluent;
using W = DocumentFormat.OpenXml.Wordprocessing;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class WordLogicalTableMutationAtomicityTests {
    [Fact]
    public void RectangularMergeRejectsExcessiveCellDepthBeforeChangingEarlierRows() {
        using WordDocument document = WordDocument.Create();
        WordTable table = document.AddTable(2, 2);
        table.Rows[0].Cells[0].Paragraphs[0].Text = "ANCHOR";
        table.Rows[0].Cells[1].Paragraphs[0].Text = "NEIGHBOR";
        OpenXmlElement target = table.Rows[1].Cells[1]._tableCell;
        WrapBeyondLimit(target, row: false);
        string before = table._table.OuterXml;

        Assert.Throws<InvalidDataException>(() => table.MergeCells(0, 0, 2, 2, copyParagraphs: true));

        Assert.Equal(before, table._table.OuterXml);
    }

    [Fact]
    public void FluentInsertionRejectsExcessiveRowDepthBeforeAppending() {
        using WordDocument document = WordDocument.Create();
        WordTable? table = null;
        WordTableBuilder? builder = null;
        document.AsFluent().Table(value => {
            builder = value;
            value.Create(2, 1);
            table = value.Table;
        });
        WrapBeyondLimit(table!.Rows[1]._tableRow, row: true);
        string before = table._table.OuterXml;

        Assert.Throws<InvalidDataException>(() => builder!.InsertRow(1, "NEW"));

        Assert.Equal(before, table._table.OuterXml);
    }

    [Theory]
    [InlineData("horizontal", false)]
    [InlineData("horizontal", true)]
    [InlineData("vertical-cell", false)]
    [InlineData("vertical-cell", true)]
    [InlineData("vertical-row", false)]
    [InlineData("vertical-row", true)]
    public void ExcessiveLogicalContainerDepthRejectsBeforeChangingMergeStateOrContent(string direction, bool split) {
        using WordDocument document = WordDocument.Create();
        WordTable table = document.AddTable(3, 3);
        WordTableCell anchor = table.Rows[0].Cells[0];
        WordTableCell first = direction == "horizontal" ? table.Rows[0].Cells[1] : table.Rows[1].Cells[0];
        WordTableCell last = direction == "horizontal" ? table.Rows[0].Cells[2] : table.Rows[2].Cells[0];
        anchor.Paragraphs[0].Text = "ANCHOR";
        first.Paragraphs[0].Text = "FIRST";
        last.Paragraphs[0].Text = "LAST";
        if (split) {
            if (direction == "horizontal") {
                anchor.HorizontalMerge = WordCellMerge.Restart;
                first.HorizontalMerge = last.HorizontalMerge = WordCellMerge.Continue;
            } else {
                anchor.VerticalMerge = WordCellMerge.Restart;
                first.VerticalMerge = last.VerticalMerge = WordCellMerge.Continue;
            }
        }
        OpenXmlElement target = direction == "vertical-row" ? table.Rows[2]._tableRow : last._tableCell;
        WrapBeyondLimit(target, direction == "vertical-row");
        string before = table._table.OuterXml;

        Assert.Throws<InvalidDataException>(() => {
            if (direction == "horizontal") {
                if (split) anchor.SplitHorizontally(2);
                else anchor.MergeHorizontally(2, copyParagraphs: true);
            } else {
                if (split) anchor.SplitVertically(2);
                else anchor.MergeVertically(2, copyParagraphs: true);
            }
        });

        Assert.Equal(before, table._table.OuterXml);
    }

    private static void WrapBeyondLimit(OpenXmlElement target, bool row) {
        for (int depth = 0; depth < 129; depth++) {
            OpenXmlElement wrapper = row ? new W.CustomXmlRow() : new W.CustomXmlCell();
            target.Parent!.ReplaceChild(wrapper, target);
            wrapper.Append(target);
            target = wrapper;
        }
    }
}
