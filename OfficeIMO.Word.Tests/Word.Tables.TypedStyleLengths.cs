using System;
using System.IO;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests;

public partial class Word {
    [Fact]
    public void TableStyleLengths_SaveRoundedValuesAndPreserveExplicitZeroAndInheritance() {
        using WordDocument document = WordDocument.Create();
        WordTable table = document.AddTable(1, 1, WordTableStyle.TableNormal);
        WordTableStyleDetails details = table.StyleDetails!;
        details.MarginDefaultTop = global::OfficeIMO.OfficeLength.Emus(318); // Just above half a twip.
        details.MarginDefaultBottom = global::OfficeIMO.OfficeLength.Zero;
        details.MarginDefaultLeft = global::OfficeIMO.OfficeLength.Points(5);
        details.MarginDefaultRight = global::OfficeIMO.OfficeLength.Points(7);
        details.MarginDefaultRight = null;
        details.CellSpacing = global::OfficeIMO.OfficeLength.Zero;
        details.TableIndentation = global::OfficeIMO.OfficeLength.Points(-2000);
        details.SetCustomBorders(topStyle: WordBorderStyle.Single,
            topSize: global::OfficeIMO.OfficeLength.Points(0.2), topColor: global::OfficeIMO.OfficeColor.Blue);
        details.SetCustomBorders(topColor: global::OfficeIMO.OfficeColor.Red);

        using WordDocument loaded = WordDocument.Load(new MemoryStream(document.ToBytes()));
        WordTableStyleDetails actual = loaded.Tables[0].StyleDetails!;
        Assert.Equal(global::OfficeIMO.OfficeLength.Points(0.05), actual.MarginDefaultTop);
        Assert.Equal(global::OfficeIMO.OfficeLength.Zero, actual.MarginDefaultBottom);
        Assert.Equal(global::OfficeIMO.OfficeLength.Points(5), actual.MarginDefaultLeft);
        Assert.Null(actual.MarginDefaultRight);
        Assert.Equal(global::OfficeIMO.OfficeLength.Zero, actual.CellSpacing);
        Assert.Equal(global::OfficeIMO.OfficeLength.Points(-2000), actual.TableIndentation);
        var border = actual.GetBorderProperties(WordTableBorderSide.Top);
        Assert.Equal(global::OfficeIMO.OfficeLength.Points(0.25), border.Size);
        Assert.Equal("FF0000", border.ColorHex);
        TableProperties native = loaded.Tables[0]._tableProperties!;
        Assert.Equal("1", native.TableCellMarginDefault!.TopMargin!.Width!.Value);
        Assert.Equal("0", native.TableCellMarginDefault.BottomMargin!.Width!.Value);
        Assert.Null(native.TableCellMarginDefault.TableCellRightMargin);
        Assert.Equal("0", native.TableCellSpacing!.Width!.Value);
        Assert.Equal(-40000, native.TableIndentation!.Width!.Value);
        Assert.Equal(2U, native.TableBorders!.TopBorder!.Size!.Value);

        actual.CellSpacing = null;
        actual.TableIndentation = null;
        using WordDocument cleared = WordDocument.Load(new MemoryStream(loaded.ToBytes()));
        Assert.Null(cleared.Tables[0].StyleDetails!.CellSpacing);
        Assert.Null(cleared.Tables[0].StyleDetails!.TableIndentation);
    }

    [Fact]
    public void TableStyleLengths_RejectInvalidGeometryBeforeChangingExistingBordersOrMargins() {
        using WordDocument document = WordDocument.Create();
        WordTable table = document.AddTable(1, 1, WordTableStyle.TableGrid);
        WordTableStyleDetails details = table.StyleDetails!;
        details.SetBordersForAllSides(WordBorderStyle.Single, global::OfficeIMO.OfficeLength.Points(1), global::OfficeIMO.OfficeColor.Blue);
        string before = table._table.OuterXml;
        var negative = global::OfficeIMO.OfficeLength.Emus(-1);
        var oversizedMargin = global::OfficeIMO.OfficeLength.Points(32768 / 20D);
        var oversizedIndent = global::OfficeIMO.OfficeLength.Emus(((long)int.MaxValue + 1) * 635);
        var oversizedBorder = global::OfficeIMO.OfficeLength.Points(((double)uint.MaxValue + 1) / 8D);
        Assert.Throws<ArgumentOutOfRangeException>(() => details.MarginDefaultTop = negative);
        Assert.Throws<ArgumentOutOfRangeException>(() => details.MarginDefaultBottom = oversizedMargin);
        Assert.Throws<ArgumentOutOfRangeException>(() => details.MarginDefaultLeft = negative);
        Assert.Throws<ArgumentOutOfRangeException>(() => details.MarginDefaultRight = oversizedMargin);
        Assert.Throws<ArgumentOutOfRangeException>(() => details.CellSpacing = negative);
        Assert.Throws<ArgumentOutOfRangeException>(() => details.TableIndentation = oversizedIndent);
        Assert.Throws<ArgumentOutOfRangeException>(() => details.SetBordersForAllSides(WordBorderStyle.Single, negative, global::OfficeIMO.OfficeColor.Red));
        Assert.Throws<ArgumentOutOfRangeException>(() => details.SetBordersOutsideInside(
            WordBorderStyle.Single, global::OfficeIMO.OfficeLength.Points(2), global::OfficeIMO.OfficeColor.Red,
            WordBorderStyle.Single, oversizedBorder, global::OfficeIMO.OfficeColor.Blue));
        Assert.Throws<ArgumentOutOfRangeException>(() => details.SetCustomBorders(
            topSize: global::OfficeIMO.OfficeLength.Points(2), insideVSize: negative));
        Assert.Equal(before, table._table.OuterXml);
    }

    [Fact]
    public void PaperStyleLengths_SaveDefaultsAndExactLineHeightAndRejectInvalidConfiguration() {
        var paper = new WordPaperStyle();
        Assert.Equal(global::OfficeIMO.OfficeLength.Points(0.5), paper.RuleSize);
        Assert.Equal(global::OfficeIMO.OfficeLength.Points(4.5), paper.TableCellVerticalPadding);
        Assert.Equal(global::OfficeIMO.OfficeLength.Points(5), paper.TableCellHorizontalPadding);
        Assert.Equal(global::OfficeIMO.OfficeLength.Points(2), paper.TableParagraphSpacing);
        Assert.Null(paper.TableLineSpacing);
        var negative = global::OfficeIMO.OfficeLength.Emus(-1);
        Assert.Throws<ArgumentOutOfRangeException>(() => paper.RuleSize = negative);
        Assert.Throws<ArgumentOutOfRangeException>(() => paper.TableCellVerticalPadding = negative);
        Assert.Throws<ArgumentOutOfRangeException>(() => paper.TableCellHorizontalPadding = global::OfficeIMO.OfficeLength.Points(2000));
        Assert.Throws<ArgumentOutOfRangeException>(() => paper.TableParagraphSpacing = negative);
        Assert.Throws<ArgumentOutOfRangeException>(() => paper.TableLineSpacing = negative);
        paper.TableLineSpacing = global::OfficeIMO.OfficeLength.Points(11);
        using WordDocument document = WordDocument.Create();
        WordTable table = document.AddTable(1, 1, WordTableStyle.TableNormal);
        paper.ApplyTable(table, hasHeader: false);
        using WordDocument loaded = WordDocument.Load(new MemoryStream(document.ToBytes()));
        WordTable actual = loaded.Tables[0];
        Assert.Equal(global::OfficeIMO.OfficeLength.Points(4.5), actual.StyleDetails!.MarginDefaultTop);
        Assert.Equal(global::OfficeIMO.OfficeLength.Points(5), actual.StyleDetails.MarginDefaultLeft);
        Assert.Equal(global::OfficeIMO.OfficeLength.Points(0.5), actual.StyleDetails.GetBorderProperties(WordTableBorderSide.InsideHorizontal).Size);
        WordParagraph paragraph = actual.Rows[0].Cells[0].Paragraphs[0];
        Assert.Equal(global::OfficeIMO.OfficeLength.Points(2), paragraph.LineSpacingBefore);
        Assert.Equal(global::OfficeIMO.OfficeLength.Points(2), paragraph.LineSpacingAfter);
        Assert.Equal(WordLineSpacing.Exact(global::OfficeIMO.OfficeLength.Points(11)), paragraph.LineSpacing);
    }
}
