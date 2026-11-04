using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using OfficeIMO.Word;
using W = DocumentFormat.OpenXml.Wordprocessing;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class WordLogicalTableMutationTests {
    [Theory]
    [InlineData("sdt")]
    [InlineData("customXml")]
    public void SetCellTextUsesTheLogicalIndexWithoutChangingItsNeighbor(string container) {
        using WordDocument document = WordDocument.Create();
        WordTable table = CreateTable(document, 1, 3);
        WrapCell(table.Rows[0].Cells[1]._tableCell, container, "updated-cell");

        table.SetCellText(0, 1, "UPDATED", bold: true);

        Assert.Equal(new[] { "R0C0", "UPDATED", "R0C2" }, table.Rows[0].Cells.Select(CellText));
        Assert.True(table.Rows[0].Cells[1].Paragraphs[0].Bold);
        Assert.Equal("updated-cell", ContainerTag(table.Rows[0].Cells[1]._tableCell.Parent!));
    }

    [Theory]
    [InlineData("sdt", false)]
    [InlineData("sdt", true)]
    [InlineData("customXml", false)]
    [InlineData("customXml", true)]
    public void HorizontalMergeTargetsTheNextLogicalCellAndPreservesTheFollowingCell(string container, bool copyParagraphs) {
        using WordDocument document = WordDocument.Create();
        WordTable table = CreateTable(document, 1, 3);
        WrapCell(table.Rows[0].Cells[1]._tableCell, container, "middle-cell");
        WordTableRow row = table.Rows[0];

        row.Cells[0].MergeHorizontally(1, copyParagraphs);

        Assert.Equal("R0C2", CellText(row.Cells[2]));
        Assert.Null(row.Cells[2].HorizontalMerge);
        Assert.Equal(WordCellMerge.Continue, row.Cells[1].HorizontalMerge);
        Assert.Equal(string.Empty, CellText(row.Cells[1]));
        Assert.Equal(copyParagraphs ? new[] { "R0C0", "R0C1" } : new[] { "R0C0" },
            row.Cells[0].Paragraphs.Select(paragraph => paragraph.Text));
        Assert.Equal("middle-cell", ContainerTag(row.Cells[1]._tableCell.Parent!));
    }

    [Theory]
    [InlineData("sdt")]
    [InlineData("customXml")]
    public void HorizontalSplitHonorsTheAdditionalLogicalCellCount(string container) {
        using WordDocument document = WordDocument.Create();
        WordTable table = CreateTable(document, 1, 3);
        table.Rows[0].Cells[0].HorizontalMerge = WordCellMerge.Restart;
        table.Rows[0].Cells[1].HorizontalMerge = WordCellMerge.Continue;
        table.Rows[0].Cells[2].HorizontalMerge = WordCellMerge.Continue;
        WrapCell(table.Rows[0].Cells[1]._tableCell, container, "middle-cell");

        table.Rows[0].Cells[0].SplitHorizontally(1);

        Assert.Null(table.Rows[0].Cells[0].HorizontalMerge);
        Assert.Null(table.Rows[0].Cells[1].HorizontalMerge);
        Assert.Equal(WordCellMerge.Continue, table.Rows[0].Cells[2].HorizontalMerge);
        Assert.Equal(new[] { "R0C0", "R0C1", "R0C2" }, table.Rows[0].Cells.Select(CellText));
    }

    [Theory]
    [InlineData("cell", false)]
    [InlineData("cell", true)]
    [InlineData("row", false)]
    [InlineData("row", true)]
    [InlineData("anchor", false)]
    [InlineData("anchor", true)]
    public void VerticalMergeUsesLogicalRowsAndTheSameLogicalCell(string shape, bool copyParagraphs) {
        using WordDocument document = WordDocument.Create();
        WordTable table = CreateVerticalFixture(document, shape);

        table.Rows[0].Cells[0].MergeVertically(1, copyParagraphs);

        Assert.Equal(WordCellMerge.Continue, table.Rows[1].Cells[0].VerticalMerge);
        Assert.Equal(string.Empty, CellText(table.Rows[1].Cells[0]));
        Assert.Equal(copyParagraphs ? new[] { "R0C0", "R1C0" } : new[] { "R0C0" },
            table.Rows[0].Cells[0].Paragraphs.Select(paragraph => paragraph.Text));
        Assert.Equal(new[] { "R0C1", "R1C1", "R2C1" }, table.Rows.Select(row => CellText(row.Cells[1])));
        Assert.All(table.Rows, row => Assert.Null(row.Cells[1].VerticalMerge));
        Assert.Equal("R2C0", CellText(table.Rows[2].Cells[0]));
        Assert.Null(table.Rows[2].Cells[0].VerticalMerge);
    }

    [Theory]
    [InlineData("cell")]
    [InlineData("row")]
    [InlineData("anchor")]
    public void VerticalSplitHonorsTheAdditionalLogicalRowCount(string shape) {
        using WordDocument document = WordDocument.Create();
        WordTable table = CreateVerticalFixture(document, shape);
        table.Rows[0].Cells[0].VerticalMerge = WordCellMerge.Restart;
        table.Rows[1].Cells[0].VerticalMerge = WordCellMerge.Continue;
        table.Rows[2].Cells[0].VerticalMerge = WordCellMerge.Continue;

        table.Rows[0].Cells[0].SplitVertically(1);

        Assert.Null(table.Rows[0].Cells[0].VerticalMerge);
        Assert.Null(table.Rows[1].Cells[0].VerticalMerge);
        Assert.Equal(WordCellMerge.Continue, table.Rows[2].Cells[0].VerticalMerge);
        Assert.Equal(new[] { "R0C0", "R1C0", "R2C0" }, table.Rows.Select(row => CellText(row.Cells[0])));
    }

    [Theory]
    [InlineData("sdt", false)]
    [InlineData("customXml", false)]
    [InlineData("sdt", true)]
    [InlineData("customXml", true)]
    public void SaveNormalizesWrappedMergeContinuationsAndPreservesOtherContainersAndParts(string container, bool omittedContinuationValue) {
        using WordDocument document = WordDocument.Create();
        WordTable table = CreateTable(document, 1, 3);
        table.Rows[0].Cells[0].HorizontalMerge = WordCellMerge.Restart;
        table.Rows[0].Cells[1].HorizontalMerge = WordCellMerge.Continue;
        if (omittedContinuationValue) table.Rows[0].Cells[1]._tableCell.TableCellProperties!.HorizontalMerge!.Val = null;
        WrapCell(table.Rows[0].Cells[1]._tableCell, container, "merged-out-cell");
        WrapCell(table.Rows[0].Cells[2]._tableCell, "sdt", "retained-cell");
        WrapRow(table.Rows[0]._tableRow, "customXml", "retained-row");
        const string customXml = "<record xmlns='urn:officeimo:logical-table'>retained-native-part</record>";
        CustomXmlPart part = document.MainDocumentPartRoot.AddCustomXmlPart(CustomXmlPartType.CustomXml);
        using (var data = new MemoryStream(System.Text.Encoding.UTF8.GetBytes(customXml))) part.FeedData(data);

        using WordDocument reopened = WordDocument.Load(new MemoryStream(document.ToBytes()));
        WordTable actual = Assert.Single(reopened.Tables);
        WordTableRow row = Assert.Single(actual.Rows);
        Assert.Equal(2, row.CellsCount);
        Assert.Equal(2, row.Cells[0].ColumnSpan);
        Assert.Equal("R0C2", CellText(row.Cells[1]));
        Assert.Equal("retained-cell", ContainerTag(row.Cells[1]._tableCell.Parent!));
        Assert.Equal("retained-row", ContainerTag(row._tableRow.Parent!));
        Assert.Empty(actual._table.Descendants<W.HorizontalMerge>());
        Assert.Equal(3, actual._table.GetFirstChild<W.TableGrid>()!.Elements<W.GridColumn>().Count());
        Assert.Empty(new OpenXmlValidator(FileFormatVersions.Office2013).Validate(actual._table));
        CustomXmlPart retained = Assert.Single(reopened.MainDocumentPartRoot.CustomXmlParts);
        using var reader = new StreamReader(retained.GetStream());
        Assert.Equal(customXml, reader.ReadToEnd());
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void CopyRowAppendsAfterTheLastLogicalRow(bool allRowsWrapped) {
        using WordDocument document = WordDocument.Create();
        WordTable table = CreateTable(document, 2, 1);
        if (allRowsWrapped) WrapRow(table.Rows[0]._tableRow, "customXml", "first-row");
        WrapRow(table.Rows[1]._tableRow, "sdt", "last-row");

        WordTableRow copy = table.CopyRow(table.Rows[0]);

        Assert.Equal(new[] { "R0C0", "R1C0", "R0C0" }, table.Rows.Select(row => CellText(row.Cells[0])));
        Assert.Same(copy, table.Rows[2]);
        Assert.Equal("last-row", ContainerTag(table.Rows[1]._tableRow.Parent!));
    }

    [Fact]
    public void InsertRowUsesTheLogicalZeroBasedPosition() {
        using WordDocument document = WordDocument.Create();
        var table = document.AddTable(3, 1);
        for (int index = 0; index < 3; index++) {
            table.Rows[index].Cells[0].Paragraphs[0].Text = "ROW" + index;
            WrapRow(table.Rows[index]._tableRow, index == 1 ? "customXml" : "sdt", "row-" + index);
        }
        table.InsertRow(1, "INSERTED");

        Assert.Equal(new[] { "ROW0", "INSERTED", "ROW1", "ROW2" }, table!.Rows.Select(row => CellText(row.Cells[0])));
        Assert.Equal(new[] { "row-0", "row-1", "row-2" }, table.Rows.Where(row => row._tableRow.Parent is not W.Table)
            .Select(row => ContainerTag(row._tableRow.Parent!)));
        Assert.Empty(new OpenXmlValidator(FileFormatVersions.Office2013).Validate(table._table));
    }

    [Fact]
    public void InteriorRowInsertionPreservesSharedContentControlAndSavedOrder() {
        using WordDocument document = WordDocument.Create();
        var table = document.AddTable(4, 1);
        for (int index = 0; index < 4; index++) table.Rows[index].Cells[0].Paragraphs[0].Text = "ROW" + index;
        W.TableRow[] rows = table.Rows.Select(row => row._tableRow).ToArray();
        WrapRow(rows[0], "sdt", "edge-row");
        var shared = new W.SdtRow(new W.SdtProperties(new W.Tag { Val = "shared-rows" }, new W.SdtId { Val = 710 }),
            new W.SdtContentRow());
        rows[1].Parent!.ReplaceChild(shared, rows[1]);
        shared.SdtContentRow!.Append(rows[1]);
        rows[2].Remove();
        shared.SdtContentRow.Append(rows[2]);

        table.InsertRow(2, "INTERIOR");

        Assert.Same(shared.SdtContentRow, table.Rows[2]._tableRow.Parent);
        Assert.Equal(new[] { "ROW1", "INTERIOR", "ROW2" }, shared.SdtContentRow.Elements<W.TableRow>()
            .Select(row => row.InnerText));
        Assert.Equal("edge-row", ContainerTag(rows[0].Parent!));
        Assert.Same(table._table, rows[3].Parent);

        using WordDocument reopened = WordDocument.Load(new MemoryStream(document.ToBytes()));
        WordTable actual = Assert.Single(reopened.Tables);
        Assert.Equal(new[] { "ROW0", "ROW1", "INTERIOR", "ROW2", "ROW3" }, actual.Rows.Select(row => CellText(row.Cells[0])));
        W.SdtRow sharedControl = Assert.Single(actual._table.Elements<W.SdtRow>(), control =>
            control.SdtProperties?.GetFirstChild<W.Tag>()?.Val?.Value == "shared-rows");
        Assert.Equal(710, sharedControl.SdtProperties!.GetFirstChild<W.SdtId>()!.Val!.Value);
        Assert.Equal(new[] { "ROW1", "INTERIOR", "ROW2" }, sharedControl.SdtContentRow!.Elements<W.TableRow>()
            .Select(row => row.InnerText));
        Assert.Equal("edge-row", ContainerTag(actual.Rows[0]._tableRow.Parent!));
        Assert.Same(actual._table, actual.Rows[4]._tableRow.Parent);
        Assert.Empty(new OpenXmlValidator(FileFormatVersions.Office2013).Validate(actual._table));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void TableCommentCoversTheFirstAndLastLogicalCells(bool wrappedRow) {
        using WordDocument document = WordDocument.Create();
        WordTable table = CreateTable(document, 1, 3);
        WrapCell(table.Rows[0].Cells[0]._tableCell, "sdt", "first-cell");
        WrapCell(table.Rows[0].Cells[2]._tableCell, "customXml", "last-cell");
        if (wrappedRow) WrapRow(table.Rows[0]._tableRow, "sdt", "only-row");

        table.AddComment("Synthetic reviewer", "SR", "Whole logical table");

        W.TableCell first = table.Rows[0].Cells[0]._tableCell;
        W.TableCell last = table.Rows[0].Cells[2]._tableCell;
        W.CommentRangeStart start = Assert.Single(first.Descendants<W.CommentRangeStart>());
        W.CommentRangeEnd end = Assert.Single(last.Descendants<W.CommentRangeEnd>());
        Assert.Equal(start.Id!.Value, end.Id!.Value);
        Assert.Equal(start.Id.Value, Assert.Single(last.Descendants<W.CommentReference>()).Id!.Value);
        Assert.Empty(table.Rows[0].Cells[1]._tableCell.Descendants<W.CommentRangeStart>());
        Assert.Equal("Whole logical table", Assert.Single(document.MainDocumentPartRoot.WordprocessingCommentsPart!.Comments!
            .Elements<W.Comment>()).InnerText);
    }

    private static WordTable CreateTable(WordDocument document, int rows, int columns) {
        WordTable table = document.AddTable(rows, columns);
        for (int row = 0; row < rows; row++)
            for (int column = 0; column < columns; column++)
                table.Rows[row].Cells[column].Paragraphs[0].Text = $"R{row}C{column}";
        return table;
    }

    private static WordTable CreateVerticalFixture(WordDocument document, string shape) {
        WordTable table = CreateTable(document, 3, 2);
        if (shape == "cell") WrapCell(table.Rows[1].Cells[0]._tableCell, "sdt", "target-cell");
        else if (shape == "row") WrapRow(table.Rows[1]._tableRow, "sdt", "target-row");
        else WrapCell(table.Rows[0].Cells[0]._tableCell, "customXml", "anchor-cell");
        return table;
    }

    private static string CellText(WordTableCell cell) => string.Concat(cell.Paragraphs.Select(paragraph => paragraph.Text));

    private static void WrapCell(W.TableCell cell, string kind, string tag) {
        OpenXmlElement wrapper = kind == "sdt"
            ? new W.SdtCell(new W.SdtProperties(new W.Tag { Val = tag }), new W.SdtContentCell())
            : new W.CustomXmlCell { Element = tag, Uri = "urn:officeimo:logical-table" };
        cell.Parent!.ReplaceChild(wrapper, cell);
        (wrapper is W.SdtCell control ? control.SdtContentCell! : wrapper).Append(cell);
    }

    private static void WrapRow(W.TableRow row, string kind, string tag) {
        OpenXmlElement wrapper = kind == "sdt"
            ? new W.SdtRow(new W.SdtProperties(new W.Tag { Val = tag }), new W.SdtContentRow())
            : new W.CustomXmlRow { Element = tag, Uri = "urn:officeimo:logical-table" };
        row.Parent!.ReplaceChild(wrapper, row);
        (wrapper is W.SdtRow control ? control.SdtContentRow! : wrapper).Append(row);
    }

    private static string? ContainerTag(OpenXmlElement parent) => parent switch {
        W.SdtContentCell content => ((W.SdtCell)content.Parent!).SdtProperties!.GetFirstChild<W.Tag>()!.Val!.Value,
        W.SdtContentRow content => ((W.SdtRow)content.Parent!).SdtProperties!.GetFirstChild<W.Tag>()!.Val!.Value,
        W.CustomXmlCell custom => custom.Element!.Value,
        W.CustomXmlRow custom => custom.Element!.Value,
        _ => null
    };
}
