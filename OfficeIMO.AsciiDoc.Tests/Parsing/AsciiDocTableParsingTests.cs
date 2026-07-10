namespace OfficeIMO.AsciiDoc.Tests;

public sealed class AsciiDocTableParsingTests {
    [Fact]
    public void PsvTable_ParsesColumnsHeaderSpansStylesAndEscapedSeparators() {
        const string source =
            "[cols=\"2,2\",%header]\n" +
            "|===\n" +
            "|Name |Value\n" +
            "2+|spanning\n" +
            ".2+^.^s|styled \\| literal\n" +
            "|other\n" +
            "|last\n" +
            "|===\n";

        AsciiDocDocument document = AsciiDocDocument.Parse(source).Document;
        AsciiDocTableBlock block = Assert.Single(document.BlocksOfType<AsciiDocTableBlock>());
        AsciiDocTable table = block.Table;

        Assert.Equal(AsciiDocTableFormat.Psv, table.Format);
        Assert.Equal("|", table.Separator);
        Assert.Equal(2, table.ColumnCount);
        Assert.Equal(6, table.Cells.Count);
        Assert.True(table.Rows[0].IsHeader);
        Assert.Equal(2, table.Cells[2].ColumnSpan);
        Assert.Equal(2, table.Cells[3].RowSpan);
        Assert.Equal('s', table.Cells[3].Style);
        Assert.Contains("\\| literal", table.Cells[3].Content, StringComparison.Ordinal);
        Assert.Equal(source, document.ToAsciiDoc());
        Assert.True(document.SyntaxTree.IsLossless);
        Assert.All(table.Cells, cell => Assert.Equal(AsciiDocSyntaxKind.TableCell, cell.Syntax.Kind));
    }

    [Fact]
    public void EditingPsvCellValue_PreservesSeparatorsSpecifiersAndRowBoundaries() {
        const string source = "[cols=2*]\r\n|===\r\n|A |B\r\n2+|old\r\n|===\r\n";
        AsciiDocDocument document = AsciiDocDocument.Parse(source).Document;
        AsciiDocTableBlock block = Assert.Single(document.BlocksOfType<AsciiDocTableBlock>());

        block.Table.Cells[2].Value = "new";

        Assert.Equal("[cols=2*]\r\n|===\r\n|A |B\r\n2+|new\r\n|===\r\n", document.ToAsciiDoc());
    }

    [Fact]
    public void CustomPsvSeparator_DoesNotSplitLiteralPipes() {
        const string source = "[cols=2*,separator=¦]\n|===\n¦A | literal ¦B\n|===\n";

        AsciiDocTable table = Assert.Single(AsciiDocDocument.Parse(source).Document.BlocksOfType<AsciiDocTableBlock>()).Table;

        Assert.Equal("¦", table.Separator);
        Assert.Equal(2, table.Cells.Count);
        Assert.Equal("A | literal", table.Cells[0].Value);
    }

    [Fact]
    public void CsvShorthand_ParsesQuotedCommasRowsAndEditsWithQuoting() {
        const string source = ",===\n\"A,B\",C\nD,E\n,===\n";
        AsciiDocDocument document = AsciiDocDocument.Parse(source).Document;
        AsciiDocTable table = Assert.Single(document.BlocksOfType<AsciiDocTableBlock>()).Table;

        Assert.Equal(AsciiDocTableFormat.Csv, table.Format);
        Assert.Equal(2, table.ColumnCount);
        Assert.Equal(2, table.Rows.Count);
        Assert.Equal("A,B", table.Cells[0].Value);
        Assert.Equal(source, document.ToAsciiDoc());

        table.Cells[3].Value = "E, F";
        Assert.Equal(",===\n\"A,B\",C\nD,\"E, F\"\n,===\n", document.ToAsciiDoc());
    }

    [Fact]
    public void DsvShorthand_UsesColonAndHonorsEscapes() {
        const string source = ":===\nA:B\\:C\nD:E\n:===\n";

        AsciiDocTable table = Assert.Single(AsciiDocDocument.Parse(source).Document.BlocksOfType<AsciiDocTableBlock>()).Table;

        Assert.Equal(AsciiDocTableFormat.Dsv, table.Format);
        Assert.Equal(2, table.ColumnCount);
        Assert.Equal("B:C", table.Cells[1].Value);
        Assert.Equal(4, table.Cells.Count);
    }

    [Fact]
    public void StyledDelimitedBlocks_ExposeAdmonitionAndStemSemantics() {
        const string source = "[WARNING]\n====\nDanger\n====\n[stem]\n++++\nx^2\n++++\n";
        AsciiDocDelimitedBlock[] blocks = AsciiDocDocument.Parse(source).Document.BlocksOfType<AsciiDocDelimitedBlock>().ToArray();

        Assert.Equal(AsciiDocAdmonitionKind.Warning, blocks[0].AdmonitionKind);
        Assert.True(blocks[1].IsStem);
    }
}
