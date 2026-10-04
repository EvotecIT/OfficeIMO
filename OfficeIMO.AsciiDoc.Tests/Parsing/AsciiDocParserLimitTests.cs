namespace OfficeIMO.AsciiDoc.Tests;

public sealed class AsciiDocParserLimitTests {
    [Theory]
    [InlineData("|===\n|a |b\n|===\n")]
    [InlineData(",===\na,b\n,===\n")]
    public void TableCellLimitsApplyDuringScanningForNativeAndDataTables(string source) {
        Assert.Throws<InvalidDataException>(() => AsciiDocDocument.Parse(source, new AsciiDocParseOptions { MaximumTableCellCount = 1 }));
        Assert.Equal(2, AsciiDocDocument.Parse(source, new AsciiDocParseOptions { MaximumTableCellCount = 2 }).BlocksOfType<AsciiDocTableBlock>().Single().Table.Cells.Count);
    }
    [Fact]
    public void TableCellLimitsAreCumulativeAndColumnOrSpanLimitsRejectOversizedDeclarations() {
        Assert.Throws<InvalidDataException>(() => AsciiDocDocument.Parse("|===\n|a\n|===\n\n|===\n|b\n|===\n", new AsciiDocParseOptions { MaximumTableCellCount = 1 }));
        Assert.Throws<InvalidDataException>(() => AsciiDocDocument.Parse("[cols=\"2147483647*\"]\n|===\n|a\n|===\n"));
        Assert.Throws<InvalidDataException>(() => AsciiDocDocument.Parse("|===\n10001+|a\n|===\n"));
    }
    [Fact]
    public void MaximumInputLength_RejectsOversizedSourceBeforeParsing() {
        var options = new AsciiDocParseOptions { MaximumInputLength = 3 };

        Assert.Throws<ArgumentException>(() => AsciiDocDocument.ParseResult("four", options));
    }

    [Fact]
    public void MaximumBlockCount_RejectsAdditionalTopLevelBlocks() {
        var options = new AsciiDocParseOptions { MaximumBlockCount = 1 };

        Assert.Throws<InvalidDataException>(() => AsciiDocDocument.ParseResult("one\n\ntwo", options));
    }

    [Fact]
    public void MaximumBlockCount_AllowsTheExactConfiguredCount() {
        var options = new AsciiDocParseOptions { MaximumBlockCount = 1 };

        AsciiDocParseResult result = AsciiDocDocument.ParseResult("one", options);

        Assert.Single(result.Document.Blocks);
    }
}
