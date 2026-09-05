namespace OfficeIMO.AsciiDoc.Tests;

public sealed class AsciiDocParserLimitTests {
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
