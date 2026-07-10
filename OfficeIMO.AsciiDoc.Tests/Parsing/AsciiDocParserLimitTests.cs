namespace OfficeIMO.AsciiDoc.Tests;

public sealed class AsciiDocParserLimitTests {
    [Fact]
    public void MaximumInputLength_RejectsOversizedSourceBeforeParsing() {
        var options = new AsciiDocParseOptions { MaximumInputLength = 3 };

        Assert.Throws<ArgumentException>(() => AsciiDocDocument.Parse("four", options));
    }

    [Fact]
    public void MaximumBlockCount_RejectsAdditionalTopLevelBlocks() {
        var options = new AsciiDocParseOptions { MaximumBlockCount = 1 };

        Assert.Throws<InvalidDataException>(() => AsciiDocDocument.Parse("one\n\ntwo", options));
    }

    [Fact]
    public void MaximumBlockCount_AllowsTheExactConfiguredCount() {
        var options = new AsciiDocParseOptions { MaximumBlockCount = 1 };

        AsciiDocParseResult result = AsciiDocDocument.Parse("one", options);

        Assert.Single(result.Document.Blocks);
    }
}
