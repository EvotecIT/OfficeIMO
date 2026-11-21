using Xunit;

namespace OfficeIMO.Yaml.Tests;

// Decoding cases the yaml-test-suite cannot reach: every suite file ends with a line break, and none has '' in a
// multi-line plain scalar. Both came from real files in the large corpus.
public class YamlScalarDecodingTests {
    [Fact]
    public void MultiLinePlainScalarsKeepDoubledQuotes() {
        YamlDocument document = YamlDocument.Parse("a: x != '' &&\n  y\n");
        Assert.Equal("x != '' && y", document.Select("a")!.AsString());
    }

    [Theory]
    [InlineData("a: |\n  x", "x")]
    [InlineData("a: >\n  x\n  y", "x y")]
    [InlineData("a: |+\n  x", "x")]
    [InlineData("a: |\n  x\n", "x\n")]
    public void BlockScalarsEndingAtEndOfInputAddNoLineBreak(string yaml, string value) {
        Assert.Equal(value, YamlDocument.Parse(yaml).Select("a")!.AsString());
    }
}
