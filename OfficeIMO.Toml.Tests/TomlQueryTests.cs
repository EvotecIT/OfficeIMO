using System.Globalization;
using Xunit;

namespace OfficeIMO.Toml.Tests;

public class TomlQueryTests {
    private const string Lock = """
        version = 3

        [[package]]
        name = "serde"
        version = "1.0.210"
        dependencies = ["serde_derive"]

        [[package]]
        name = "serde_derive"
        version = "1.0.210"
        checksum = "abc"

        [[package]]
        name = "toml"
        version = "0.8.19"

        [tool.limits]
        max = 10
        ratio = 0.5
        strict = true
        """;

    [Theory]
    [InlineData("package[*].name", "serde,serde_derive,toml")]
    [InlineData("package.*.name", "serde,serde_derive,toml")]
    [InlineData("package[-1].name", "toml")]
    [InlineData("package[?(@.name == 'serde')].version", "1.0.210")]
    [InlineData("package[?(@.checksum)].name", "serde_derive")]
    [InlineData("package[?(@.dependencies[0] == 'serde_derive')].name", "serde")]
    [InlineData("package[?(@.version != '1.0.210')].name", "toml")]
    [InlineData("tool[?(@.max > 5)].strict", "True")]
    [InlineData("tool.limits[?(@ == 0.5)]", "0.5")]
    [InlineData("tool.limits[?(@ == true)]", "True")]
    public void QueriesSelectNodesInDocumentOrder(string query, string expected) {
        IReadOnlyList<TomlQueryMatch> matches = TomlDocument.Parse(Lock).Query(query);
        Assert.Equal(expected, string.Join(",", matches.Select(m => Convert.ToString(((TomlValue)m.Node).Value, CultureInfo.InvariantCulture))));
    }

    [Fact]
    public void PathsAddressEachMatchAndNodesKeepTheirPositions() {
        TomlDocument document = TomlDocument.Parse(Lock);

        IReadOnlyList<TomlQueryMatch> matches = document.Query("..version");

        Assert.Equal(new[] { "version", "package[0].version", "package[1].version", "package[2].version" }, matches.Select(m => m.Path));
        Assert.All(matches, m => Assert.Same(m.Node, document.Select(m.Path)));
        Assert.Equal(new[] { 1, 5, 10, 15 }, matches.Select(m => m.Node.Start.Line));
        Assert.Equal("$", document.Query("$").Single().Path);
    }

    [Fact]
    public void EverySelectPathIsAQueryForTheSameNode() {
        TomlDocument document = TomlDocument.Parse(Lock);
        foreach (string path in new[] { "package[1].checksum", "$.tool.limits.max", "[\"version\"]", "package[0].dependencies[0]" }) {
            Assert.Same(document.Select(path), document.Query(path).Single().Node);
        }
        Assert.Empty(document.Query("package[9].name"));
    }

    [Fact]
    public void NodeQueriesAreRelativeToTheNode() {
        TomlNode packages = TomlDocument.Parse(Lock).Select("package")!;

        TomlQueryMatch match = packages.Query("[?(@.name == 'toml')].version").Single();

        Assert.Equal("[2].version", match.Path);
        Assert.Same(match.Node, packages.Select(match.Path));
    }

    [Theory]
    [InlineData("package[")]
    [InlineData("package..")]
    [InlineData("package[?(@.name ==)]")]
    [InlineData("package[?(1)]")]
    public void MalformedQueriesThrowWhateverTheData(string query) {
        Assert.Throws<ArgumentException>(() => TomlDocument.Parse(Lock).Query(query));
        Assert.Throws<ArgumentException>(() => TomlDocument.Parse(string.Empty).Query(query));
    }
}
