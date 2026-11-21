using Xunit;

namespace OfficeIMO.Yaml.Tests;

// Consumers that report a file which does not load (Reader.Yaml, Markdown front matter, PowerShell) take the failure in
// the shared diagnostic form, next to the warnings of files that do load.
public class YamlExceptionTests {
    [Theory]
    [InlineData("a: [1, 2\n", "yaml-syntax-unexpected-token", "2:1")]
    [InlineData("a: 1\na: 2\n", "yaml-duplicate-key", "2:1")]
    [InlineData("a: *missing\n", "yaml-load-undefined-alias", "1:4")]
    [InlineData("[[[x]]]\n", "yaml-limit-exceeded", "1:3")]
    public void FailuresConvertToPositionedErrorDiagnosticsWithCatalogCodes(string yaml, string code, string location) {
        var options = new YamlLoadOptions { Strict = true, MaxDepth = 2 };
        YamlException error = Assert.ThrowsAny<YamlException>(() => YamlDocument.Parse(yaml, options));

        OfficeDiagnostic diagnostic = error.ToDiagnostic();

        Assert.Equal(code, diagnostic.Code);
        Assert.Equal(OfficeDiagnosticSeverity.Error, diagnostic.Severity);
        Assert.Equal(location, diagnostic.Location);
        Assert.Equal("YAML", diagnostic.Format);
        // The location is a field of its own; the message keeps the source excerpt but not the position text.
        Assert.DoesNotContain("(line ", diagnostic.Message, StringComparison.Ordinal);
        Assert.EndsWith("^", diagnostic.Message, StringComparison.Ordinal);
    }
}
