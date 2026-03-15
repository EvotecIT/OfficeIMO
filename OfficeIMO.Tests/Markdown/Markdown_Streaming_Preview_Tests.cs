using OfficeIMO.Markdown;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class MarkdownStreamingPreviewTests {
    [Fact]
    public void NormalizeIntelligenceXTranscript_RepairsLineStartBulletsConservatively() {
        const string markdown = "-AD1\nstarkes Muster\n—** AD2** eher Secure-Channel";

        var normalized = MarkdownStreamingPreviewNormalizer.NormalizeIntelligenceXTranscript(markdown);

        Assert.Equal("- AD1 starkes Muster\n- **AD2** eher Secure-Channel", normalized);
    }

    [Fact]
    public void NormalizeIntelligenceXTranscript_RepairsSignalFlowTypographyArtifacts() {
        var markdown = string.Join("\n", [
            "- Signal **Catalog count includes hidden/disabled/deprecated rules -> **Why it matters:**external/custom rules can drift or disappear between hosts ->**Next action:**break down `rule_origin` (`builtin` vs `external`) and confirm expected external rules are present.**",
            "- TestimoX rules available ****359****"
        ]);

        var normalized = MarkdownStreamingPreviewNormalizer.NormalizeIntelligenceXTranscript(markdown);

        var expected = string.Join("\n", [
            "- Signal **Catalog count includes hidden/disabled/deprecated rules** -> **Why it matters:** external/custom rules can drift or disappear between hosts -> **Next action:** break down `rule_origin` (`builtin` vs `external`) and confirm expected external rules are present.",
            "- TestimoX rules available **359**"
        ]);

        Assert.Equal(expected, normalized);
    }

    [Fact]
    public void NormalizeIntelligenceXTranscript_DoesNotRewriteSignalFlowTypographyInsideFencedCode() {
        const string markdown = """
```text
- Signal **Catalog count includes hidden rules -> **Why it matters:**external drift ->**Next action:**compare inventories.**
- TestimoX rules available ****359****
```
""";

        var normalized = MarkdownStreamingPreviewNormalizer.NormalizeIntelligenceXTranscript(markdown);

        Assert.Equal(markdown, normalized);
    }
}
