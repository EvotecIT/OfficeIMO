using OfficeIMO.Markdown;
using Xunit;

namespace OfficeIMO.Tests.MarkdownSuite;

public class Markdown_Input_Normalizer_Tests {
    [Fact]
    public void Normalize_DefaultOptions_LeavesInputUnchanged() {
        var markdown = "**Status\nHEALTHY** and `a\nb`";

        var normalized = MarkdownInputNormalizer.Normalize(markdown);
        Assert.Equal(markdown, normalized);
    }

    [Fact]
    public void Normalize_SoftWrappedStrong_WhenEnabled() {
        var options = new MarkdownInputNormalizationOptions {
            NormalizeSoftWrappedStrongSpans = true
        };

        var normalized = MarkdownInputNormalizer.Normalize("**Status\nHEALTHY**", options);
        Assert.Equal("**Status HEALTHY**", normalized);
    }

    [Fact]
    public void Normalize_InlineCodeLineBreaks_WhenEnabled() {
        var options = new MarkdownInputNormalizationOptions {
            NormalizeInlineCodeSpanLineBreaks = true
        };

        var normalized = MarkdownInputNormalizer.Normalize("`a\nb`", options);
        Assert.Equal("`a b`", normalized);
    }

    [Fact]
    public void Normalize_EscapedInlineCodeSpans_WhenEnabled() {
        var options = new MarkdownInputNormalizationOptions {
            NormalizeEscapedInlineCodeSpans = true
        };

        var normalized = MarkdownInputNormalizer.Normalize(@"Use \`/act act_001\` now.", options);
        Assert.Equal("Use `/act act_001` now.", normalized);
    }

    [Fact]
    public void Normalize_TightStrongBoundaries_WhenEnabled() {
        var options = new MarkdownInputNormalizationOptions {
            NormalizeTightStrongBoundaries = true
        };

        var normalized = MarkdownInputNormalizer.Normalize("Status **Healthy**next", options);
        Assert.Equal("Status **Healthy** next", normalized);
    }

    [Fact]
    public void Normalize_DoesNotChangeFencedCodeBlocks_ForEscapedCodeAndStrongSpacing() {
        var options = new MarkdownInputNormalizationOptions {
            NormalizeEscapedInlineCodeSpans = true,
            NormalizeTightStrongBoundaries = true
        };

        var markdown = """
```text
Use \`/act act_001\`
Status **Healthy**next
```
""";

        var normalized = MarkdownInputNormalizer.Normalize(markdown, options);
        Assert.Equal(markdown, normalized);
    }

    [Fact]
    public void Normalize_DoesNotChangeTildeFencedCodeBlocks() {
        var options = new MarkdownInputNormalizationOptions {
            NormalizeEscapedInlineCodeSpans = true,
            NormalizeTightStrongBoundaries = true
        };

        var markdown = """
~~~text
Use \`/act act_001\`
Status **Healthy**next
~~~
""";

        var normalized = MarkdownInputNormalizer.Normalize(markdown, options);
        Assert.Equal(markdown, normalized);
    }
}
