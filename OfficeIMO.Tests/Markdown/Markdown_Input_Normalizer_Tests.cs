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
}
