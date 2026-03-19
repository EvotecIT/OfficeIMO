using System;
using OfficeIMO.MarkdownRenderer;
using Xunit;

namespace OfficeIMO.MarkdownRenderer.Wpf.Tests;

public class MarkdownViewSupportTests {
    [Fact]
    public void AppendCss_CombinesTrimmedBlocksWithSingleSeparator() {
        string result = MarkdownViewSupport.AppendCss("  body { color: red; }  ", "  :root { color-scheme: dark; }  ");

        Assert.Equal("body { color: red; }" + Environment.NewLine + ":root { color-scheme: dark; }", result);
    }

    [Fact]
    public void AppendCss_ReturnsExistingWhenAdditionalIsBlank() {
        string result = MarkdownViewSupport.AppendCss("  body { color: red; }  ", "   ");

        Assert.Equal("body { color: red; }", result);
    }

    [Theory]
    [InlineData("https://github.com/EvotecIT", "https://github.com/EvotecIT")]
    [InlineData(" file:///C:/Temp/readme.md ", "file:///C:/Temp/readme.md")]
    public void TryNormalizeBaseHref_AcceptsAbsoluteUris(string rawValue, string expected) {
        bool success = MarkdownViewSupport.TryNormalizeBaseHref(rawValue, out string normalized);

        Assert.True(success);
        Assert.Equal(expected, normalized);
    }

    [Theory]
    [InlineData("")]
    [InlineData("relative/path")]
    [InlineData("not a uri")]
    public void TryNormalizeBaseHref_RejectsInvalidValues(string rawValue) {
        bool success = MarkdownViewSupport.TryNormalizeBaseHref(rawValue, out string normalized);

        Assert.False(success);
        Assert.Equal(string.Empty, normalized);
    }

    [Theory]
    [InlineData("https://example.com/readme", true)]
    [InlineData("mailto:test@example.com", true)]
    [InlineData("about:blank", false)]
    [InlineData("data:text/html,hello", false)]
    [InlineData("javascript:alert(1)", false)]
    public void TryGetExternalNavigationUri_FiltersUnsafeSchemes(string rawUri, bool expected) {
        bool success = MarkdownViewSupport.TryGetExternalNavigationUri(rawUri, out Uri navigationUri);

        Assert.Equal(expected, success);
        if (expected) {
            Assert.NotNull(navigationUri);
        }
    }

    [Fact]
    public void TryGetClipboardText_ReturnsCopyPayload() {
        bool success = MarkdownViewSupport.TryGetClipboardText("""{"type":"omd.copy","text":"copied value"}""", out string clipboardText);

        Assert.True(success);
        Assert.Equal("copied value", clipboardText);
    }

    [Theory]
    [InlineData("""{"type":"other","text":"ignored"}""")]
    [InlineData("""{"type":"omd.copy"}""")]
    [InlineData("""{"type":"omd.copy","text":""}""")]
    [InlineData("""not json""")]
    public void TryGetClipboardText_RejectsMalformedOrUnsupportedMessages(string payload) {
        bool success = MarkdownViewSupport.TryGetClipboardText(payload, out string clipboardText);

        Assert.False(success);
        Assert.Equal(string.Empty, clipboardText);
    }

    [Fact]
    public void CreateEffectiveOptions_AppliesPresetAndHostOverrides() {
        MarkdownRendererOptions options = MarkdownViewSupport.CreateEffectiveOptions(
            MarkdownViewPreset.Relaxed,
            "https://github.com/EvotecIT",
            ":root { color-scheme: dark; }",
            static effectiveOptions => {
                effectiveOptions.EnableCodeCopyButtons = true;
                effectiveOptions.EnableTableCopyButtons = true;
            });

        Assert.Equal("https://github.com/EvotecIT", options.BaseHref);
        Assert.Contains(":root { color-scheme: dark; }", options.ShellCss);
        Assert.True(options.EnableCodeCopyButtons);
        Assert.True(options.EnableTableCopyButtons);
        Assert.False(options.HtmlOptions.BlockExternalHttpImages);
        Assert.False(options.HtmlOptions.RestrictHttpLinksToBaseOrigin);
    }

    [Fact]
    public void CreateEffectiveOptions_UsesStrictMinimalPresetDefaults() {
        MarkdownRendererOptions options = MarkdownViewSupport.CreateEffectiveOptions(
            MarkdownViewPreset.StrictMinimal,
            null,
            null,
            null);

        Assert.False(options.EnableCodeCopyButtons);
        Assert.False(options.EnableTableCopyButtons);
        Assert.False(options.Math.Enabled);
        Assert.False(options.Mermaid.Enabled);
        Assert.False(options.Chart.Enabled);
        Assert.False(options.Network.Enabled);
    }
}
