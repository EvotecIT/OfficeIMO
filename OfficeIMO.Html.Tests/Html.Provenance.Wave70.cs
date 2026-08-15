using AngleSharp.Html.Dom;
using OfficeIMO.Html;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class HtmlProvenanceWave70Tests {
    [Fact]
    public void SvgHrefNormalizationPreservesDoubleEscapedScriptText() {
        const string payload = "<!--<script><svg><image href='new' xlink:href='old'></image></svg></script>-->";
        string html = "<script>" + payload + "</script><p>after</p>";

        IHtmlDocument document = HtmlDocumentParser.ParseDocument(html);

        Assert.Equal(payload, document.QuerySelector("script")!.TextContent);
        Assert.Equal("after", document.QuerySelector("p")!.TextContent);
    }

    [Theory]
    [InlineData("1s cubic-bezier(2,0,0,1) spin")]
    [InlineData("1s steps(1,jump-none) spin")]
    [InlineData("1s linear(nope,1) spin")]
    [InlineData("1s cubic-bezier(2,0,0,1)")]
    public void AnimationShorthandRejectsMalformedTimingFunctions(string value) {
        Assert.False(HtmlResourcePipeline.TryExpandAnimationShorthandNames(value, out _));
    }

    [Theory]
    [InlineData("1s cubic-bezier(.2,0,.8,1) spin")]
    [InlineData("1s steps(2,jump-none) spin")]
    [InlineData("1s linear(0, 1 100%) spin")]
    public void AnimationShorthandAcceptsValidTimingFunctions(string value) {
        Assert.True(HtmlResourcePipeline.TryExpandAnimationShorthandNames(value, out string names));
        Assert.Equal("spin", names);
    }

    [Fact]
    public void AnimationShorthandAllowsTimingKeywordAsNameAfterTimingIsSet() {
        Assert.True(HtmlResourcePipeline.TryExpandAnimationShorthandNames("1s ease linear", out string names));
        Assert.Equal("linear", names);
    }
}
