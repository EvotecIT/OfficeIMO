using OfficeIMO.Html;
using Xunit;

namespace OfficeIMO.Tests;

public class HtmlDomTraversalTests {
    [Theory]
    [InlineData("br")]
    [InlineData("IMG")]
    [InlineData(" input ")]
    public void HtmlDomElementFacts_Recognizes_Html_Void_Elements(string localName) {
        Assert.True(HtmlDomElementFacts.IsVoidElement(localName));
    }

    [Theory]
    [InlineData(null)]
    [InlineData("")]
    [InlineData("p")]
    [InlineData("section")]
    public void HtmlDomElementFacts_Rejects_NonVoid_Element_Names(string? localName) {
        Assert.False(HtmlDomElementFacts.IsVoidElement(localName));
    }

    [Fact]
    public void HtmlDomLimitTracker_Returns_Null_When_Unbounded() {
        Assert.Null(HtmlDomLimitTracker.Create(maxHtmlNodes: null, maxHtmlDepth: null));
    }

    [Fact]
    public void HtmlDomLimitTracker_Stops_When_Node_Limit_Is_Exceeded() {
        HtmlDomLimitTracker tracker = new HtmlDomLimitTracker(maxHtmlNodes: 1, maxHtmlDepth: null);
        tracker.RecordNode();

        HtmlDomLimitException exception = Assert.Throws<HtmlDomLimitException>(() => tracker.RecordNode());

        Assert.Equal("HtmlNodeLimitExceeded", exception.Code);
        Assert.Equal("MaxHtmlNodes", exception.LimitSource);
        Assert.Equal(2, exception.Actual);
        Assert.Equal(1, exception.Limit);
        Assert.Contains("Actual=2", exception.Detail, StringComparison.Ordinal);
        Assert.Contains("Limit=1", exception.Detail, StringComparison.Ordinal);
    }

    [Fact]
    public void HtmlDomLimitTracker_Stops_When_Depth_Limit_Is_Exceeded() {
        HtmlDomLimitTracker tracker = new HtmlDomLimitTracker(maxHtmlNodes: null, maxHtmlDepth: 2);

        HtmlDomLimitException exception = Assert.Throws<HtmlDomLimitException>(() => tracker.RecordElementStart(3));

        Assert.Equal("HtmlDepthLimitExceeded", exception.Code);
        Assert.Equal("MaxHtmlDepth", exception.LimitSource);
        Assert.Equal(3, exception.Actual);
        Assert.Equal(2, exception.Limit);
    }
}
