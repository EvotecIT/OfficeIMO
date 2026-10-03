using OfficeIMO.Adf;
using Xunit;

namespace OfficeIMO.Adf.Tests;

public sealed class AdfMediaShapeTests {
    [Theory]
    [InlineData("media", null, true)]
    [InlineData("media", "caption", true)]
    [InlineData("media", "media", false)]
    [InlineData("caption", "media", false)]
    [InlineData("caption", null, false)]
    public void MediaSingleRequiresMediaThenOptionalCaption(string first, string? second, bool valid) {
        var container = new AdfNode("mediaSingle");
        container.Content.Add(Child(first));
        if (second != null) container.Content.Add(Child(second));
        var document = new AdfDocument(new[] { container });
        var result = document.Validate();

        Assert.Equal(valid, result.IsValid);
        if (valid) Assert.Empty(result.Issues);
        else Assert.Contains(result.Issues, issue => issue.Code == "ADF_MEDIA_SINGLE_CONTENT");
    }

    [Fact]
    public void CaptionAllowsInlineContentButRejectsBlocks() {
        var caption = new AdfNode("caption") { Content = { AdfNode.TextNode("caption", new[] { new AdfMark("strong") }), new AdfNode("hardBreak") } };
        var container = new AdfNode("mediaSingle") { Content = { Child("media"), caption } };
        var document = new AdfDocument(new[] { container });
        Assert.Empty(document.Validate().Issues);
        Assert.Empty(AdfDocument.Parse(document.ToJson()).Validate().Issues);

        caption.Content.Add(new AdfNode("paragraph"));
        Assert.Contains(document.Validate().Issues, issue => issue.Code == "ADF_NODE_CHILD" && issue.Path == "$.content[0].content[1].content[2]");
    }

    private static AdfNode Child(string type) => type == "media"
        ? new AdfNode("media").SetAttribute("type", "external").SetAttribute("url", "https://example.com/a.png")
        : new AdfNode(type);
}
