using OfficeIMO.Html;
using Xunit;

namespace OfficeIMO.Tests;

public partial class Html {
    [Theory]
    [InlineData(1)]
    [InlineData(2)]
    public void SemanticDocument_RetainsPromotedHeadingFormattingAndSource(int level) {
        HtmlSemanticSection section = Assert.Single(HtmlConversionDocument.Parse(
            $"<main><h{level} id='intro' style='color:#123456'><em>Intro</em></h{level}><p>Body</p></main>")
            .SemanticDocument.Sections);

        Assert.Equal("Intro", section.Title);
        HtmlSemanticBlock heading = Assert.IsType<HtmlSemanticBlock>(section.TitleHeading);
        Assert.Equal(HtmlSemanticBlockKind.Heading, heading.Kind);
        Assert.Equal("Intro", heading.Text);
        Assert.Equal(level, heading.Level);
        Assert.Contains(heading.Runs, run => run.Italic && run.Text == "Intro");
        Assert.Equal("#123456", heading.Style!.Properties["color"]);
        Assert.Equal("h" + level, heading.SourceLocation!.ElementName);
        Assert.Contains("intro", heading.SourceLocation.Selector);
        Assert.Equal("Body", Assert.Single(section.Blocks).Text);
    }

    [Fact]
    public void SemanticDocument_RetainsEachImplicitTitleHeadingInSectionOrder() {
        IReadOnlyList<HtmlSemanticSection> sections = HtmlConversionDocument.Parse(
            "<main><h1>First</h1><p>One</p><h2>Second</h2><p>Two</p><h1>Third</h1></main>")
            .SemanticDocument.Sections;

        Assert.Equal(new[] { "First", "Second", "Third" }, sections.Select(section => section.TitleHeading!.Text));
        Assert.Equal(new[] { 1, 2, 1 }, sections.Select(section => section.TitleHeading!.Level));
        Assert.Equal(new[] { "One", "Two" }, sections.SelectMany(section => section.Blocks).Select(block => block.Text));
    }

    [Theory]
    [InlineData("<article><h1>Title</h1><p>Body</p></article>")]
    [InlineData("<html><head><title>Title</title></head><body><h1>Title</h1><p>Body</p></body></html>")]
    public void SemanticDocument_DoesNotDuplicateHeadingRetainedInSectionBody(string html) {
        HtmlSemanticSection section = Assert.Single(HtmlConversionDocument.Parse(html).SemanticDocument.Sections);

        Assert.Null(section.TitleHeading);
        Assert.Equal("Title", Assert.Single(section.Blocks, block => block.Kind == HtmlSemanticBlockKind.Heading).Text);
    }
}
