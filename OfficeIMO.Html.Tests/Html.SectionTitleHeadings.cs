using OfficeIMO.Html;
using Xunit;

namespace OfficeIMO.Tests;

public partial class Html {
    [Theory]
    [InlineData("")]
    [InlineData("<p>Before</p>")]
    public void SemanticDocument_RetainsImageOnlySectionHeading(string before) {
        HtmlSemanticDocument document = HtmlConversionDocument.Parse(
            "<main>" + before + "<h1 id='brand'><img src='https://example.org/logo.png' alt='Brand'></h1><p>Body</p></main>")
            .SemanticDocument;

        HtmlSemanticBlock heading = document.Sections.Last().TitleHeading
            ?? Assert.Single(document.Sections.Last().Blocks, block => block.Kind == HtmlSemanticBlockKind.Heading);
        Assert.Equal(HtmlSemanticBlockKind.Heading, heading.Kind);
        Assert.Equal(1, heading.Level);
        Assert.Equal("h1", heading.SourceLocation!.ElementName);
        Assert.Equal("Brand", Assert.Single(heading.InlineResources).AlternateText);
        Assert.Equal("Brand", Assert.Single(document.Resources).AlternateText);
        Assert.Equal("Body", Assert.Single(document.Sections.Last().Blocks, block => block.Kind == HtmlSemanticBlockKind.Paragraph).Text);
    }

    [Fact]
    public void Preflight_ReportsRichLinkedPromotedHeading() {
        HtmlConversionPreflight preflight = HtmlConversionDocument.Parse(
            "<h1><a href='https://example.org/'><em>Title</em></a></h1><p>Body</p>")
            .AnalyzeFor(HtmlConversionTarget.Word);

        foreach (HtmlSemanticFeature feature in new[] { HtmlSemanticFeature.Headings, HtmlSemanticFeature.RichText, HtmlSemanticFeature.Links }) {
            HtmlFeaturePreflightResult result = preflight.Get(feature);
            Assert.True(result.IsPresent);
            Assert.Equal(1, result.OccurrenceCount);
        }
    }

    [Fact]
    public void FidelityScore_DetectsStyleLossInPromotedHeading() {
        HtmlRoundTripScore score = HtmlRoundTripScorer.Compare(
            "<h1 style='color:red'><em>Title</em></h1><p>Body</p>",
            "<h1>Title</h1><p>Body</p>");

        Assert.True(score.Dimensions.ContainsKey("styles"));
        Assert.True(score.Dimensions["styles"] < 1);
    }

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
