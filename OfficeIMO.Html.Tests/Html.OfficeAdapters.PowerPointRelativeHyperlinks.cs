using OfficeIMO.Html;
using OfficeIMO.PowerPoint;
using OfficeIMO.PowerPoint.Html;
using Xunit;

namespace OfficeIMO.Tests;

public partial class HtmlOfficeAdapters {
    [Fact]
    public void PowerPointHtmlImportsPolicyApprovedRelativeRunHyperlinks() {
        const string html = "<p><a href='#slide-2'>Fragment</a> <a href='/docs'>Root</a> <a href='javascript:alert(1)'>Rejected</a></p>";

        HtmlToPowerPointResult result = HtmlConversionDocument.Parse(html)
            .ToPowerPointPresentationResult(new HtmlToPowerPointOptions { Mode = HtmlImportMode.Generic });
        using PowerPointPresentation presentation = result.RequireValue();
        IReadOnlyList<PowerPointTextRun> runs = Assert.Single(presentation.Slides).TextBoxes
            .SelectMany(textBox => textBox.Paragraphs)
            .SelectMany(paragraph => paragraph.Runs)
            .ToList();

        Assert.Equal("#slide-2", Assert.Single(runs, run => run.Text == "Fragment").Hyperlink?.OriginalString);
        Assert.Equal("/docs", Assert.Single(runs, run => run.Text == "Root").Hyperlink?.OriginalString);
        Assert.Null(Assert.Single(runs, run => run.Text == "Rejected").Hyperlink);
    }

    [Fact]
    public void PowerPointSemanticRunsRespectTheTargetHyperlinkPolicy() {
        const string html = "<p><a href='https://example.com/allowed'>Allowed</a> " +
            "<a href='http://example.com/rejected'>Rejected</a></p>";
        var options = new HtmlToPowerPointOptions { Mode = HtmlImportMode.Generic };
        options.HyperlinkUrlPolicy.RestrictUrlSchemes = true;
        options.HyperlinkUrlPolicy.AllowedUrlSchemes.Clear();
        options.HyperlinkUrlPolicy.AllowedUrlSchemes.Add("https");

        HtmlToPowerPointResult result = HtmlConversionDocument.Parse(html)
            .ToPowerPointPresentationResult(options);
        using PowerPointPresentation presentation = result.RequireValue();
        IReadOnlyList<PowerPointTextRun> runs = Assert.Single(presentation.Slides).TextBoxes
            .SelectMany(textBox => textBox.Paragraphs)
            .SelectMany(paragraph => paragraph.Runs)
            .ToList();

        Assert.Equal("https://example.com/allowed", Assert.Single(runs, run => run.Text == "Allowed").Hyperlink?.OriginalString);
        Assert.Null(Assert.Single(runs, run => run.Text == "Rejected").Hyperlink);
    }
}
