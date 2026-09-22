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

    [Fact]
    public void PowerPointSemanticRunsApplyTargetUrlTransformToOrdinaryLinks() {
        const string html = "<p><a href='https://example.com/rewrite'>Rewrite</a> " +
            "<a href='https://example.com/reject'>Reject</a></p>";
        var options = new HtmlToPowerPointOptions { Mode = HtmlImportMode.Generic };
        options.HyperlinkUrlPolicy.ResolvedUrlTransform = value =>
            value.EndsWith("/reject", StringComparison.Ordinal) ? null : "https://safe.example/rewritten";

        using PowerPointPresentation presentation = HtmlConversionDocument.Parse(html)
            .ToPowerPointPresentationResult(options).RequireValue();
        IReadOnlyList<PowerPointTextRun> runs = Assert.Single(presentation.Slides).TextBoxes
            .SelectMany(textBox => textBox.Paragraphs).SelectMany(paragraph => paragraph.Runs).ToList();

        Assert.Equal("https://safe.example/rewritten", Assert.Single(runs, run => run.Text == "Rewrite").Hyperlink?.OriginalString);
        Assert.Null(Assert.Single(runs, run => run.Text == "Reject").Hyperlink);
    }

    [Fact]
    public void PowerPointOrdinaryLinksApplyDocumentAndAdapterTransformsOnceEach() {
        var documentOptions = new HtmlConversionDocumentOptions();
        documentOptions.UrlPolicy.ResolvedUrlTransform = value => value + "?document=1";
        var adapterOptions = new HtmlToPowerPointOptions { Mode = HtmlImportMode.Generic };
        adapterOptions.HyperlinkUrlPolicy.ResolvedUrlTransform = value => value + "&adapter=1";

        using PowerPointPresentation presentation = HtmlConversionDocument.Parse(
            "<p><a href='https://example.com/path'>Link</a></p>", documentOptions)
            .ToPowerPointPresentationResult(adapterOptions).RequireValue();
        PowerPointTextRun run = Assert.Single(Assert.Single(presentation.Slides).TextBoxes
            .SelectMany(textBox => textBox.Paragraphs).SelectMany(paragraph => paragraph.Runs),
            item => item.Text == "Link");

        Assert.Equal("https://example.com/path?document=1&adapter=1", run.Hyperlink?.OriginalString);
    }

    [Fact]
    public void PowerPointTableLinksApplyDocumentAndAdapterTransformsOnceEach() {
        var documentOptions = new HtmlConversionDocumentOptions();
        documentOptions.UrlPolicy.ResolvedUrlTransform = value => value + "?document=1";
        var adapterOptions = new HtmlToPowerPointOptions { Mode = HtmlImportMode.Generic };
        adapterOptions.HyperlinkUrlPolicy.ResolvedUrlTransform = value => value + "&adapter=1";

        using PowerPointPresentation presentation = HtmlConversionDocument.Parse(
            "<table><tr><td><a href='https://example.com/path'>Link</a></td></tr></table>", documentOptions)
            .ToPowerPointPresentationResult(adapterOptions).RequireValue();
        PowerPointTextRun run = Assert.Single(Assert.Single(Assert.Single(presentation.Slides).Tables)
            .GetCell(0, 0).Paragraphs.SelectMany(paragraph => paragraph.Runs),
            item => item.Text == "Link");

        Assert.Equal("https://example.com/path?document=1&adapter=1", run.Hyperlink?.OriginalString);
    }
}
