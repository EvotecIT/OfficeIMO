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

    [Theory]
    [InlineData("<p><a href='https://www.cdc.gov/'><img src='{0}' alt='CDC'></a></p>")]
    [InlineData("<a href='https://www.cdc.gov/'><img src='{0}' alt='CDC'></a>")]
    public void PowerPointHtmlRetainsLinkedPicturesAfterNativeReopen(string template) {
        const string image = "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNg+P//HwAF/gL9HjcXBgAAAABJRU5ErkJggg==";
        string html = string.Format(System.Globalization.CultureInfo.InvariantCulture, template, image);
        HtmlToPowerPointResult result = HtmlConversionDocument.Parse(html).ToPowerPointPresentationResult(
            new HtmlToPowerPointOptions { Mode = HtmlImportMode.Generic, ImportEditableLayoutRegions = false });
        using PowerPointPresentation presentation = result.RequireValue();
        using var stream = new MemoryStream();
        presentation.Save(stream);
        using PowerPointPresentation reopened = PowerPointPresentation.Load(new MemoryStream(stream.ToArray()),
            new PowerPointLoadOptions { AccessMode = OfficeIMO.DocumentAccessMode.ReadOnly });

        PowerPointPicture picture = Assert.Single(reopened.Slides.SelectMany(slide => slide.Pictures));
        Assert.Equal("https://www.cdc.gov/", picture.Hyperlink?.AbsoluteUri);
        Assert.DoesNotContain(result.Report.Diagnostics,
            diagnostic => diagnostic.Code == HtmlConversionDiagnosticCodes.ContentOmitted);
    }

    [Fact]
    public void PowerPointHtmlRetainsResolvableSlideFragmentOnPicture() {
        const string image = "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNg+P//HwAF/gL9HjcXBgAAAABJRU5ErkJggg==";
        HtmlToPowerPointResult result = HtmlConversionDocument.Parse(
            "<p><a href='#slide-1'><img src='" + image + "' alt='Back to start'></a></p>")
            .ToPowerPointPresentationResult(new HtmlToPowerPointOptions {
                Mode = HtmlImportMode.Generic,
                ImportEditableLayoutRegions = false
            });
        using PowerPointPresentation presentation = result.RequireValue();
        PowerPointPicture picture = Assert.Single(Assert.Single(presentation.Slides).Pictures);

        Assert.Equal("#slide-1", picture.Hyperlink?.OriginalString);
        Assert.DoesNotContain(result.Report.Diagnostics,
            diagnostic => diagnostic.Code == HtmlConversionDiagnosticCodes.ContentOmitted);
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
