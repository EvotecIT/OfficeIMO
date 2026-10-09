using OfficeIMO.Html;
using OfficeIMO.Markdown.Html;
using OfficeIMO.Reader;
using OfficeIMO.Reader.Chm;
using OfficeIMO.Reader.Html;
using System.Threading;

namespace OfficeIMO.Chm.Tests;

public sealed class ProjectionPolicyTests {
    [Theory]
    [InlineData("markdown")]
    [InlineData("normalization")]
    [InlineData("conversion")]
    public void ReaderTopicIdentityOverridesCallerProjectionBasesWithoutMutatingOptions(string callerBase) {
        var projectionBase = new Uri("https://example.invalid/unrelated/");
        var conversionBase = new Uri("https://other.invalid/unrelated/");
        var htmlOptions = ReaderHtmlOptions.CreateOfficeIMOProfile();
        htmlOptions.HtmlToMarkdownOptions!.BaseUri = callerBase == "markdown" ? projectionBase : null;
        htmlOptions.ConversionOptions = new HtmlConversionDocumentOptions { BaseUri = conversionBase };
        if (callerBase == "normalization") htmlOptions.ConversionOptions.NormalizationOptions = new HtmlNormalizationOptions {
            BaseUri = projectionBase, BaseElementBaseUri = projectionBase
        };
        var options = new ReaderChmOptions { HtmlOptions = htmlOptions };
        var reader = new OfficeDocumentReaderBuilder().AddChmHandler(options).Build();
        using var stream = new MemoryStream(ChmFixture.Archive(new Dictionary<string, byte[]> {
            ["/guide/topic.html"] = ChmFixture.Html("<a href='next.html#anchor'>Next</a><img src='../pixel.png'><video src='clip.mp4'></video>"),
            ["/guide/next.html"] = ChmFixture.Html("<p id='anchor'>Next topic</p>"),
            ["/pixel.png"] = ChmDocument.Load(ChmFixture.Book()).FindEntry("/pixel.png")!.GetBytes()
        }));
        var result = reader.ReadDocument(stream, "manual.chm");
        Assert.Contains(result.Links, link => link.Uri == "chm://archive/guide/next.html#anchor");
        Assert.Contains(result.Visuals, visual => visual.SourceName == "chm://archive/pixel.png");
        Assert.Contains(result.Visuals, visual => visual.SourceName == "chm://archive/guide/clip.mp4");
        Assert.DoesNotContain("example.invalid", result.Markdown!);
        Assert.DoesNotContain("other.invalid", result.Markdown!);
        Assert.Equal(callerBase == "markdown" ? projectionBase : null, htmlOptions.HtmlToMarkdownOptions.BaseUri);
        Assert.Equal(conversionBase, htmlOptions.ConversionOptions.BaseUri);
        if (callerBase == "normalization") {
            Assert.Equal(projectionBase, htmlOptions.ConversionOptions.NormalizationOptions!.BaseUri);
            Assert.Equal(projectionBase, htmlOptions.ConversionOptions.NormalizationOptions.BaseElementBaseUri);
        }
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void ReaderCancellationStopsPolicyNormalizationAtTheFirstCallback(bool chm) {
        using var cancelled = new CancellationTokenSource();
        int calls = 0;
        var options = ReaderHtmlOptions.CreateOfficeIMOProfile();
        options.ConversionOptions = new HtmlConversionDocumentOptions();
        options.ConversionOptions.UrlPolicy.ResolvedUrlTransform = value => {
            calls++;
            cancelled.Cancel();
            return value;
        };
        var builder = new OfficeDocumentReaderBuilder();
        if (chm) builder.AddChmHandler(new ReaderChmOptions { HtmlOptions = options });
        else builder.AddHtmlHandler(options);
        var reader = builder.Build();
        string html = string.Concat(Enumerable.Range(0, 8).Select(index => "<a href='https://example.invalid/" + index + "'>Link</a>"));
        byte[] bytes = chm ? ChmFixture.Archive(new Dictionary<string, byte[]> { ["/topic.html"] = ChmFixture.Html(html) }) : ChmFixture.Html(html);
        using var stream = new MemoryStream(bytes);
        OperationCanceledException error = Assert.ThrowsAny<OperationCanceledException>(() =>
            reader.ReadDocument(stream, chm ? "manual.chm" : "topic.html", cancellationToken: cancelled.Token));
        Assert.Equal(cancelled.Token, error.CancellationToken);
        Assert.Equal(1, calls);
    }

    [Theory]
    [InlineData(false, false)]
    [InlineData(false, true)]
    [InlineData(true, false)]
    [InlineData(true, true)]
    public void ReaderAuthoredRelativeBaseResolvesOnceAcrossMarkdownAndRichReferences(bool chm, bool filtered) {
        var options = ReaderHtmlOptions.CreateOfficeIMOProfile();
        var sourceBase = new Uri("https://example.invalid/guide/topic.html");
        options.HtmlToMarkdownOptions!.BaseUri = sourceBase;
        if (filtered) options.HtmlToMarkdownOptions.ExcludeSelectors.Add(".remove");
        string html = "<html><head><base href='assets/'></head><body><p class='remove'>Omit</p><a href='next.html'>Next</a><img src='image.png' alt='Image'></body></html>";
        var result = ReadTopic(html, options, chm);
        string root = chm ? "chm://archive/guide/assets/" : "https://example.invalid/guide/assets/";
        Assert.Contains(result.Links, link => link.Uri == root + "next.html");
        Assert.Contains(result.Visuals, visual => visual.SourceName == root + "image.png");
        Assert.Contains("[Next](" + root + "next.html)", result.Markdown!);
        Assert.Contains("![Image](" + root + "image.png)", result.Markdown!);
        Assert.DoesNotContain("assets/assets/", result.Markdown!);
        Assert.Equal(sourceBase, options.HtmlToMarkdownOptions.BaseUri);
    }

    [Theory]
    [InlineData(false, false)]
    [InlineData(false, true)]
    [InlineData(true, false)]
    [InlineData(true, true)]
    public void ReaderAppliesDocumentAndProjectionTransformsOnceToLinksAndImages(bool chm, bool filtered) {
        var options = ReaderHtmlOptions.CreateOfficeIMOProfile();
        options.HtmlToMarkdownOptions!.BaseUri = new Uri("https://example.invalid/guide/topic.html");
        options.ConversionOptions = new HtmlConversionDocumentOptions();
        options.ConversionOptions.UrlPolicy.ResolvedUrlTransform = uri => uri + "?document=1";
        options.ConversionOptions.ResourceUrlPolicy.ResolvedUrlTransform = uri => uri + "?document=1";
        options.HtmlToMarkdownOptions.UrlPolicy.ResolvedUrlTransform = uri => uri + "&projection=1";
        options.HtmlToMarkdownOptions.ResourceUrlPolicy = new HtmlUrlPolicy { ResolvedUrlTransform = uri => uri + "&projection=1" };
        if (filtered) options.HtmlToMarkdownOptions.ElementFilters.Add(element => false);
        var result = ReadTopic("<a href='next.html'>Next</a><img src='image.png' alt='Image'>", options, chm);
        string root = chm ? "chm://archive/guide/" : "https://example.invalid/guide/";
        string link = root + "next.html?document=1&projection=1";
        string image = root + "image.png?document=1&projection=1";
        Assert.Contains(result.Links, item => item.Uri == link);
        Assert.Contains(result.Visuals, item => item.SourceName == image);
        Assert.Contains("[Next](" + link + ")", result.Markdown!);
        Assert.Contains("![Image](" + image + ")", result.Markdown!);
        Assert.DoesNotContain("?document=1?document=1", result.Markdown!);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void ReaderPreparedProjectionKeepsDocumentRestrictionsOnProjectionTransforms(bool chm) {
        var options = ReaderHtmlOptions.CreateOfficeIMOProfile();
        options.ConversionOptions = new HtmlConversionDocumentOptions();
        options.ConversionOptions.UrlPolicy.DisallowFileUrls = true;
        options.ConversionOptions.ResourceUrlPolicy.DisallowFileUrls = true;
        options.HtmlToMarkdownOptions!.UrlPolicy.ResolvedUrlTransform = uri => "file:///private/link";
        options.HtmlToMarkdownOptions.ResourceUrlPolicy = new HtmlUrlPolicy { ResolvedUrlTransform = uri => "file:///private/image" };
        options.HtmlToMarkdownOptions.ElementFilters.Add(element => false);
        var result = ReadTopic("<a href='https://example.invalid/next'>Next</a><img src='https://example.invalid/image.png' alt='Image'>", options, chm);
        Assert.DoesNotContain("file:", result.Markdown!);
        Assert.DoesNotContain("![Image]", result.Markdown!);
    }

    private static OfficeDocumentReadResult ReadTopic(string html, ReaderHtmlOptions options, bool chm) {
        var builder = new OfficeDocumentReaderBuilder();
        if (chm) builder.AddChmHandler(new ReaderChmOptions { HtmlOptions = options });
        else builder.AddHtmlHandler(options);
        byte[] bytes = chm ? ChmFixture.Archive(new Dictionary<string, byte[]> { ["/guide/topic.html"] = ChmFixture.Html(html) }) : ChmFixture.Html(html);
        using var stream = new MemoryStream(bytes);
        return builder.Build().ReadDocument(stream, chm ? "manual.chm" : "topic.html");
    }
}
