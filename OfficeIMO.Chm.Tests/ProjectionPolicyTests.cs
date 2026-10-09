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
}
