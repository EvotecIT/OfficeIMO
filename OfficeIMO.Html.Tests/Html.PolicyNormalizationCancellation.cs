using System.Threading;
using OfficeIMO.Html;
using OfficeIMO.Html.Dom;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class HtmlPolicyNormalizationCancellationTests {
    [Theory]
    [InlineData("<a href='one.html'>One</a><a href='two.html'>Two</a>", false)]
    [InlineData("<img src='one.png'><img src='two.png'>", true)]
    [InlineData("<img srcset='one.png 1x, two.png 2x'>", false)]
    [InlineData("<style>p { background: url(one.png), url(two.png); }</style>", true)]
    [InlineData("<style>@import 'one.css'; @import 'two.css';</style>", false)]
    [InlineData("<p style='background: image-set(&quot;one.png&quot; 1x, &quot;two.png&quot; 2x)'>Text</p>", true)]
    [InlineData("<meta http-equiv='refresh' content='0; url=one.html; url=two.html'>", false)]
    [InlineData("<iframe srcdoc='&lt;a href=&quot;one.html&quot;>One&lt;/a>&lt;a href=&quot;two.html&quot;>Two&lt;/a>'></iframe>", true)]
    public void CancellationDuringNormalizationStopsCallbacksAndAllowsLaterReuse(string html, bool explicitMedia) {
        using var cancelled = new CancellationTokenSource();
        int calls = 0;
        bool cancel = true;
        var policy = HtmlUrlPolicy.CreateOfficeIMOProfile();
        policy.ResolvedUrlTransform = value => {
            calls++;
            if (cancel) cancelled.Cancel();
            return value;
        };
        var source = HtmlConversionDocument.Parse("<input value='authored'>" + html).Document.CloneAttached();
        source.QuerySelector("input")!.FormState = new HtmlFormControlState(HtmlFormControlStateKind.Input, "live");
        var document = HtmlConversionDocument.FromDocument(source, new HtmlConversionDocumentOptions {
            BaseUri = new Uri("https://example.invalid/"), UrlPolicy = policy, ResourceUrlPolicy = policy
        });
        OperationCanceledException error = Assert.ThrowsAny<OperationCanceledException>(() => {
            if (explicitMedia) document.CreateDocumentForConversion(HtmlCssMediaContext.Print, cancelled.Token);
            else document.CreateDocumentForConversion(cancelled.Token);
        });
        Assert.Equal(cancelled.Token, error.CancellationToken);
        Assert.Equal(1, calls);
        cancel = false;
        HtmlDocument projected = document.CreateDocumentForConversion();
        Assert.Equal("live", projected.QuerySelector("input")!.FormState!.Value);
        projected.QuerySelector("input")!.FormState = new HtmlFormControlState(HtmlFormControlStateKind.Input, "edited");
        Assert.Equal("live", document.Document.QuerySelector("input")!.FormState!.Value);
        Assert.Equal("live", source.QuerySelector("input")!.FormState!.Value);
    }

    [Fact]
    public void ACancelledWarmProjectionDoesNotReturnOrChangeTheSharedSnapshot() {
        var document = HtmlConversionDocument.Parse("<style>@media print { p { color: red; } }</style><p>Text</p>");
        document.CreateDocumentForConversion();
        var cancelled = new CancellationToken(true);
        Assert.ThrowsAny<OperationCanceledException>(() => document.CreateDocumentForConversion(cancelled));
        Assert.ThrowsAny<OperationCanceledException>(() => document.CreateDocumentForConversion(HtmlCssMediaContext.Print, cancelled));
        Assert.Equal("Text", document.CreateDocumentForConversion(HtmlCssMediaContext.Print).QuerySelector("p")!.TextContent);
        Assert.Contains("@media print", document.SourceHtml);
    }
}
