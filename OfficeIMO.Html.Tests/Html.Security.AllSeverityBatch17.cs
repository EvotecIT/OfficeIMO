using OfficeIMO.Html;
using OfficeIMO.Html.Pdf;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class HtmlAllSeverityBatch17SecurityTests {
    [Fact]
    public void GenericRenderOptionsCannotWidenPdfHyperlinkSchemes() {
        int transformCalls = 0;
        HtmlUrlPolicy callerPolicy = HtmlUrlPolicy.CreateHyperlinkProfile();
        callerPolicy.AllowMailtoUrls = false;
        callerPolicy.AllowProtocolRelativeUrls = false;
        callerPolicy.AllowedUrlSchemes.Remove(Uri.UriSchemeHttp);
        callerPolicy.ResolvedUrlTransform = value => {
            transformCalls++;
            return value;
        };
        var generic = new HtmlRenderOptions {
            UrlPolicy = callerPolicy
        };

        var pdf = new HtmlPdfSaveOptions(generic);

        Assert.False(HtmlUrlPolicyEvaluator.IsAllowed("file:///private/secret.txt", pdf.UrlPolicy));
        Assert.False(HtmlUrlPolicyEvaluator.IsAllowed("data:text/html,secret", pdf.UrlPolicy));
        Assert.False(HtmlUrlPolicyEvaluator.IsAllowed("smb://server/share", pdf.UrlPolicy));
        Assert.False(HtmlUrlPolicyEvaluator.IsAllowed("http://example.test/report", pdf.UrlPolicy));
        Assert.False(HtmlUrlPolicyEvaluator.IsAllowed("mailto:security@example.test", pdf.UrlPolicy));
        Assert.False(HtmlUrlPolicyEvaluator.IsAllowed("//example.test/report", pdf.UrlPolicy));
        Assert.True(HtmlUrlPolicyEvaluator.IsAllowed("https://example.test/report", pdf.UrlPolicy));
        Assert.Equal(
            "https://example.test/report",
            HtmlUrlPolicyEvaluator.ResolveUrl("https://example.test/report", null, pdf.UrlPolicy));
        Assert.True(transformCalls > 0);
    }

    [Fact]
    public void PdfSpecificOptionSnapshotsRetainExplicitPdfPolicy() {
        var source = new HtmlPdfSaveOptions {
            UrlPolicy = HtmlUrlPolicy.CreateOfficeIMOProfile()
        };

        var copy = new HtmlPdfSaveOptions(source);

        Assert.False(copy.UrlPolicy.RestrictUrlSchemes);
    }
}
