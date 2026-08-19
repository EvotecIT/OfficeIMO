using OfficeIMO.Drawing;
using OfficeIMO.Html;
using OfficeIMO.Rtf;
using Xunit;

namespace OfficeIMO.Tests.Rtf;

public class RtfHtmlOptionsTests {
    [Fact]
    public void HtmlToRtfOptions_Profiles_Expose_Bounded_Untrusted_Profile() {
        HtmlToRtfOptions defaultProfile = HtmlToRtfOptions.CreateOfficeIMOProfile();
        HtmlToRtfOptions untrustedProfile = HtmlToRtfOptions.CreateUntrustedHtmlProfile();

        Assert.Null(defaultProfile.MaxHtmlNodes);
        Assert.Null(defaultProfile.MaxHtmlDepth);
        Assert.NotNull(defaultProfile.UrlPolicy);
        Assert.Equal(10000, untrustedProfile.MaxHtmlNodes);
        Assert.Equal(64, untrustedProfile.MaxHtmlDepth);
        Assert.True(untrustedProfile.IgnoreInsignificantWhitespace);
        Assert.False(untrustedProfile.PreserveUnknownTagsAsText);
    }

    [Fact]
    public void HtmlToRtfOptions_Clone_Copies_Configuration_Without_Diagnostics() {
        var options = new HtmlToRtfOptions {
            BaseUri = new Uri("https://example.test/root/"),
            PreserveUnknownTagsAsText = true,
            IgnoreInsignificantWhitespace = false,
            UrlPolicy = HtmlUrlPolicy.CreateWebOnlyProfile(),
            MaxHtmlNodes = 33,
            MaxHtmlDepth = 44
        };

        HtmlToRtfOptions clone = options.Clone();

        Assert.NotSame(options, clone);
        Assert.Equal(options.BaseUri, clone.BaseUri);
        Assert.Equal(options.PreserveUnknownTagsAsText, clone.PreserveUnknownTagsAsText);
        Assert.Equal(options.IgnoreInsignificantWhitespace, clone.IgnoreInsignificantWhitespace);
        Assert.NotSame(options.UrlPolicy, clone.UrlPolicy);
        Assert.True(clone.UrlPolicy.RestrictUrlSchemes);
        Assert.Equal(options.MaxHtmlNodes, clone.MaxHtmlNodes);
        Assert.Equal(options.MaxHtmlDepth, clone.MaxHtmlDepth);
    }

    [Fact]
    public void RtfToHtmlOptions_Clone_Copies_Configuration() {
        var options = new RtfToHtmlOptions {
            Profile = OfficeHtmlConversionProfile.RtfPrintReview,
            Theme = OfficeVisualThemeKind.Report,
            IncludeDefaultStyles = true,
            FragmentOnly = false,
            IncludeMetadata = false,
            Title = "Clinical note",
            UrlPolicy = HtmlUrlPolicy.CreateHyperlinkProfile(),
            IncludeRoundTripMetadata = true,
            EmbedImagesAsDataUri = false,
            MaxEmbeddedImageBytes = 123,
            ImageSourceResolver = _ => "https://example.test/image.png",
            NewLine = "\n"
        };

        RtfToHtmlOptions clone = options.Clone();

        Assert.NotSame(options, clone);
        Assert.Equal(options.Profile, clone.Profile);
        Assert.Equal(options.Theme, clone.Theme);
        Assert.Equal(options.IncludeDefaultStyles, clone.IncludeDefaultStyles);
        Assert.Equal(options.FragmentOnly, clone.FragmentOnly);
        Assert.Equal(options.IncludeMetadata, clone.IncludeMetadata);
        Assert.Equal(options.Title, clone.Title);
        Assert.NotSame(options.UrlPolicy, clone.UrlPolicy);
        Assert.True(clone.UrlPolicy.DisallowFileUrls);
        Assert.Equal(options.IncludeRoundTripMetadata, clone.IncludeRoundTripMetadata);
        Assert.Equal(options.EmbedImagesAsDataUri, clone.EmbedImagesAsDataUri);
        Assert.Equal(options.MaxEmbeddedImageBytes, clone.MaxEmbeddedImageBytes);
        Assert.Same(options.ImageSourceResolver, clone.ImageSourceResolver);
        Assert.Equal(options.NewLine, clone.NewLine);
    }

    [Fact]
    public void RtfToHtmlOptions_PrintReviewProfile_UsesSharedPremiumDocumentShell() {
        RtfDocument document = RtfDocument.Create();
        document.AddParagraph("Premium RTF review");

        RtfToHtmlOptions options = RtfToHtmlOptions.CreatePrintReviewProfile(OfficeVisualThemeKind.GitHubLike);
        string html = document.ToHtml(options);

        Assert.Equal(OfficeHtmlConversionProfile.RtfPrintReview, options.Profile);
        Assert.False(options.FragmentOnly);
        Assert.Contains("class=\"officeimo-html officeimo-rtf-html\"", html, StringComparison.Ordinal);
        Assert.Contains("data-officeimo-profile=\"RtfPrintReview\"", html, StringComparison.Ordinal);
        Assert.Contains("--officeimo-accent:#0969DA", html, StringComparison.Ordinal);
        Assert.Contains("@media print", html, StringComparison.Ordinal);
        Assert.Contains("<p>Premium RTF review</p>", html, StringComparison.Ordinal);
    }

    [Fact]
    public void RtfPrintReviewProfile_RoundTripsStructuredHtmlThroughNativeRtf() {
        const string sourceHtml = "<h1>Review</h1><ul><li>Ready</li><li>Verified</li></ul><table><thead><tr><th>State</th><th>Count</th></tr></thead><tbody><tr><td>Open</td><td>2</td></tr></tbody></table>";
        HtmlToRtfResult import = HtmlConversionDocument.Parse(sourceHtml).ToRtfDocumentResult(HtmlToRtfOptions.CreateOfficeIMOProfile());

        RtfToHtmlResult export = import.Value.ToHtmlResult(RtfToHtmlOptions.CreatePrintReviewProfile(OfficeVisualThemeKind.Report));
        string html = export.RequireValue();

        Assert.Contains("<h1", html, StringComparison.Ordinal);
        Assert.Contains("Review", html, StringComparison.Ordinal);
        Assert.Contains("</h1>", html, StringComparison.Ordinal);
        Assert.Contains("<ul>", html, StringComparison.Ordinal);
        Assert.Contains("<li>Ready</li>", html, StringComparison.Ordinal);
        Assert.Contains("<table", html, StringComparison.Ordinal);
        Assert.Contains("<th", html, StringComparison.Ordinal);
        Assert.Contains("State", html, StringComparison.Ordinal);
        Assert.Contains("--officeimo-accent:#1D4ED8", html, StringComparison.Ordinal);
        Assert.DoesNotContain(import.Report.Diagnostics, diagnostic => diagnostic.Severity == HtmlDiagnosticSeverity.Error);
        Assert.DoesNotContain(export.Report.Diagnostics, diagnostic => diagnostic.Severity == HtmlDiagnosticSeverity.Error);
    }

    [Fact]
    public void RtfToHtmlOptions_RejectsNonRtfProfile() {
        Assert.Throws<ArgumentOutOfRangeException>(() =>
            new RtfToHtmlOptions().Profile = OfficeHtmlConversionProfile.WordSemanticDocument);
    }

    [Fact]
    public void Html_ToRtfDocument_MaxHtmlNodes_Stops_With_Diagnostic() {
        var options = new HtmlToRtfOptions {
            MaxHtmlNodes = 1
        };

        HtmlRtfConversionLimitException exception = Assert.Throws<HtmlRtfConversionLimitException>(() =>
            HtmlConversionDocument.Parse("<p>One</p><p>Two</p>").ToRtfDocument(options));

        Assert.Equal("HtmlNodeLimitExceeded", exception.Code);
        Assert.Equal("MaxHtmlNodes", exception.LimitSource);
        Assert.True(exception.Actual > exception.Limit);
        Assert.Equal(1, exception.Limit);
    }

    [Fact]
    public void Html_ToRtfDocument_MaxHtmlDepth_Stops_With_Diagnostic() {
        var options = new HtmlToRtfOptions {
            MaxHtmlDepth = 2
        };

        HtmlRtfConversionLimitException exception = Assert.Throws<HtmlRtfConversionLimitException>(() =>
            HtmlConversionDocument.Parse("<div><section><p>Too deep</p></section></div>").ToRtfDocument(options));

        Assert.Equal("HtmlDepthLimitExceeded", exception.Code);
        Assert.Equal("MaxHtmlDepth", exception.LimitSource);
        Assert.True(exception.Actual > exception.Limit);
        Assert.Equal(2, exception.Limit);
    }

    [Fact]
    public void Options_Can_Be_Reused_Without_Leaking_Diagnostics_Between_Results() {
        var options = new RtfToHtmlOptions();
        RtfDocument lossy = RtfDocument.Create();
        lossy.AddImage(RtfImageFormat.Png, new byte[] { 1, 2, 3 });
        RtfDocument clean = RtfDocument.Create();
        clean.AddParagraph("Clean");

        RtfToHtmlResult first = lossy.ToHtmlResult(options);
        RtfToHtmlResult second = clean.ToHtmlResult(options);

        Assert.NotEmpty(first.RtfDiagnostics);
        Assert.True(first.Report.HasLoss);
        Assert.Empty(second.RtfDiagnostics);
        Assert.False(second.Report.HasLoss);
    }
}
