using OfficeIMO.Html;
using OfficeIMO.Rtf.Html;
using Xunit;

namespace OfficeIMO.Tests.Rtf;

public class RtfHtmlOptionsTests {
    [Fact]
    public void RtfHtmlReadOptions_Profiles_Expose_Bounded_Untrusted_Profile() {
        RtfHtmlReadOptions defaultProfile = RtfHtmlReadOptions.CreateOfficeIMOProfile();
        RtfHtmlReadOptions untrustedProfile = RtfHtmlReadOptions.CreateUntrustedHtmlProfile();

        Assert.Null(defaultProfile.MaxHtmlNodes);
        Assert.Null(defaultProfile.MaxHtmlDepth);
        Assert.NotNull(defaultProfile.UrlPolicy);
        Assert.Equal(10000, untrustedProfile.MaxHtmlNodes);
        Assert.Equal(64, untrustedProfile.MaxHtmlDepth);
        Assert.True(untrustedProfile.IgnoreInsignificantWhitespace);
        Assert.False(untrustedProfile.PreserveUnknownTagsAsText);
    }

    [Fact]
    public void RtfHtmlReadOptions_Clone_Copies_Configuration_Without_Diagnostics() {
        Action<RtfHtmlConversionDiagnostic> handler = _ => { };
        var options = new RtfHtmlReadOptions {
            BaseUri = new Uri("https://example.test/root/"),
            PreserveUnknownTagsAsText = true,
            IgnoreInsignificantWhitespace = false,
            UrlPolicy = HtmlUrlPolicy.CreateWebOnlyProfile(),
            MaxHtmlNodes = 33,
            MaxHtmlDepth = 44,
            DiagnosticHandler = handler
        };
        options.Diagnostics.Add(new RtfHtmlConversionDiagnostic("Existing", "Existing diagnostic"));

        RtfHtmlReadOptions clone = options.Clone();

        Assert.NotSame(options, clone);
        Assert.Equal(options.BaseUri, clone.BaseUri);
        Assert.Equal(options.PreserveUnknownTagsAsText, clone.PreserveUnknownTagsAsText);
        Assert.Equal(options.IgnoreInsignificantWhitespace, clone.IgnoreInsignificantWhitespace);
        Assert.NotSame(options.UrlPolicy, clone.UrlPolicy);
        Assert.True(clone.UrlPolicy.RestrictUrlSchemes);
        Assert.Equal(options.MaxHtmlNodes, clone.MaxHtmlNodes);
        Assert.Equal(options.MaxHtmlDepth, clone.MaxHtmlDepth);
        Assert.Same(handler, clone.DiagnosticHandler);
        Assert.Empty(clone.Diagnostics);
    }

    [Fact]
    public void RtfHtmlSaveOptions_Clone_Copies_Configuration() {
        var options = new RtfHtmlSaveOptions {
            FragmentOnly = false,
            IncludeMetadata = false,
            Title = "Clinical note",
            EmbedImagesAsDataUri = false,
            NewLine = "\n"
        };

        RtfHtmlSaveOptions clone = options.Clone();

        Assert.NotSame(options, clone);
        Assert.Equal(options.FragmentOnly, clone.FragmentOnly);
        Assert.Equal(options.IncludeMetadata, clone.IncludeMetadata);
        Assert.Equal(options.Title, clone.Title);
        Assert.Equal(options.EmbedImagesAsDataUri, clone.EmbedImagesAsDataUri);
        Assert.Equal(options.NewLine, clone.NewLine);
    }

    [Fact]
    public void Html_ToRtfDocument_MaxHtmlNodes_Stops_With_Diagnostic() {
        var callbackDiagnostics = new List<RtfHtmlConversionDiagnostic>();
        var options = new RtfHtmlReadOptions {
            MaxHtmlNodes = 1,
            DiagnosticHandler = callbackDiagnostics.Add
        };

        RtfHtmlConversionLimitException exception = Assert.Throws<RtfHtmlConversionLimitException>(() =>
            "<p>One</p><p>Two</p>".LoadFromHtml(options));

        Assert.Equal("HtmlNodeLimitExceeded", exception.Code);
        Assert.Equal("MaxHtmlNodes", exception.LimitSource);
        Assert.True(exception.Actual > exception.Limit);
        Assert.Equal(1, exception.Limit);
        RtfHtmlConversionDiagnostic diagnostic = Assert.Single(options.Diagnostics);
        Assert.Same(diagnostic, Assert.Single(callbackDiagnostics));
        Assert.Equal(RtfHtmlConversionDiagnosticSeverity.Error, diagnostic.Severity);
        Assert.Equal("HtmlNodeLimitExceeded", diagnostic.Code);
        Assert.Equal("MaxHtmlNodes", diagnostic.Source);
    }

    [Fact]
    public void Html_ToRtfDocument_MaxHtmlDepth_Stops_With_Diagnostic() {
        var options = new RtfHtmlReadOptions {
            MaxHtmlDepth = 2
        };

        RtfHtmlConversionLimitException exception = Assert.Throws<RtfHtmlConversionLimitException>(() =>
            "<div><section><p>Too deep</p></section></div>".LoadFromHtml(options));

        Assert.Equal("HtmlDepthLimitExceeded", exception.Code);
        Assert.Equal("MaxHtmlDepth", exception.LimitSource);
        Assert.True(exception.Actual > exception.Limit);
        Assert.Equal(2, exception.Limit);
        RtfHtmlConversionDiagnostic diagnostic = Assert.Single(options.Diagnostics);
        Assert.Equal(RtfHtmlConversionDiagnosticSeverity.Error, diagnostic.Severity);
        Assert.Equal("HtmlDepthLimitExceeded", diagnostic.Code);
        Assert.Equal("MaxHtmlDepth", diagnostic.Source);
    }
}
