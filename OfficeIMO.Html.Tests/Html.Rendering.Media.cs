using OfficeIMO.Drawing;
using OfficeIMO.Html;
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class HtmlRenderingTests {
    [Fact]
    public void HtmlRender_MediaLengthsUseTheActiveContinuousSurface() {
        const string html = "<style>"
            + ".target{color:#0000ff}"
            + "@media (max-width:300px){.target{color:#ff0000}}"
            + "@media (min-width:350px) and (max-width:450px){.target{color:#008000}}"
            + "</style><p class='target'>Media marker</p>";

        HtmlRenderDocument medium = HtmlRenderTestDriver.Render(HtmlConversionDocument.Parse(html), new HtmlRenderOptions {
            ViewportWidth = 400D,
            ViewportHeight = 200D,
            Margins = HtmlRenderMargins.All(0D)
        });
        HtmlRenderDocument wide = HtmlRenderTestDriver.Render(HtmlConversionDocument.Parse(html), new HtmlRenderOptions {
            ViewportWidth = 800D,
            ViewportHeight = 600D,
            Margins = HtmlRenderMargins.All(0D)
        });

        HtmlRenderText mediumText = Assert.Single(medium.Pages[0].Visuals.OfType<HtmlRenderText>(), text => text.Text.Contains("Media", StringComparison.Ordinal));
        HtmlRenderText wideText = Assert.Single(wide.Pages[0].Visuals.OfType<HtmlRenderText>(), text => text.Text.Contains("Media", StringComparison.Ordinal));
        Assert.Equal(OfficeColor.FromRgb(0, 128, 0), mediumText.Color);
        Assert.Equal(OfficeColor.Blue, wideText.Color);
        Assert.False(HtmlComputedStyleEngine.IsApplicableMedia("(max-width:1px)", HtmlCssMediaContext.Screen, 400D, 200D));
        Assert.True(HtmlComputedStyleEngine.IsApplicableMedia("(orientation:landscape)", HtmlCssMediaContext.Screen, 400D, 200D));
    }

    [Fact]
    public void HtmlRender_MediaLengthsUseTheActivePagedSurface() {
        const string html = "<style>.target{color:#0000ff}@media print and (max-width:300px){.target{color:#ff0000}}</style><p class='target'>Paged media</p>";
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(HtmlConversionDocument.Parse(html), new HtmlRenderOptions {
            Mode = HtmlRenderMode.Paged,
            PageSize = new OfficePageSize(4D, 3D),
            HonorCssPageRules = false,
            Margins = HtmlRenderMargins.All(0D)
        });

        HtmlRenderText text = Assert.Single(rendered.Pages[0].Visuals.OfType<HtmlRenderText>(), item => item.Text.Contains("Paged", StringComparison.Ordinal));
        Assert.Equal(OfficeColor.Blue, text.Color);
    }

    [Fact]
    public void HtmlRender_MediaFeaturesSelectDeterministicStaticPreferences() {
        const string html = """
            <style>
              .target { color:#0000ff }
              @media (prefers-color-scheme:dark) and (prefers-reduced-motion:reduce) and (pointer:none) and (hover:none) {
                .target { color:#ff0000 }
              }
            </style>
            <p class="target">Static preferences</p>
            """;
        var options = new HtmlRenderOptions {
            ViewportWidth = 400D,
            ViewportHeight = 200D,
            Margins = HtmlRenderMargins.All(0D),
            MediaFeatures = new HtmlRenderMediaFeatures {
                PreferredColorScheme = HtmlPreferredColorScheme.Dark,
                ReducedMotion = HtmlReducedMotionPreference.Reduce
            }
        };

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(HtmlConversionDocument.Parse(html), options);

        HtmlRenderText text = Assert.Single(
            rendered.Pages[0].Visuals.OfType<HtmlRenderText>(),
            item => item.Text.Contains("Static", StringComparison.Ordinal));
        Assert.Equal(OfficeColor.Red, text.Color);
        Assert.True(HtmlComputedStyleEngine.IsApplicableMedia(
            "(min-resolution:1dppx) and (max-resolution:96dpi) and (scripting:none) and (update:none)",
            HtmlCssMediaContext.Screen,
            400D,
            200D,
            options.MediaFeatures));
        Assert.False(HtmlComputedStyleEngine.IsApplicableMedia(
            "(pointer:fine)",
            HtmlCssMediaContext.Screen,
            400D,
            200D,
            options.MediaFeatures));
        Assert.True(HtmlComputedStyleEngine.IsApplicableMedia(
            "(prefers-color-scheme) and (prefers-reduced-motion)",
            HtmlCssMediaContext.Screen,
            400D,
            200D,
            options.MediaFeatures));
        Assert.False(HtmlComputedStyleEngine.IsApplicableMedia(
            "(prefers-reduced-motion)",
            HtmlCssMediaContext.Screen,
            400D,
            200D,
            new HtmlRenderMediaFeatures()));
    }

    [Fact]
    public void HtmlRender_MediaFeaturesHonorBooleanPointerAndHoverQueries() {
        var interactive = new HtmlRenderMediaFeatures {
            Pointer = HtmlPointerCapability.Fine,
            AnyPointer = HtmlPointerCapability.Coarse,
            Hover = HtmlHoverCapability.Hover,
            AnyHover = HtmlHoverCapability.Hover
        };

        Assert.True(HtmlComputedStyleEngine.IsApplicableMedia(
            "(pointer) and (any-pointer) and (hover) and (any-hover) and (resolution)",
            HtmlCssMediaContext.Screen,
            400D,
            200D,
            interactive));
        Assert.False(HtmlComputedStyleEngine.IsApplicableMedia(
            "(pointer), (any-pointer), (hover), (any-hover)",
            HtmlCssMediaContext.Screen,
            400D,
            200D,
            new HtmlRenderMediaFeatures()));
    }

    [Theory]
    [InlineData("(prefers-color-scheme:dark) or (prefers-reduced-motion:reduce)", true)]
    [InlineData("(prefers-reduced-motion:reduce) or (prefers-color-scheme:dark)", true)]
    [InlineData("(prefers-color-scheme:light) or (prefers-reduced-motion:reduce)", false)]
    [InlineData("(prefers-color-scheme:dark) and (prefers-reduced-motion:reduce)", false)]
    [InlineData("or (prefers-color-scheme:dark)", false)]
    [InlineData("(prefers-color-scheme:dark) or", false)]
    [InlineData("(max-resolution:1e2dpi)", true)]
    [InlineData("(resolution:9.6E1dpi)", true)]
    [InlineData("(min-resolution:9.7e1dpi)", false)]
    [InlineData("(color) and not (hover)", true)]
    [InlineData("(color) and not (hover:none)", false)]
    public void HtmlRender_MediaQueriesHonorLogicalOrAndResolutionExponents(
        string mediaQuery,
        bool expected) {
        var features = new HtmlRenderMediaFeatures {
            PreferredColorScheme = HtmlPreferredColorScheme.Dark
        };

        Assert.Equal(
            expected,
            HtmlComputedStyleEngine.IsApplicableMedia(
                mediaQuery,
                HtmlCssMediaContext.Screen,
                400D,
                200D,
                features));
    }

    [Theory]
    [InlineData("(max-width:4e2px)", true)]
    [InlineData("(width:4.0E+2px)", true)]
    [InlineData("(min-height:2e2px)", true)]
    [InlineData("(max-width:3.99e2px)", false)]
    [InlineData("(width:4e+px)", false)]
    public void HtmlRender_MediaLengthsHonorExponentNotation(string mediaQuery, bool expected) {
        Assert.Equal(
            expected,
            HtmlComputedStyleEngine.IsApplicableMedia(
                mediaQuery,
                HtmlCssMediaContext.Screen,
                400D,
                200D));
    }

    [Fact]
    public void HtmlRender_AdditionalStylesheetsParticipateInTheBoundedAuthorCascade() {
        var options = new HtmlRenderOptions {
            ViewportWidth = 400D,
            ViewportHeight = 200D,
            Margins = HtmlRenderMargins.All(0D)
        };
        options.AdditionalStylesheets.Add(".target { color:#008000; width:clamp(40px, 25%, 80px) }");

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(
            HtmlConversionDocument.Parse("<style>.target{color:#0000ff}</style><p class='target'>Caller style</p>"),
            options);

        HtmlRenderText text = Assert.Single(
            rendered.Pages[0].Visuals.OfType<HtmlRenderText>(),
            item => item.Text.Contains("Caller", StringComparison.Ordinal));
        Assert.Equal(OfficeColor.FromRgb(0, 128, 0), text.Color);
        HtmlRenderOptions clone = options.Clone();
        Assert.Single(clone.AdditionalStylesheets);
        Assert.NotSame(options.MediaFeatures, clone.MediaFeatures);
    }

    [Fact]
    public void HtmlRender_AdditionalStylesheetsFollowBodyPositionedDocumentStyles() {
        var options = new HtmlRenderOptions {
            ViewportWidth = 400D,
            ViewportHeight = 200D,
            Margins = HtmlRenderMargins.All(0D)
        };
        options.AdditionalStylesheets.Add(".target { color:#008000 }");

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(
            HtmlConversionDocument.Parse("<p class='target'>Caller last</p><style>.target{color:#0000ff}</style>"),
            options);

        HtmlRenderText text = Assert.Single(
            rendered.Pages[0].Visuals.OfType<HtmlRenderText>(),
            item => item.Text.Contains("Caller last", StringComparison.Ordinal));
        Assert.Equal(OfficeColor.FromRgb(0, 128, 0), text.Color);
    }
}
