using System.Text;
using OfficeIMO.Drawing;
using OfficeIMO.Html;
using OfficeIMO.Html.Pdf;
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class HtmlRenderingTests {
    [Fact]
    public void HtmlRadialGradient_UsesOverflowSafeCornerDistance() {
        Assert.True(HtmlCssRadialGradientParser.TryParse(
            "radial-gradient(circle farthest-corner at 1e308px 1e308px,red,blue)",
            maximumStops: 8,
            out HtmlCssRadialGradientDefinition? definition,
            out bool stopLimitExceeded));
        Assert.False(stopLimitExceeded);
        Assert.NotNull(definition);
        Assert.True(definition!.TryResolve(40D, 20D, 16D, 16D, out OfficeRadialGradient? gradient));
        Assert.NotNull(gradient);
        Assert.False(double.IsNaN(gradient!.EndRadiusX));
        Assert.False(double.IsInfinity(gradient.EndRadiusX));
        Assert.False(double.IsNaN(gradient.EndRadiusY));
        Assert.False(double.IsInfinity(gradient.EndRadiusY));
    }

    [Fact]
    public void HtmlPdf_SkipsLinkedWhitespaceOnlySvgText() {
        string svg = Convert.ToBase64String(Encoding.UTF8.GetBytes(
            "<svg xmlns='http://www.w3.org/2000/svg' width='20' height='10'><a href='https://example.test'><text x='1' y='8'>   </text></a></svg>"));

        byte[] pdf = HtmlConversionDocument.Parse(
            "<img src='data:image/svg+xml;base64," + svg + "'><p>AfterWhitespaceVector</p>").ToPdf();

        Assert.Contains("AfterWhitespaceVector", OfficeIMO.Pdf.PdfReadDocument.Open(pdf).ExtractText(), StringComparison.Ordinal);
    }

    [Fact]
    public void HtmlMargins_NegativeCollapsedPairsDoNotInflateFollowingFlow() {
        const string html = "<div style='margin:0'>"
            + "<div id='first' style='height:10px;margin:0 0 -100px;background:red'></div>"
            + "<div id='second' style='height:10px;margin:-100px 0 0;background:blue'></div>"
            + "<div id='tail' style='height:10px;margin:0;background:lime'></div>"
            + "</div>";

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, new HtmlRenderOptions {
            ViewportWidth = 40D,
            ViewportHeight = 40D,
            Margins = HtmlRenderMargins.All(0D),
            BackgroundColor = OfficeColor.Transparent
        });

        HtmlRenderShape tail = Assert.Single(rendered.Pages[0].Visuals.OfType<HtmlRenderShape>(), shape => shape.Source == "div#tail");
        Assert.InRange(tail.Y, 9.9D, 10.1D);
    }

    [Fact]
    public void HtmlImages_RejectAuthoredDimensionsBeforeBuildingOversizedVisuals() {
        var options = new HtmlRenderOptions {
            ViewportWidth = 100D,
            Margins = HtmlRenderMargins.All(0D),
            MaxSurfaceWidth = 256,
            MaxSurfaceHeight = 256
        };

        InvalidOperationException exception = Assert.Throws<InvalidOperationException>(() =>
            HtmlRenderTestDriver.Render("<img style='display:block;width:1e308px;height:20px'>", options));

        Assert.Contains("replaced content", exception.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void HtmlColumns_RejectGeneratedColumnsDuringPlanConstruction() {
        var html = new StringBuilder("<div style='height:1px;column-count:1;column-fill:auto;margin:0'>");
        for (int index = 0; index < 100; index++) html.Append("<div style='height:1px;margin:0'>x</div>");
        html.Append("</div>");
        var options = new HtmlRenderOptions {
            ViewportWidth = 100D,
            Margins = HtmlRenderMargins.All(0D),
            MaxColumnCount = 2
        };

        HtmlDomLimitException exception = Assert.Throws<HtmlDomLimitException>(() => HtmlRenderTestDriver.Render(html.ToString(), options));

        Assert.Equal(HtmlRenderDiagnosticCodes.MultiColumnLimitExceeded, exception.Code);
        Assert.Equal(3L, exception.Actual);
    }

    [Fact]
    public void HtmlTransforms_RejectFiniteInputsWhoseCompositionOverflows() {
        const string html = "<div id='overflowed' style='width:10px;height:10px;margin:0;background:red;transform:scale(1e308) scale(1e308)'></div>";

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, new HtmlRenderOptions {
            ViewportWidth = 40D,
            Margins = HtmlRenderMargins.All(0D)
        });

        Assert.Contains(rendered.Diagnostics, diagnostic =>
            diagnostic.Code == HtmlRenderDiagnosticCodes.TransformValueUnsupported
            && diagnostic.Source == "div#overflowed");
        Assert.DoesNotContain(rendered.Pages[0].Visuals, visual => visual is HtmlRenderEffectGroup group && group.Source == "div#overflowed");
    }

    [Theory]
    [InlineData("rotate(1e308deg)")]
    [InlineData("rotate(1e308grad)")]
    [InlineData("rotate(1e308rad)")]
    [InlineData("rotate(1e308turn)")]
    [InlineData("skew(1e308deg,-1e308deg)")]
    public void HtmlTransforms_NormalizeHugeFiniteAnglesBeforeMatrixConstruction(string transformValue) {
        string html = "<div id='large-angle' style='width:10px;height:10px;margin:0;background:red;transform-origin:0 0;transform:"
            + transformValue + "'></div>";

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, new HtmlRenderOptions {
            ViewportWidth = 40D,
            Margins = HtmlRenderMargins.All(0D)
        });

        HtmlRenderEffectGroup group = Assert.Single(
            EnumerateRenderVisuals(rendered.Pages[0].Visuals).OfType<HtmlRenderEffectGroup>(),
            visual => visual.Source == "div#large-angle");
        Assert.DoesNotContain(rendered.Diagnostics, diagnostic => diagnostic.Code == HtmlRenderDiagnosticCodes.TransformValueUnsupported);
        Assert.All(new[] { group.Transform.M11, group.Transform.M12, group.Transform.M21, group.Transform.M22 },
            value => Assert.False(double.IsNaN(value) || double.IsInfinity(value)));
    }

    [Theory]
    [InlineData("rotate(1e308deg)")]
    [InlineData("rotate(1e308grad)")]
    [InlineData("rotate(1e308rad)")]
    [InlineData("rotate(1e308turn)")]
    [InlineData("skew(1e308deg,-1e308deg)")]
    public void HtmlSupports_AcceptsHugeFiniteTransformAnglesWithoutThrowing(string transformValue) {
        string html = "<style>@supports (transform:" + transformValue + "){#supported{display:none}}</style>"
            + "<p id='supported'>SupportsLargeAngle</p>";
        var document = HtmlDocumentParser.ParseDocument(html);
        var element = document.QuerySelector("#supported")!;

        IReadOnlyDictionary<AngleSharp.Dom.IElement, HtmlComputedStyle> styles = HtmlComputedStyleEngine.Compute(document);

        Assert.Equal("none", styles[element].GetValue("display"));
    }

    [Fact]
    public void HtmlTransforms_RejectOverflowingAbsoluteOriginsBeforeTranslationConstruction() {
        bool parsed = HtmlCssTransformParser.TryParse(
            "rotate(1deg)",
            "1e308px 1e308px",
            1e308D,
            1e308D,
            10D,
            10D,
            16D,
            16D,
            out _,
            out string detail);

        Assert.False(parsed);
        Assert.Contains("transform-origin", detail, StringComparison.Ordinal);
    }

    [Theory]
    [InlineData("linear-gradient")]
    [InlineData("radial-gradient")]
    public void HtmlGradients_StopBeforeMaterializingUnboundedArgumentLists(string function) {
        string background = function + "(" + string.Join(",", Enumerable.Repeat("red", 10_000)) + ")";
        string html = "<div style='width:10px;height:10px;background:" + background + "'></div>";

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, new HtmlRenderOptions {
            ViewportWidth = 20D,
            Margins = HtmlRenderMargins.All(0D),
            MaxGradientStops = 8
        });

        Assert.Contains(rendered.Diagnostics, diagnostic => diagnostic.Code == HtmlRenderDiagnosticCodes.GradientStopLimitExceeded);
    }

    [Fact]
    public void HtmlFlex_NormalizesHugeGrowFactorsWithoutNaN() {
        const string html = "<div style='display:flex;width:300px;height:10px;margin:0'>"
            + "<div id='large' style='flex-grow:1e308;flex-basis:0;height:10px;background:red'></div>"
            + "<div id='half' style='flex-grow:5e307;flex-basis:0;height:10px;background:blue'></div>"
            + "</div>";

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, new HtmlRenderOptions {
            ViewportWidth = 300D,
            Margins = HtmlRenderMargins.All(0D)
        });
        HtmlRenderShape large = Assert.Single(rendered.Pages[0].Visuals.OfType<HtmlRenderShape>(), shape => shape.Source == "div#large");
        HtmlRenderShape half = Assert.Single(rendered.Pages[0].Visuals.OfType<HtmlRenderShape>(), shape => shape.Source == "div#half");

        Assert.Equal(200D, large.Width, 3);
        Assert.Equal(100D, half.Width, 3);
        Assert.False(double.IsNaN(large.Width));
        Assert.False(double.IsNaN(half.Width));
    }

    [Fact]
    public void HtmlTable_HandlesLongHeaderPrefixesWithOneSuffixPass() {
        var html = new StringBuilder("<table style='border-collapse:collapse;margin:0'>");
        for (int index = 0; index < 1_500; index++) html.Append("<tr><th>H</th></tr>");
        html.Append("<tr><td id='body-cell'>Body</td></tr></table>");

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html.ToString(), new HtmlRenderOptions {
            ViewportWidth = 100D,
            Margins = HtmlRenderMargins.All(0D),
            MaxTableRows = 2_000,
            MaxSurfaceHeight = 100_000
        });

        Assert.Contains("Body", rendered.Text, StringComparison.Ordinal);
    }

    [Fact]
    public void HtmlResources_PreloadAsStyleDoesNotBecomeAnActiveStylesheet() {
        const string html = "<link rel='preload' as='style' href='https://example.test/preloaded.css'>"
            + "<link rel='preload stylesheet' as='style' href='https://example.test/active.css'>";

        HtmlResourceManifest manifest = HtmlResourcePipeline.BuildManifest(html, new HtmlResourcePipelineOptions {
            UrlPolicy = HtmlUrlPolicy.CreateWebOnlyProfile()
        });

        Assert.Contains(manifest.Resources, resource => resource.Source == "https://example.test/preloaded.css" && resource.Kind == HtmlResourceKind.Other);
        Assert.Contains(manifest.Resources, resource => resource.Source == "https://example.test/active.css" && resource.Kind == HtmlResourceKind.Stylesheet);
    }
}
