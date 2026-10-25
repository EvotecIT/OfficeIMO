using OfficeIMO.Drawing;
using OfficeIMO.Html;
using OfficeIMO.TestAssets;
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class HtmlRenderingTests {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task HtmlDefaultFontMetrics_RawEstimatesAreTypedAndStrictPolicyRejectsThem(bool asynchronous) {
        const string html = "<style>@page{size:240px 180px;margin:24px;"
            + "@top-left{content:'HEADER';font:12px Unselected}}body,p{margin:0}"
            + "#subject{font:700 16px/20px Unselected}</style><p id='subject'>VISIBLE</p>";
        HtmlConversionDocument source = HtmlConversionDocument.Parse(html);
        HtmlRenderOptions options = DefaultMetricOptions();
        HtmlRenderDocument rendered = asynchronous
            ? await HtmlRenderEngine.RenderAsync(source, options)
            : HtmlRenderEngine.Render(source, options);
        HtmlDiagnostic[] estimates = rendered.Diagnostics
            .Where(item => item.Code == HtmlRenderDiagnosticCodes.FontMetricsEstimated).ToArray();

        Assert.Contains(estimates, item => item.Source == "p#subject" && item.Detail!.Contains("weight=700", StringComparison.Ordinal));
        Assert.Contains(estimates, item => item.Source == "@page @top-left");
        Assert.All(estimates, item => {
            Assert.Equal(HtmlDiagnosticSeverity.Warning, item.Severity);
            Assert.Equal(OfficeConversionLossKind.Approximation, item.LossKind);
            Assert.Contains("estimator=OfficeTextMeasurer;axes=width;family=Unselected", item.Detail, StringComparison.Ordinal);
        });
        Assert.True(rendered.HasLoss);
        Assert.Throws<HtmlConversionException>(() => rendered.RequireNoLoss());
        options.FidelityPolicy = HtmlRenderFidelityPolicy.RequireNoLoss;
        if (asynchronous) await Assert.ThrowsAsync<HtmlConversionException>(() => HtmlRenderEngine.RenderAsync(source, options));
        else Assert.Throws<HtmlConversionException>(() => HtmlRenderEngine.Render(source, options));
    }

    [Fact]
    public void HtmlDefaultFontMetrics_RegisteredProgramKeepsVisibleBodyAndMarginLossless() {
        const string html = "<style>@page{size:240px 180px;margin:24px;"
            + "@top-left{content:'HEADER';font:12px Unselected}}body,p{margin:0}"
            + "#subject{font:16px Unselected}</style><p id='subject'>VISIBLE</p>";
        HtmlRenderOptions options = DefaultMetricOptions(withFont: true);
        HtmlRenderDocument rendered = HtmlRenderEngine.Render(HtmlConversionDocument.Parse(html), options);
        HtmlRenderText body = Assert.Single(rendered.Pages.SelectMany(page => EnumerateRenderVisuals(page.Scene))
            .OfType<HtmlRenderText>(), item => item.Source == "p#subject");
        Assert.True(rendered.Fonts.TryResolveFaceForText(body.Text, body.Font.FamilyName,
            body.FontDescriptor, body.Font.Size, out OfficeFontFace? face));

        Assert.Equal(face!.Program.Measure(body.Text, body.Font.Size), body.TextAdvanceWidth!.Value, 6);
        Assert.DoesNotContain(rendered.Diagnostics, item => item.Code == HtmlRenderDiagnosticCodes.FontMetricsEstimated);
        rendered.RequireNoLoss();
    }

    [Theory]
    [InlineData("display:none")]
    [InlineData("visibility:hidden")]
    [InlineData("font-size:0")]
    [InlineData("opacity:0")]
    [InlineData("color:transparent")]
    public void HtmlDefaultFontMetrics_NonPaintedRequestsDoNotInventEstimateLoss(string hiddenStyle) {
        string html = "<style>body,p{margin:0;font:16px Unselected}</style>"
            + "<p id='hidden' style='font-family:Unregistered;" + hiddenStyle + "'>HIDDEN</p><p>VISIBLE</p>";
        HtmlRenderOptions options = DefaultMetricOptions(withFont: true);
        options.Mode = HtmlRenderMode.Continuous;
        HtmlRenderDocument rendered = HtmlRenderEngine.Render(HtmlConversionDocument.Parse(html), options);

        Assert.DoesNotContain(rendered.Diagnostics, item => item.Code == HtmlRenderDiagnosticCodes.FontMetricsEstimated);
        rendered.RequireNoLoss();
    }

    [Fact]
    public void HtmlDefaultFontMetrics_UnusedRunningCaptureDoesNotAttributeDiscardedText() {
        const string html = "<style>@page{size:240px 180px;margin:24px}body,p{margin:0;font:16px Unselected}"
            + "#unused{position:running(unused);font-family:Unregistered}</style>"
            + "<div id='unused'>DISCARDED</div><p>VISIBLE</p>";
        HtmlRenderDocument rendered = HtmlRenderEngine.Render(HtmlConversionDocument.Parse(html), DefaultMetricOptions(withFont: true));

        Assert.Equal("VISIBLE", rendered.Text.Trim());
        Assert.DoesNotContain(rendered.Diagnostics, item => item.Code == HtmlRenderDiagnosticCodes.FontMetricsEstimated);
        rendered.RequireNoLoss();
    }

    [Theory]
    [InlineData("")]
    [InlineData("<div></div>")]
    [InlineData("<p>\u200d\ufe0f</p>")]
    public void HtmlDefaultFontMetrics_EmptyAndControlOnlyContentDoesNotInventLoss(string html) {
        HtmlRenderDocument rendered = HtmlRenderEngine.Render(HtmlConversionDocument.Parse(html), DefaultMetricOptions());

        Assert.DoesNotContain(rendered.Diagnostics, item => item.Code == HtmlRenderDiagnosticCodes.FontMetricsEstimated);
    }

    [Theory]
    [InlineData("", true)]
    [InlineData("display:none", false)]
    [InlineData("visibility:hidden", false)]
    public void HtmlDefaultFontMetrics_IframeOnlyEstimatesFollowRetainedChildPaint(string frameStyle, bool expectedLoss) {
        string child = System.Net.WebUtility.HtmlEncode(
            "<p id='nested' style='margin:0;font:16px Unselected'>IFRAME</p>");
        HtmlConversionDocument source = HtmlConversionDocument.Parse(
            "<style>body{margin:0}</style><iframe style='width:200px;height:80px;border:0;"
            + frameStyle + "' srcdoc=\"" + child + "\"></iframe>");
        HtmlRenderOptions options = DefaultMetricOptions();
        options.Mode = HtmlRenderMode.Continuous;
        HtmlRenderDocument rendered = HtmlRenderEngine.Render(source, options);

        Assert.Equal(expectedLoss, rendered.HasLoss);
        if (expectedLoss) {
            HtmlDiagnostic estimate = Assert.Single(rendered.Diagnostics,
                item => item.Code == HtmlRenderDiagnosticCodes.FontMetricsEstimated);
            Assert.Equal("p#nested", estimate.Source);
            Assert.Equal(OfficeConversionLossKind.Approximation, estimate.LossKind);
            Assert.Contains("IFRAME", rendered.Text, StringComparison.Ordinal);
            Assert.Throws<HtmlConversionException>(() => rendered.RequireNoLoss());
        } else {
            Assert.DoesNotContain(rendered.Diagnostics, item => item.Code == HtmlRenderDiagnosticCodes.FontMetricsEstimated);
            rendered.RequireNoLoss();
        }
    }

    private static HtmlRenderOptions DefaultMetricOptions(bool withFont = false) {
        var options = new HtmlRenderOptions {
            Mode = HtmlRenderMode.Paged,
            DefaultFontFamily = "Unselected",
            AllowSystemFontFallback = false,
            Margins = HtmlRenderMargins.All(0D)
        };
        if (withFont) options.Fonts.Add("Unselected", ManagedTextShapingTestAssets.CreateFontWithLineBoxMetrics(
            800, -200, 100, 200, Enumerable.Range(32, 95).ToArray()));
        return options;
    }
}
