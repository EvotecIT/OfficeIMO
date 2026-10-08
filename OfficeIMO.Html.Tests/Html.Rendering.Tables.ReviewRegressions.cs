using OfficeIMO.Html;
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class HtmlRenderingTests {
    [Theory]
    [InlineData("inline", "width")]
    [InlineData("inline", "min-width")]
    [InlineData("inline", "max-width")]
    [InlineData("inline", "height")]
    [InlineData("inline", "min-height")]
    [InlineData("inline", "max-height")]
    [InlineData("contents", "width")]
    [InlineData("contents", "min-width")]
    [InlineData("contents", "max-width")]
    [InlineData("contents", "height")]
    [InlineData("contents", "min-height")]
    [InlineData("contents", "max-height")]
    public void HtmlTableGeometry_InapplicableIntrinsicDimensionsDoNotRejectNoLoss(string display, string property) {
        string html = TableGeometrySource("<div><span style='display:" + display + ";" + property + ":max-content'>Text</span></div>");
        var options = TableGeometryOptions();
        options.FidelityPolicy = HtmlRenderFidelityPolicy.RequireNoLoss;
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, options);

        Assert.DoesNotContain(rendered.Diagnostics, diagnostic => diagnostic.Code == HtmlRenderDiagnosticCodes.IntrinsicSizeUnsupported);
        Assert.False(rendered.HasLoss);
        rendered.RequireNoLoss();
        Assert.Contains(rendered.Pages[0].Visuals.OfType<HtmlRenderText>(), text => text.Text == "Text");
    }

    [Theory]
    [InlineData("width")]
    [InlineData("min-width")]
    [InlineData("max-width")]
    [InlineData("height")]
    [InlineData("min-height")]
    [InlineData("max-height")]
    public void HtmlTableGeometry_InlineReplacedIntrinsicDimensionsRetainLoss(string property) {
        const string pixel = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNgYAAAAAMAASsJTYQAAAAASUVORK5CYII=";
        string html = TableGeometrySource("<img style='display:inline;" + property + ":max-content' src='data:image/png;base64," + pixel + "'>");
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, TableGeometryOptions());

        HtmlDiagnostic diagnostic = Assert.Single(rendered.Diagnostics, item => item.Code == HtmlRenderDiagnosticCodes.IntrinsicSizeUnsupported);
        Assert.Contains(property + "=max-content", diagnostic.Detail);
        Assert.Throws<HtmlConversionException>(() => rendered.RequireNoLoss());
    }

    [Theory]
    [InlineData("display:inline-block", "block")]
    [InlineData("position:absolute", "block")]
    [InlineData("float:left", "block")]
    [InlineData("", "flex")]
    [InlineData("", "grid")]
    public void HtmlTableGeometry_BlockifiedInlineIntrinsicDimensionsRetainLoss(string inlineStyle, string parentDisplay) {
        string html = TableGeometrySource("<div style='display:" + parentDisplay + "'><span style='width:max-content;"
            + inlineStyle + "'>Text</span></div>");
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, TableGeometryOptions());

        Assert.Single(rendered.Diagnostics, item => item.Code == HtmlRenderDiagnosticCodes.IntrinsicSizeUnsupported);
        Assert.Throws<HtmlConversionException>(() => rendered.RequireNoLoss());
    }

    [Theory]
    [InlineData("flex")]
    [InlineData("grid")]
    public void HtmlTableGeometry_ContentsDescendantsRetainIntrinsicLossWhenBlockified(string parentDisplay) {
        string html = TableGeometrySource("<div style='display:" + parentDisplay
            + "'><span style='display:contents'><span style='width:max-content'>Text</span></span></div>");
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, TableGeometryOptions());

        Assert.Single(rendered.Diagnostics, item => item.Code == HtmlRenderDiagnosticCodes.IntrinsicSizeUnsupported);
    }

    [Theory]
    [InlineData("flex", "before")]
    [InlineData("flex", "after")]
    [InlineData("grid", "before")]
    [InlineData("grid", "after")]
    public void HtmlTableGeometry_ContentsPseudoItemsRetainIntrinsicLossWhenBlockified(string parentDisplay, string pseudo) {
        foreach (bool nested in new[] { false, true }) {
            string html = TableGeometrySource("<style>#contents{display:contents}#contents::" + pseudo
                + "{content:'Generated';min-width:max-content}</style><div style='display:" + parentDisplay + "'>"
                + (nested ? "<span style='display:contents'>" : string.Empty) + "<span id='contents'></span>"
                + (nested ? "</span>" : string.Empty) + "</div>");
            HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, TableGeometryOptions());

            HtmlDiagnostic diagnostic = Assert.Single(rendered.Diagnostics, item => item.Code == HtmlRenderDiagnosticCodes.IntrinsicSizeUnsupported);
            Assert.Contains("min-width=max-content", diagnostic.Detail);
            Assert.Single(rendered.Pages[0].Visuals.OfType<HtmlRenderText>(), text => text.Text == "Generated");
            Assert.Throws<HtmlConversionException>(() => rendered.RequireNoLoss());
            var strict = TableGeometryOptions();
            strict.FidelityPolicy = HtmlRenderFidelityPolicy.RequireNoLoss;
            Assert.Throws<HtmlConversionException>(() => HtmlRenderTestDriver.Render(html, strict));
        }
    }

    [Theory]
    [InlineData("block", "contents")]
    [InlineData("flex", "block")]
    public void HtmlTableGeometry_InlinePseudoDimensionsDoNotBorrowOuterFlexApplicability(string parentDisplay, string originatingDisplay) {
        string html = TableGeometrySource("<style>#contents{display:" + originatingDisplay
            + "}#contents::before{content:'Generated';min-width:max-content}</style><div style='display:"
            + parentDisplay + "'><span id='contents'></span></div>");
        var options = TableGeometryOptions();
        options.FidelityPolicy = HtmlRenderFidelityPolicy.RequireNoLoss;
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, options);

        Assert.DoesNotContain(rendered.Diagnostics, item => item.Code == HtmlRenderDiagnosticCodes.IntrinsicSizeUnsupported);
        Assert.Single(rendered.Pages[0].Visuals.OfType<HtmlRenderText>(), text => text.Text == "Generated");
        rendered.RequireNoLoss();
    }

    [Theory]
    [InlineData("width")]
    [InlineData("min-width")]
    [InlineData("max-width")]
    [InlineData("height")]
    [InlineData("min-height")]
    [InlineData("max-height")]
    public void HtmlTableGeometry_StylesheetIntrinsicDimensionsRetainApplicableLoss(string property) {
        string html = TableGeometrySource("<style>#sized{" + property + ":max-content}</style><div id='sized'>Text</div>");
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, TableGeometryOptions());

        HtmlDiagnostic diagnostic = Assert.Single(rendered.Diagnostics, item => item.Code == HtmlRenderDiagnosticCodes.IntrinsicSizeUnsupported);
        Assert.Contains(property + "=max-content", diagnostic.Detail);
        Assert.Throws<HtmlConversionException>(() => rendered.RequireNoLoss());
    }

    [Theory]
    [InlineData("width:max-content;width:100px", false, 100D)]
    [InlineData("width:100px;width:max-content", true, 600D)]
    [InlineData("width:100px!important;width:max-content", false, 100D)]
    [InlineData("width:max-content!important;width:100px", true, 600D)]
    [InlineData("width:100px;width:bogus", false, 100D)]
    [InlineData("width:100px;width:-10px", false, 100D)]
    [InlineData("width:100px;width:fit-content(auto)", false, 100D)]
    [InlineData("width:100px;width:fit-content()", false, 100D)]
    [InlineData("width:100px;width:fit-content(120px)", true, 600D)]
    [InlineData("width:100px;width:fit-content(calc(50px + 10%))", true, 600D)]
    [InlineData("width:100px;width:var(--size);--size:max-content", true, 600D)]
    [InlineData("width:100px;width:initial", false, 600D)]
    public void HtmlTableGeometry_StylesheetDimensionsPreserveCascadeAndRejectInvalidLaterValues(string declarations, bool loss, double width) {
        string html = TableGeometrySource("<style>#sized{" + declarations + ";background:white}</style><div id='sized'>Text</div>");
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, TableGeometryOptions());

        Assert.Equal(width, TableGeometryShape(rendered, "div#sized").Width, 3);
        Assert.Equal(loss, rendered.HasLoss);
        if (loss) {
            Assert.Single(rendered.Diagnostics, item => item.Code == HtmlRenderDiagnosticCodes.IntrinsicSizeUnsupported);
            Assert.Throws<HtmlConversionException>(() => rendered.RequireNoLoss());
        } else {
            Assert.DoesNotContain(rendered.Diagnostics, item => item.Code == HtmlRenderDiagnosticCodes.IntrinsicSizeUnsupported);
            rendered.RequireNoLoss();
        }
    }

    [Theory]
    [InlineData("<svg style='display:inline;width:max-content' width='10' height='10'><rect width='10' height='10'/></svg>")]
    [InlineData("<input style='display:inline;width:max-content' value='Text'>")]
    public void HtmlTableGeometry_InlineSvgAndControlBoxDimensionsRetainLoss(string body) {
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(TableGeometrySource(body), TableGeometryOptions());

        Assert.Single(rendered.Diagnostics, item => item.Code == HtmlRenderDiagnosticCodes.IntrinsicSizeUnsupported);
    }

    [Theory]
    [InlineData("")]
    [InlineData("min-")]
    [InlineData("max-")]
    public void HtmlTableGeometry_InapplicableTableInternalAxesDoNotRejectNoLoss(string prefix) {
        string width = prefix + "width:max-content";
        string height = prefix + "height:max-content";
        string html = TableGeometrySource("<table style='width:200px'><colgroup style='" + height + "'><col style='" + height
            + "'></colgroup><thead style='" + width + "'><tr style='" + width + "'><th>Header</th></tr></thead>"
            + "<tbody style='" + width + "'><tr style='" + width + "'><td>Body</td></tr></tbody>"
            + "<tfoot style='" + width + "'><tr style='" + width + "'><td>Footer</td></tr></tfoot></table>");
        var options = TableGeometryOptions();
        options.FidelityPolicy = HtmlRenderFidelityPolicy.RequireNoLoss;
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, options);

        Assert.DoesNotContain(rendered.Diagnostics, item => item.Code == HtmlRenderDiagnosticCodes.IntrinsicSizeUnsupported);
        rendered.RequireNoLoss();
    }

    [Theory]
    [InlineData("content-box", 230D, 85D)]
    [InlineData("border-box", 200D, 100D)]
    [InlineData("", 200D, 100D)]
    public void HtmlTableGeometry_AuthoredWidthAndAutoMarginsRespectBoxSizing(string sizing, double width, double x) {
        string html = TableGeometrySource("<div style='width:400px'><table id='sized' style='width:200px;padding:10px;border:5px solid black;"
            + "border-spacing:0;margin:0 auto;background:white;" + (sizing.Length == 0 ? string.Empty : "box-sizing:" + sizing + ";")
            + "'><tr><td style='padding:0'>Sized</td></tr></table></div>");
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, TableGeometryOptions());
        HtmlRenderShape table = TableGeometryShape(rendered, "table#sized");

        Assert.Equal(width, table.Width, 3);
        Assert.Equal(x, table.X, 3);
        Assert.False(rendered.HasLoss);
    }
}
