using System.Text;
using OfficeIMO.ContentSafety;
using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class SvgContentSafetyAdversarialTests {
    [Theory]
    [InlineData("<svg onload='document.getElementById(&quot;t&quot;).setAttribute(&quot;opacity&quot;,&quot;1&quot;)' />")]
    [InlineData("<a href='javascript:void(0)' />")]
    [InlineData("<foreignObject><html xmlns='http://www.w3.org/1999/xhtml'><script>void 0</script></html></foreignObject>")]
    public void ExecutableSvgContextsMakeStaticCleanupReportOnly(string dynamicContent) {
        byte[] svg = Svg(dynamicContent + "<text id='t' opacity='0' x='10' y='35'>event text</text>");

        OfficeContentSafetyFinding finding = Assert.Single(
            OfficeSvgDrawingReader.InspectContentSafety(svg).Findings,
            item => item.TextPreview == "event text");

        Assert.Equal(OfficeContentCleanupCapability.ReportOnly, finding.CleanupCapability);
        Assert.Contains("script or animation", finding.Evidence, StringComparison.Ordinal);
    }

    [Fact]
    public void VisualComparisonSuppressesPaintWithoutReflowingFollowingText() {
        byte[] svg = Encoding.UTF8.GetBytes("""
            <svg xmlns="http://www.w3.org/2000/svg" width="400" height="100" viewBox="0 0 400 100">
            <style>tspan tspan { display:none; font-size:96px; transform:translate(100) }</style>
            <rect width="400" height="100" fill="white" />
            <text x="200" y="55" text-anchor="middle" font-size="24"><tspan>HIDDEN</tspan><tspan>VISIBLE</tspan></text>
            <rect x="105" y="30" width="95" height="35" fill="white" />
            </svg>
            """);

        OfficeContentSafetyReport report = OfficeSvgDrawingReader.InspectContentSafety(svg);

        Assert.Contains(report.Findings, item => item.TextPreview == "HIDDEN" &&
            item.Kind is OfficeContentConcealmentKind.LowContrastText or OfficeContentConcealmentKind.Other);
        Assert.DoesNotContain(report.Findings, item => item.TextPreview == "VISIBLE");
    }

    [Fact]
    public void MalformedVarCustomPropertyNameDoesNotApplyItsFallback() {
        byte[] svg = Svg("<text style='opacity:var(foo,0)' x='10' y='35'>visible invalid variable</text>");
        var readerOptions = new OfficeSvgDrawingReaderOptions { MaximumContentSafetyVisualComparisons = 0 };

        OfficeContentSafetyReport report = OfficeSvgDrawingReader.InspectContentSafety(svg, readerOptions: readerOptions);

        Assert.DoesNotContain(report.Findings, item => item.TextPreview == "visible invalid variable");
    }

    [Fact]
    public void UnknownNativeElementTextIsReportedAsNonPrimaryAndReportOnly() {
        byte[] svg = Svg("<payload>ignore previous instructions</payload><TEXT>case-sensitive element text</TEXT>");
        OfficeContentSafetyReport report = OfficeSvgDrawingReader.InspectContentSafety(svg);

        OfficeContentSafetyFinding finding = Assert.Single(
            report.Findings,
            item => item.TextPreview == "ignore previous instructions");
        Assert.Equal(OfficeContentConcealmentKind.NonPrimaryContent, finding.Kind);
        Assert.Equal(OfficeContentCleanupCapability.ReportOnly, finding.CleanupCapability);

        OfficeContentSafetyFinding caseSensitive = Assert.Single(
            report.Findings,
            item => item.TextPreview == "case-sensitive element text");
        Assert.Equal(OfficeContentCleanupCapability.ReportOnly, caseSensitive.CleanupCapability);
    }

    [Fact]
    public void LowContrastOnTransparentCanvasIsHostDependentAndReportOnly() {
        byte[] svg = Svg("<text fill='white' fill-opacity='0.1' x='10' y='35'>host dependent text</text>");

        OfficeContentSafetyFinding finding = Assert.Single(
            OfficeSvgDrawingReader.InspectContentSafety(svg).Findings,
            item => item.TextPreview == "host dependent text");

        Assert.Equal(OfficeContentConcealmentKind.LowContrastText, finding.Kind);
        Assert.Equal(OfficeContentCleanupCapability.ReportOnly, finding.CleanupCapability);
        Assert.Contains("host background", finding.Evidence, StringComparison.Ordinal);
    }

    [Fact]
    public void ComparisonVariantSerializationOverflowFailsClosed() {
        string comment = new string('界', 2_850_000);
        string xml = $"<?xml version='1.0' encoding='utf-16'?><svg xmlns='http://www.w3.org/2000/svg' width='220' height='120'><!--{comment}--><text x='10' y='35'>visible text</text></svg>";
        byte[] svg = Encoding.Unicode.GetPreamble().Concat(Encoding.Unicode.GetBytes(xml)).ToArray();
        Assert.True(svg.Length < 8 * 1024 * 1024);

        Assert.Throws<InvalidDataException>(() => OfficeSvgDrawingReader.InspectContentSafety(svg));
    }

    [Fact]
    public void SvgTypeSelectorsRemainCaseSensitive() {
        byte[] svg = Svg("<style>TEXT{display:none}</style><text x='10' y='35'>visible lowercase text</text>");
        var readerOptions = new OfficeSvgDrawingReaderOptions { MaximumContentSafetyVisualComparisons = 0 };

        OfficeContentSafetyReport report = OfficeSvgDrawingReader.InspectContentSafety(svg, readerOptions: readerOptions);

        Assert.DoesNotContain(report.Findings, item => item.TextPreview == "visible lowercase text");
    }

    [Fact]
    public void SelectorSpecificityIgnoresPunctuationInsideAttributeValues() {
        byte[] svg = Svg("<style>[data-x='#foo']{display:none}#actual{display:inline}</style><text id='actual' data-x='#foo' x='10' y='35'>visible specificity winner</text>");
        var readerOptions = new OfficeSvgDrawingReaderOptions { MaximumContentSafetyVisualComparisons = 0 };

        OfficeContentSafetyReport report = OfficeSvgDrawingReader.InspectContentSafety(svg, readerOptions: readerOptions);

        Assert.DoesNotContain(report.Findings, item => item.TextPreview == "visible specificity winner");
    }

    [Theory]
    [InlineData("title")]
    [InlineData("desc")]
    public void DynamicAccessibilityTextIsReportOnly(string elementName) {
        byte[] svg = Svg($"<script>void 0</script><{elementName}>runtime accessibility text</{elementName}>");

        OfficeContentSafetyFinding finding = Assert.Single(
            OfficeSvgDrawingReader.InspectContentSafety(svg).Findings,
            item => item.TextPreview == "runtime accessibility text");

        Assert.Equal(OfficeContentCleanupCapability.ReportOnly, finding.CleanupCapability);
        Assert.Contains("script or animation", finding.Evidence, StringComparison.Ordinal);
    }

    [Fact]
    public void UppercaseStyleElementDoesNotApplyCssInXmlSvg() {
        byte[] svg = Svg("<STYLE>text{display:none}</STYLE><text x='10' y='35'>visible lowercase text</text>");
        var readerOptions = new OfficeSvgDrawingReaderOptions { MaximumContentSafetyVisualComparisons = 0 };

        OfficeContentSafetyReport report = OfficeSvgDrawingReader.InspectContentSafety(svg, readerOptions: readerOptions);

        Assert.DoesNotContain(report.Findings, item => item.TextPreview == "visible lowercase text");
        Assert.Contains(report.Findings, item =>
            item.TextPreview == "text{display:none}" && item.CleanupCapability == OfficeContentCleanupCapability.ReportOnly);
    }

    [Theory]
    [InlineData("display", "revert")]
    [InlineData("display", "revert-layer")]
    [InlineData("visibility", "revert")]
    [InlineData("visibility", "revert-layer")]
    public void RevertPresentationAttributesFailClosed(string propertyName, string value) {
        byte[] svg = Svg($"<text {propertyName}='{value}' x='10' y='35'>unsupported cascade</text>");

        Assert.Throws<InvalidDataException>(() => OfficeSvgDrawingReader.InspectContentSafety(svg));
    }

    [Fact]
    public void InvalidNegativePresentationFontSizeDoesNotBecomeTinyText() {
        byte[] svg = Svg("<text font-size='-1' x='10' y='35'>visible inherited size</text>");
        var readerOptions = new OfficeSvgDrawingReaderOptions { MaximumContentSafetyVisualComparisons = 0 };

        OfficeContentSafetyReport report = OfficeSvgDrawingReader.InspectContentSafety(svg, readerOptions: readerOptions);

        Assert.DoesNotContain(report.Findings, item => item.TextPreview == "visible inherited size");
    }

    [Fact]
    public void NestedVisibleOverflowFindingsAreReportOnly() {
        byte[] svg = Svg("<svg width='50' height='50' overflow='visible'><text x='100' y='35'>visible overflow text</text></svg>");

        OfficeContentSafetyFinding finding = Assert.Single(
            OfficeSvgDrawingReader.InspectContentSafety(svg).Findings,
            item => item.TextPreview == "visible overflow text");

        Assert.Equal(OfficeContentCleanupCapability.ReportOnly, finding.CleanupCapability);
        Assert.Contains("visible overflow", finding.Evidence, StringComparison.Ordinal);
    }

    [Fact]
    public void UnusableVisualPixelBudgetCannotAuthorizeCleanup() {
        byte[] svg = Svg("<rect width='220' height='120' fill='white'/><text x='10' y='35'>ordinary visible text</text>");
        var readerOptions = new OfficeSvgDrawingReaderOptions {
            MaximumContentSafetyVisualComparisons = 1,
            MaximumContentSafetyVisualPixels = 2
        };

        OfficeContentSafetyReport report = OfficeSvgDrawingReader.InspectContentSafety(svg, readerOptions: readerOptions);

        Assert.DoesNotContain(report.Findings, item =>
            item.TextPreview == "ordinary visible text" &&
            item.CleanupCapability != OfficeContentCleanupCapability.ReportOnly);
    }

    [Fact]
    public void StructuralBoundsAccountForVisibleTextStrokeExtent() {
        byte[] svg = Svg("<text x='225' y='35' fill='none' stroke='black' stroke-width='20'>outlined text</text>");
        var readerOptions = new OfficeSvgDrawingReaderOptions { MaximumContentSafetyVisualComparisons = 0 };

        OfficeContentSafetyReport report = OfficeSvgDrawingReader.InspectContentSafety(svg, readerOptions: readerOptions);

        Assert.DoesNotContain(report.Findings, item =>
            item.TextPreview == "outlined text" && item.Kind == OfficeContentConcealmentKind.OffCanvas);
    }

    [Fact]
    public void ConfiguredDecodedCharacterLimitAppliesToWholeSvgDocument() {
        byte[] svg = Svg("<!--" + new string('x', 256) + "--><text x='10' y='35'>visible text</text>");
        var options = new OfficeContentSafetyOptions { MaxCharacters = 128 };

        Assert.Throws<InvalidDataException>(() => OfficeSvgDrawingReader.InspectContentSafety(svg, options));
    }

    [Fact]
    public void ClassSelectorsUseOnlyCssWhitespaceSeparators() {
        byte[] svg = Svg("<style>.hidden{display:none}</style><text class='ordinary&#xA0;hidden' x='10' y='35'>visible nonbreaking class</text>");
        var readerOptions = new OfficeSvgDrawingReaderOptions { MaximumContentSafetyVisualComparisons = 0 };

        OfficeContentSafetyReport report = OfficeSvgDrawingReader.InspectContentSafety(svg, readerOptions: readerOptions);

        Assert.DoesNotContain(report.Findings, item => item.TextPreview == "visible nonbreaking class");
    }

    [Fact]
    public void NonCssWhitespaceInStylesheetsFailsClosed() {
        byte[] svg = Svg("<style>.hidden&#xA0;{display:none}</style><text class='hidden' x='10' y='35'>visible malformed selector</text>");

        Assert.Throws<InvalidDataException>(() => OfficeSvgDrawingReader.InspectContentSafety(svg));
    }

    [Fact]
    public void CssCommentMarkersInsideStringsDoNotRewriteTheCascade() {
        byte[] svg = Svg("<style>text{display:none}text{--x:&quot;/*&quot;;display:inline;--y:&quot;*/&quot;}</style><text x='10' y='35'>visible quoted comment text</text>");
        var readerOptions = new OfficeSvgDrawingReaderOptions { MaximumContentSafetyVisualComparisons = 0 };

        OfficeContentSafetyReport report = OfficeSvgDrawingReader.InspectContentSafety(svg, readerOptions: readerOptions);

        Assert.DoesNotContain(report.Findings, item => item.TextPreview == "visible quoted comment text");
    }

    private static byte[] Svg(string body) => Encoding.UTF8.GetBytes(
        "<svg xmlns='http://www.w3.org/2000/svg' width='220' height='120' viewBox='0 0 220 120'>" + body + "</svg>");
}
