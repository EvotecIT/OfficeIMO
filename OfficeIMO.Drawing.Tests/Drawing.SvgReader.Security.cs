using System.Text;
using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Tests;

public class DrawingSvgReaderSecurityTests {
    [Fact]
    public void SvgSafetyPredicateCountsInheritedPatternPaintPerRenderedElement() {
        const string onePatternUse = "<svg xmlns='http://www.w3.org/2000/svg' width='16' height='8' style='fill:none;stroke:url(#p)'>"
            + "<defs><pattern id='p'><rect width='1' height='1'/><rect x='2' width='1' height='1'/></pattern></defs>"
            + "<g><rect width='4' height='4'/></g></svg>";
        const string twoPatternUses = "<svg xmlns='http://www.w3.org/2000/svg' width='16' height='8' style='fill:none;stroke:url(#p)'>"
            + "<defs><pattern id='p'><rect width='1' height='1'/><rect x='2' width='1' height='1'/></pattern></defs>"
            + "<g><rect width='4' height='4'/><rect x='5' width='4' height='4'/></g></svg>";
        var options = new OfficeSvgDrawingReaderOptions { MaximumElements = 7 };

        Assert.True(OfficeSvgDrawingReader.IsWithinSafetyLimits(Encoding.UTF8.GetBytes(onePatternUse), options));
        Assert.False(OfficeSvgDrawingReader.IsWithinSafetyLimits(Encoding.UTF8.GetBytes(twoPatternUses), options));
    }

    [Fact]
    public void SvgSafetyPredicateCountsInheritedPatternDefinitions() {
        const string onePatternUse = "<svg xmlns='http://www.w3.org/2000/svg' width='16' height='8'>"
            + "<defs><pattern id='base'><rect width='1' height='1'/><rect x='2' width='1' height='1'/></pattern>"
            + "<pattern id='derived' href='#base'/></defs><rect width='4' height='4' fill='url(#derived)'/></svg>";
        const string twoPatternUses = "<svg xmlns='http://www.w3.org/2000/svg' width='16' height='8'>"
            + "<defs><pattern id='base'><rect width='1' height='1'/><rect x='2' width='1' height='1'/></pattern>"
            + "<pattern id='derived' href='#base'/></defs><rect width='4' height='4' fill='url(#derived)'/>"
            + "<rect x='5' width='4' height='4' fill='url(#derived)'/></svg>";
        var options = new OfficeSvgDrawingReaderOptions { MaximumElements = 7 };

        Assert.True(OfficeSvgDrawingReader.IsWithinSafetyLimits(Encoding.UTF8.GetBytes(onePatternUse), options));
        Assert.False(OfficeSvgDrawingReader.IsWithinSafetyLimits(Encoding.UTF8.GetBytes(twoPatternUses), options));
    }

    [Fact]
    public void SvgSafetyPredicateDoesNotExpandGradientPaintServersAsPatterns() {
        const string svg = "<svg xmlns='http://www.w3.org/2000/svg' width='16' height='8'>"
            + "<defs><linearGradient id='g'><stop offset='0' stop-color='red'/><stop offset='1' stop-color='blue'/></linearGradient></defs>"
            + "<rect width='4' height='4' fill='url(#g)'/><rect x='5' width='4' height='4' style='stroke:url(#g)'/></svg>";
        var options = new OfficeSvgDrawingReaderOptions { MaximumElements = 6 };

        Assert.True(OfficeSvgDrawingReader.IsWithinSafetyLimits(Encoding.UTF8.GetBytes(svg), options));
    }

    [Fact]
    public void SvgSafetyPredicateRejectsAnyStylesheetLocalReferenceBeforeRasterFallback() {
        const string svg = "<svg xmlns='http://www.w3.org/2000/svg' width='16' height='8'>"
            + "<style>.clipped { clip-path: url( '#c' ); }</style>"
            + "<defs><clipPath id='c'><rect width='1' height='1'/></clipPath></defs>"
            + "<rect class='clipped' width='4' height='4'/></svg>";

        Assert.False(OfficeSvgDrawingReader.IsWithinSafetyLimits(Encoding.UTF8.GetBytes(svg)));
    }

    [Fact]
    public void SvgSafetyPredicateRejectsEscapedStylesheetReferenceToRenderedDefinition() {
        const string svg = "<svg xmlns='http://www.w3.org/2000/svg' width='16' height='8'>"
            + "<style>.clipped { clip-path: url(\\#c); }</style>"
            + "<defs><clipPath id='c'><rect width='1' height='1'/></clipPath></defs>"
            + "<rect class='clipped' width='4' height='4'/></svg>";

        Assert.False(OfficeSvgDrawingReader.IsWithinSafetyLimits(Encoding.UTF8.GetBytes(svg)));
    }

    [Fact]
    public void SvgSafetyPredicateAllowsStylesheetExternalUrlsWithoutLocalExpansion() {
        const string svg = "<svg xmlns='http://www.w3.org/2000/svg' width='16' height='8'>"
            + "<style>.external { fill: url(https://example.test/pattern.svg); }</style>"
            + "<rect class='external' width='4' height='4'/></svg>";

        Assert.True(OfficeSvgDrawingReader.IsWithinSafetyLimits(Encoding.UTF8.GetBytes(svg)));
    }

    [Fact]
    public void SvgSafetyPredicateChargesInheritedMarkerMidPerVertex() {
        const string shortPath = "<svg xmlns='http://www.w3.org/2000/svg' width='16' height='8' marker-mid='url(#m)'>"
            + "<defs><marker id='m'><rect width='1' height='1'/><rect x='2' width='1' height='1'/></marker></defs>"
            + "<path d='M0 0 L1 1 L2 0'/></svg>";
        const string longPath = "<svg xmlns='http://www.w3.org/2000/svg' width='16' height='8' marker-mid='url(#m)'>"
            + "<defs><marker id='m'><rect width='1' height='1'/><rect x='2' width='1' height='1'/></marker></defs>"
            + "<path d='M0 0 L1 1 L2 0 L3 1'/></svg>";
        var options = new OfficeSvgDrawingReaderOptions { MaximumElements = 5 };

        Assert.True(OfficeSvgDrawingReader.IsWithinSafetyLimits(Encoding.UTF8.GetBytes(shortPath), options));
        Assert.False(OfficeSvgDrawingReader.IsWithinSafetyLimits(Encoding.UTF8.GetBytes(longPath), options));
    }

    [Fact]
    public void SvgSafetyPredicateChargesMarkerShorthandPerVertex() {
        const string svg = "<svg xmlns='http://www.w3.org/2000/svg' width='16' height='8'>"
            + "<defs><marker id='m'><rect width='1' height='1'/></marker></defs>"
            + "<polyline points='0,0 1,1 2,0 3,1' style='marker:url(#m)'/></svg>";
        var options = new OfficeSvgDrawingReaderOptions { MaximumElements = 5 };

        Assert.False(OfficeSvgDrawingReader.IsWithinSafetyLimits(Encoding.UTF8.GetBytes(svg), options));
    }

    [Fact]
    public void SvgSafetyPredicateChargesMarkerStartPerSubpath() {
        const string svg = "<svg xmlns='http://www.w3.org/2000/svg' width='16' height='8'>"
            + "<defs><marker id='m'><rect width='1' height='1'/><rect x='2' width='1' height='1'/></marker></defs>"
            + "<path d='M0 0 L1 1 M2 0 L3 1' marker-start='url(#m)'/></svg>";
        var options = new OfficeSvgDrawingReaderOptions { MaximumElements = 5 };

        Assert.False(OfficeSvgDrawingReader.IsWithinSafetyLimits(Encoding.UTF8.GetBytes(svg), options));
    }
}
