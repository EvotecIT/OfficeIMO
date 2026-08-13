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
    public void SvgSafetyPredicateRejectsStylesheetPatternPaintBeforeRasterFallback() {
        const string svg = "<svg xmlns='http://www.w3.org/2000/svg' width='16' height='8'>"
            + "<style>.patterned { fill: url(#p); }</style>"
            + "<defs><pattern id='p'><rect width='1' height='1'/></pattern></defs>"
            + "<rect class='patterned' width='4' height='4'/></svg>";

        Assert.False(OfficeSvgDrawingReader.IsWithinSafetyLimits(Encoding.UTF8.GetBytes(svg)));
    }
}
