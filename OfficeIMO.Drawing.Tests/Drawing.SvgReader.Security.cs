using System.Linq;
using System.Text;
using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Tests;

public class DrawingSvgReaderSecurityTests {
    [Theory]
    [InlineData("clip-path='none' x:clip-path='url(#c)'", false)]
    [InlineData("x:clip-path='url(#c)' clip-path='none'", true)]
    public void SvgSafetyPredicateUsesRasterizerPresentationAttributeIdentity(string attributes, bool expected) {
        string svg = "<svg xmlns='http://www.w3.org/2000/svg' xmlns:x='urn:test' width='16' height='8'>"
            + "<defs><clipPath id='c'><rect width='1' height='1'/><rect x='2' width='1' height='1'/></clipPath></defs>"
            + "<rect width='4' height='4' " + attributes + "/>"
            + "<rect x='5' width='4' height='4' " + attributes + "/></svg>";
        var options = new OfficeSvgDrawingReaderOptions { MaximumElements = 6 };

        Assert.Equal(expected, OfficeSvgDrawingReader.IsWithinSafetyLimits(Encoding.UTF8.GetBytes(svg), options));
    }

    [Fact]
    public void SvgSafetyPredicateCountsDefinitionPaintInheritedFromDomAncestors() {
        const string oneUse = "<svg xmlns='http://www.w3.org/2000/svg' width='16' height='8' fill='url(#large)'>"
            + "<defs><pattern id='large' fill='none'><rect width='1' height='1'/><rect x='2' width='1' height='1'/></pattern>"
            + "<pattern id='small'><rect width='1' height='1'/></pattern></defs>"
            + "<rect width='4' height='4' fill='url(#small)'/></svg>";
        const string twoUses = "<svg xmlns='http://www.w3.org/2000/svg' width='16' height='8' fill='url(#large)'>"
            + "<defs><pattern id='large' fill='none'><rect width='1' height='1'/><rect x='2' width='1' height='1'/></pattern>"
            + "<pattern id='small'><rect width='1' height='1'/></pattern></defs>"
            + "<rect width='4' height='4' fill='url(#small)'/>"
            + "<rect x='5' width='4' height='4' fill='url(#small)'/></svg>";
        var options = new OfficeSvgDrawingReaderOptions { MaximumElements = 8 };

        Assert.True(OfficeSvgDrawingReader.IsWithinSafetyLimits(Encoding.UTF8.GetBytes(oneUse), options));
        Assert.False(OfficeSvgDrawingReader.IsWithinSafetyLimits(Encoding.UTF8.GetBytes(twoUses), options));
    }

    [Theory]
    [InlineData("width='100000' height='100000' style='width:16px;height:8px'")]
    [InlineData("WIDTH='100000' HEIGHT='100000' style='width:16px;height:8px'")]
    [InlineData("xmlns:x='urn:test' x:width='100000' x:height='100000' style='width:16px;height:8px'")]
    public void SvgSafetyPredicateUsesRasterizerViewportAttributeIdentity(string dimensions) {
        string svg = "<svg xmlns='http://www.w3.org/2000/svg' " + dimensions + ">"
            + "<rect width='4' height='4'/></svg>";

        Assert.False(OfficeSvgDrawingReader.IsWithinSafetyLimits(Encoding.UTF8.GetBytes(svg)));
    }

    [Fact]
    public void SvgSafetyPredicateIgnoresInlineStyleForRootRasterViewport() {
        const string svg = "<svg xmlns='http://www.w3.org/2000/svg' style='width:100000px;height:100000px'>"
            + "<rect width='4' height='4'/></svg>";

        Assert.True(OfficeSvgDrawingReader.IsWithinSafetyLimits(Encoding.UTF8.GetBytes(svg)));
    }

    [Fact]
    public void SvgSafetyPredicateRejectsOverDeepInheritedGradientChain() {
        Assert.True(OfficeSvgDrawingReader.IsWithinSafetyLimits(Encoding.UTF8.GetBytes(BuildGradientChain(15))));
        Assert.False(OfficeSvgDrawingReader.IsWithinSafetyLimits(Encoding.UTF8.GetBytes(BuildGradientChain(16))));

        static string BuildGradientChain(int lastIndex) {
            var svg = new StringBuilder("<svg xmlns='http://www.w3.org/2000/svg' width='16' height='8'><defs>");
            svg.Append("<linearGradient id='g0'><stop stop-color='red'/></linearGradient>");
            for (int index = 1; index <= lastIndex; index++) {
                svg.Append("<linearGradient id='g").Append(index).Append("' href='#g").Append(index - 1).Append("'/>");
            }
            return svg.Append("</defs><rect width='4' height='4' fill='url(#g")
                .Append(lastIndex)
                .Append(")'/></svg>")
                .ToString();
        }
    }

    [Theory]
    [InlineData("<style>.painted { fill:url(#g16); }</style>", "class='painted'")]
    [InlineData("", "style='--paint:url(#g16);fill:var(--paint)'")]
    public void SvgSafetyPredicateRejectsIndirectOverDeepGradientReferences(string styleBlock, string consumerAttributes) {
        var svg = new StringBuilder("<svg xmlns='http://www.w3.org/2000/svg' width='16' height='8'>")
            .Append(styleBlock)
            .Append("<defs><linearGradient id='g0'><stop stop-color='red'/></linearGradient>");
        for (int index = 1; index <= 16; index++) {
            svg.Append("<linearGradient id='g").Append(index).Append("' href='#g").Append(index - 1).Append("'/>");
        }
        svg.Append("</defs><rect width='4' height='4' ").Append(consumerAttributes).Append("/></svg>");

        Assert.False(OfficeSvgDrawingReader.IsWithinSafetyLimits(Encoding.UTF8.GetBytes(svg.ToString())));
    }

    [Theory]
    [InlineData("x:style='--paint:url(#g16);fill:var(--paint)'")]
    [InlineData("style='fill:none' x:style='--paint:url(#g16);fill:var(--paint)'")]
    public void SvgSafetyPredicateUsesRasterizerInlineStyleIdentity(string consumerAttributes) {
        var svg = new StringBuilder("<svg xmlns='http://www.w3.org/2000/svg' xmlns:x='urn:test' width='16' height='8'>")
            .Append("<defs><linearGradient id='g0'><stop stop-color='red'/></linearGradient>");
        for (int index = 1; index <= 16; index++) {
            svg.Append("<linearGradient id='g").Append(index).Append("' href='#g").Append(index - 1).Append("'/>");
        }
        svg.Append("</defs><rect width='4' height='4' ").Append(consumerAttributes).Append("/></svg>");

        Assert.False(OfficeSvgDrawingReader.IsWithinSafetyLimits(Encoding.UTF8.GetBytes(svg.ToString())));
    }

    [Fact]
    public void SvgSafetyPredicateUsesLastRasterizerInlineStyleIdentity() {
        var svg = new StringBuilder("<svg xmlns='http://www.w3.org/2000/svg' xmlns:x='urn:test' width='16' height='8'>")
            .Append("<defs><linearGradient id='g0'><stop stop-color='red'/></linearGradient>");
        for (int index = 1; index <= 16; index++) {
            svg.Append("<linearGradient id='g").Append(index).Append("' href='#g").Append(index - 1).Append("'/>");
        }
        svg.Append("</defs><rect width='4' height='4' x:style='--paint:url(#g16);fill:var(--paint)' style='fill:none'/></svg>");

        Assert.True(OfficeSvgDrawingReader.IsWithinSafetyLimits(Encoding.UTF8.GetBytes(svg.ToString())));
    }

    [Fact]
    public void SvgSafetyPredicateCountsInheritedPatternPaintPerRenderedElement() {
        const string onePatternUse = "<svg xmlns='http://www.w3.org/2000/svg' width='16' height='8' style='fill:none;stroke:url(#p)'>"
            + "<defs stroke='none'><pattern id='p'><rect width='1' height='1'/><rect x='2' width='1' height='1'/></pattern></defs>"
            + "<g><rect width='4' height='4'/></g></svg>";
        const string twoPatternUses = "<svg xmlns='http://www.w3.org/2000/svg' width='16' height='8' style='fill:none;stroke:url(#p)'>"
            + "<defs stroke='none'><pattern id='p'><rect width='1' height='1'/><rect x='2' width='1' height='1'/></pattern></defs>"
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

    [Theory]
    [InlineData("clip-path:u/**/rl(#c)")]
    [InlineData("clip-path:url(/*comment*/#c)")]
    public void SvgSafetyPredicateRejectsCommentNormalizedStylesheetReferences(string declaration) {
        string svg = "<svg xmlns='http://www.w3.org/2000/svg' width='16' height='8'>"
            + "<style>.clipped { " + declaration + "; }</style>"
            + "<defs><clipPath id='c'><rect width='1' height='1'/></clipPath></defs>"
            + "<rect class='clipped' width='4' height='4'/></svg>";

        Assert.False(OfficeSvgDrawingReader.IsWithinSafetyLimits(Encoding.UTF8.GetBytes(svg)));
    }

    [Fact]
    public void SvgSafetyPredicateIgnoresLocalLookingReferencesInStylesheetCommentsAndStrings() {
        const string svg = "<svg xmlns='http://www.w3.org/2000/svg' width='16' height='8'>"
            + "<style>.external { fill:url(https://example.test/pattern.svg);"
            + "content:'url(#unused)';/* clip-path:url(#unused); */ }</style>"
            + "<defs><clipPath id='unused'><rect width='1' height='1'/></clipPath></defs>"
            + "<rect class='external' width='4' height='4'/></svg>";

        Assert.True(OfficeSvgDrawingReader.IsWithinSafetyLimits(Encoding.UTF8.GetBytes(svg)));
    }

    [Theory]
    [InlineData("u\\72l")]
    [InlineData("\\75rl")]
    [InlineData("u\\000072 l")]
    [InlineData("ur\\6c")]
    public void SvgSafetyPredicateRejectsEscapedStylesheetUrlFunction(string functionName) {
        string svg = "<svg xmlns='http://www.w3.org/2000/svg' width='16' height='8'>"
            + "<style>.clipped { clip-path: " + functionName + "(#c); }</style>"
            + "<defs><clipPath id='c'><rect width='1' height='1'/></clipPath></defs>"
            + "<rect class='clipped' width='4' height='4'/></svg>";

        Assert.False(OfficeSvgDrawingReader.IsWithinSafetyLimits(Encoding.UTF8.GetBytes(svg)));
    }

    [Fact]
    public void SvgSafetyPredicateAllowsStylesheetExternalUrlsWithoutLocalExpansion() {
        const string svg = "<svg xmlns='http://www.w3.org/2000/svg' width='16' height='8'>"
            + "<style>.external { fill: url(https://example.test/pattern.svg); }</style>"
            + "<defs><clipPath id='unused'><rect width='1' height='1'/></clipPath></defs>"
            + "<rect class='external' width='4' height='4'/></svg>";

        Assert.True(OfficeSvgDrawingReader.IsWithinSafetyLimits(Encoding.UTF8.GetBytes(svg)));
    }

    [Fact]
    public void SvgSafetyPredicateRejectsAmbiguousLocalUseReferences() {
        const string ambiguous = "<svg xmlns='http://www.w3.org/2000/svg' width='16' height='8'>"
            + "<defs><g id='duplicate'><rect width='1' height='1'/></g>"
            + "<g id='duplicate'><circle r='1'/></g></defs><use href='#duplicate'/></svg>";
        const string external = "<svg xmlns='http://www.w3.org/2000/svg' width='16' height='8'>"
            + "<defs><g id='local'><rect width='1' height='1'/></g></defs>"
            + "<use href='https://example.test/shapes.svg#external'/></svg>";

        Assert.False(OfficeSvgDrawingReader.IsWithinSafetyLimits(Encoding.UTF8.GetBytes(ambiguous)));
        Assert.True(OfficeSvgDrawingReader.IsWithinSafetyLimits(Encoding.UTF8.GetBytes(external)));
    }

    [Theory]
    [InlineData("xlink:href='https://example.test/shapes.svg#external' href='#local'")]
    [InlineData("href='#local' xlink:href='https://example.test/shapes.svg#external'")]
    public void SvgSafetyPredicateRejectsNamespaceCollidingUseReferences(string attributes) {
        string svg = "<svg xmlns='http://www.w3.org/2000/svg' xmlns:xlink='http://www.w3.org/1999/xlink' width='16' height='8'>"
            + "<defs><g id='local'><rect width='1' height='1'/></g></defs>"
            + "<use " + attributes + "/></svg>";

        Assert.False(OfficeSvgDrawingReader.IsWithinSafetyLimits(Encoding.UTF8.GetBytes(svg)));
    }

    [Fact]
    public void SvgSafetyPredicateRejectsAmbiguousInheritedPatternReferences() {
        const string ambiguous = "<svg xmlns='http://www.w3.org/2000/svg' width='16' height='8'>"
            + "<defs><pattern id='base'><rect width='1' height='1'/></pattern>"
            + "<pattern id='base'><circle r='1'/></pattern>"
            + "<pattern id='derived' href='#base'/></defs>"
            + "<rect width='4' height='4' fill='url(#derived)'/></svg>";

        Assert.False(OfficeSvgDrawingReader.IsWithinSafetyLimits(Encoding.UTF8.GetBytes(ambiguous)));
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

    [Fact]
    public void SvgSafetyPredicateChargesImportantInlineReferencesPerConsumer() {
        const string oneUse = "<svg xmlns='http://www.w3.org/2000/svg' width='16' height='8'>"
            + "<defs><clipPath id='c'><rect width='1' height='1'/><rect x='2' width='1' height='1'/></clipPath></defs>"
            + "<rect width='4' height='4' style='clip-path:url(#c) !important;clip-path:none'/></svg>";
        const string twoUses = "<svg xmlns='http://www.w3.org/2000/svg' width='16' height='8'>"
            + "<defs><clipPath id='c'><rect width='1' height='1'/><rect x='2' width='1' height='1'/></clipPath></defs>"
            + "<rect width='4' height='4' style='clip-path:url(#c) !important;clip-path:none'/>"
            + "<rect x='5' width='4' height='4' style='clip-path:url(#c) !important;clip-path:none'/></svg>";
        var options = new OfficeSvgDrawingReaderOptions { MaximumElements = 7 };

        Assert.True(OfficeSvgDrawingReader.IsWithinSafetyLimits(Encoding.UTF8.GetBytes(oneUse), options));
        Assert.False(OfficeSvgDrawingReader.IsWithinSafetyLimits(Encoding.UTF8.GetBytes(twoUses), options));
    }

    [Fact]
    public void SvgSafetyPredicateRejectsEscapedPriorityOnUrlBearingLocalReference() {
        const string svg = "<svg xmlns='http://www.w3.org/2000/svg' width='16' height='8'>"
            + "<defs><clipPath id='c'><rect width='1' height='1'/></clipPath></defs>"
            + "<rect width='4' height='4' style='clip-path:url(#c)!imp\\ortant;clip-path:none'/></svg>";

        Assert.False(OfficeSvgDrawingReader.IsWithinSafetyLimits(Encoding.UTF8.GetBytes(svg)));
    }

    [Fact]
    public void SvgSafetyPredicateRejectsEscapedInlineUrlFunction() {
        const string svg = "<svg xmlns='http://www.w3.org/2000/svg' width='16' height='8'>"
            + "<defs><clipPath id='c'><rect width='1' height='1'/></clipPath></defs>"
            + "<rect width='4' height='4' style='clip-path:u\\72l(#c)'/></svg>";

        Assert.False(OfficeSvgDrawingReader.IsWithinSafetyLimits(Encoding.UTF8.GetBytes(svg)));
    }

    [Fact]
    public void SvgSafetyPredicateChargesTextCharactersPerExpansion() {
        string text = new string('x', 201);
        string oneUse = "<svg xmlns='http://www.w3.org/2000/svg' width='16' height='8'><defs><text id='t'>"
            + text + "</text></defs><use href='#t'/></svg>";
        string twoUses = "<svg xmlns='http://www.w3.org/2000/svg' width='16' height='8'><defs><text id='t'>"
            + text + "</text></defs><use href='#t'/><use href='#t' x='1'/></svg>";
        var options = new OfficeSvgDrawingReaderOptions { MaximumElements = 5 };

        Assert.True(OfficeSvgDrawingReader.IsWithinSafetyLimits(Encoding.UTF8.GetBytes(oneUse), options));
        Assert.False(OfficeSvgDrawingReader.IsWithinSafetyLimits(Encoding.UTF8.GetBytes(twoUses), options));
    }

    [Fact]
    public void SvgSafetyPredicateChargesNestedTextCharactersPerExpansion() {
        string text = new string('x', 201);
        string oneUse = "<svg xmlns='http://www.w3.org/2000/svg' width='16' height='8'><defs><text id='t'><a>"
            + text + "</a></text></defs><use href='#t'/></svg>";
        string twoUses = "<svg xmlns='http://www.w3.org/2000/svg' width='16' height='8'><defs><text id='t'><a>"
            + text + "</a></text></defs><use href='#t'/><use href='#t' x='1'/></svg>";
        var options = new OfficeSvgDrawingReaderOptions { MaximumElements = 7 };

        Assert.True(OfficeSvgDrawingReader.IsWithinSafetyLimits(Encoding.UTF8.GetBytes(oneUse), options));
        Assert.False(OfficeSvgDrawingReader.IsWithinSafetyLimits(Encoding.UTF8.GetBytes(twoUses), options));
    }

    [Fact]
    public void SvgSafetyPredicateChargesTextPathGeometryPerConsumer() {
        string geometry = "M0 0" + string.Concat(Enumerable.Repeat(" L1 1", 10000));
        string oneUse = "<svg xmlns='http://www.w3.org/2000/svg' width='16' height='8'><defs><path id='p' d='"
            + geometry + "'/></defs><text><textPath href='#p'>a</textPath></text></svg>";
        string twoUses = "<svg xmlns='http://www.w3.org/2000/svg' width='16' height='8'><defs><path id='p' d='"
            + geometry + "'/></defs><text><textPath href='#p'>a</textPath>"
            + "<textPath href='#p'>b</textPath></text></svg>";

        Assert.True(OfficeSvgDrawingReader.IsWithinSafetyLimits(Encoding.UTF8.GetBytes(oneUse)));
        Assert.False(OfficeSvgDrawingReader.IsWithinSafetyLimits(Encoding.UTF8.GetBytes(twoUses)));
    }

    [Fact]
    public void SvgSafetyPredicateChargesAttributePayloadPerExpansion() {
        string href = "data:image/png;base64," + new string('A', 200);
        string oneUse = "<svg xmlns='http://www.w3.org/2000/svg' width='16' height='8'><defs><image id='i' href='"
            + href + "' width='1' height='1'/></defs><use href='#i'/></svg>";
        string twoUses = "<svg xmlns='http://www.w3.org/2000/svg' width='16' height='8'><defs><image id='i' href='"
            + href + "' width='1' height='1'/></defs><use href='#i'/><use href='#i' x='1'/></svg>";
        var options = new OfficeSvgDrawingReaderOptions { MaximumElements = 5 };

        Assert.True(OfficeSvgDrawingReader.IsWithinSafetyLimits(Encoding.UTF8.GetBytes(oneUse), options));
        Assert.False(OfficeSvgDrawingReader.IsWithinSafetyLimits(Encoding.UTF8.GetBytes(twoUses), options));
    }
}
