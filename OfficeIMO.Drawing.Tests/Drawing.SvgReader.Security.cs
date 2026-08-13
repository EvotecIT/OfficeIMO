using System.Linq;
using System.Text;
using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Tests;

public class DrawingSvgReaderSecurityTests {
    [Fact]
    public void SvgSafetyPredicateTokenizesNestedCustomPropertyDeclarations() {
        const string svg = "<svg xmlns='http://www.w3.org/2000/svg' width='16' height='8'>"
            + "<defs><pattern id='p;x'><rect width='1' height='1'/></pattern></defs>"
            + "<rect width='4' height='4' style='--paint:url(#p;x);fill:var(--paint)'/></svg>";

        Assert.False(OfficeSvgDrawingReader.IsWithinSafetyLimits(Encoding.UTF8.GetBytes(svg)));
    }

    [Fact]
    public void SvgSafetyPredicateUsesRasterDefinitionIdSemantics() {
        const string svg = "<svg xmlns='http://www.w3.org/2000/svg' xmlns:x='urn:test' width='16' height='8'>"
            + "<style>.paint{clip-path:url(#clip)}</style>"
            + "<defs><clipPath id='unused' x:id='clip'><rect width='1' height='1'/></clipPath></defs>"
            + "<rect class='paint' width='4' height='4'/></svg>";

        Assert.False(OfficeSvgDrawingReader.IsWithinSafetyLimits(Encoding.UTF8.GetBytes(svg)));
    }

    [Fact]
    public void SvgSafetyPredicateBoundsEmbeddedRasterImagePixels() {
        byte[] png = OfficePngWriter.Encode(new OfficeRasterImage(1, 1, OfficeColor.White));
        string safeSvg = CreateEmbeddedImageSvg(png);
        WriteBigEndian(png, 16, 100_000);
        WriteBigEndian(png, 20, 100_000);
        string oversizedSvg = CreateEmbeddedImageSvg(png);

        Assert.True(OfficeSvgDrawingReader.IsWithinSafetyLimits(Encoding.UTF8.GetBytes(safeSvg)));
        Assert.False(OfficeSvgDrawingReader.IsWithinSafetyLimits(Encoding.UTF8.GetBytes(oversizedSvg)));
    }

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
    public void SvgSafetyPredicateChargesObjectBoundingBoxPatternChildrenPerConsumer() {
        static string CreateSvg(int patternChildren) {
            var svg = new StringBuilder("<svg xmlns='http://www.w3.org/2000/svg' width='16' height='16'>"
                + "<defs><pattern id='p' width='1' height='1' patternContentUnits='objectBoundingBox'>");
            for (int index = 0; index < patternChildren; index++) {
                svg.Append("<rect width='1' height='1'/>");
            }
            return svg.Append("</pattern></defs><rect width='16' height='16' fill='url(#p)'/></svg>").ToString();
        }

        Assert.True(OfficeSvgDrawingReader.IsWithinSafetyLimits(Encoding.UTF8.GetBytes(CreateSvg(254))));
        Assert.False(OfficeSvgDrawingReader.IsWithinSafetyLimits(Encoding.UTF8.GetBytes(CreateSvg(255))));
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
    public void SvgSafetyPredicateMatchesInlinePropertyNamesCaseInsensitively() {
        const string svg = "<svg xmlns='http://www.w3.org/2000/svg' width='16' height='8'>"
            + "<defs><clipPath id='c'><rect width='1' height='1'/><rect x='2' width='1' height='1'/></clipPath></defs>"
            + "<rect width='4' height='4' style='CLIP-PATH:url(#c)'/>"
            + "<rect x='5' width='4' height='4' style='CLIP-PATH:url(#c)'/></svg>";
        var options = new OfficeSvgDrawingReaderOptions { MaximumElements = 7 };

        Assert.False(OfficeSvgDrawingReader.IsWithinSafetyLimits(Encoding.UTF8.GetBytes(svg), options));
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
        byte[] png = OfficePngWriter.Encode(new OfficeRasterImage(1, 1, OfficeColor.White));
        string href = "data:image/png;base64," + Convert.ToBase64String(png);
        string oneUse = "<svg xmlns='http://www.w3.org/2000/svg' width='16' height='8'><defs><image id='i' href='"
            + href + "' width='1' height='1'/></defs><use href='#i'/></svg>";
        string twoUses = "<svg xmlns='http://www.w3.org/2000/svg' width='16' height='8'><defs><image id='i' href='"
            + href + "' width='1' height='1'/></defs><use href='#i'/><use href='#i' x='1'/></svg>";
        int imagePayloadUnits = (href.Length + 2) / 100;
        var options = new OfficeSvgDrawingReaderOptions { MaximumElements = 3 + imagePayloadUnits };

        Assert.True(OfficeSvgDrawingReader.IsWithinSafetyLimits(Encoding.UTF8.GetBytes(oneUse), options));
        Assert.False(OfficeSvgDrawingReader.IsWithinSafetyLimits(Encoding.UTF8.GetBytes(twoUses), options));
    }

    [Fact]
    public void SvgSafetyPredicateUsesRasterizerGeometryAttributeIdentity() {
        string commands = "M0 0" + string.Concat(Enumerable.Repeat(" L1 1", 20000));
        string svg = "<svg xmlns='http://www.w3.org/2000/svg' xmlns:x='urn:test' width='16' height='8'>"
            + "<path d='' x:d='" + commands + "'/></svg>";

        Assert.False(OfficeSvgDrawingReader.IsWithinSafetyLimits(Encoding.UTF8.GetBytes(svg)));
    }

    [Fact]
    public void SvgSafetyPredicateUsesRasterizerGeometryIdentityForMarkerPlacements() {
        const string svg = "<svg xmlns='http://www.w3.org/2000/svg' xmlns:x='urn:test' width='16' height='8'>"
            + "<defs><marker id='m'><rect width='1' height='1'/></marker></defs>"
            + "<polyline points='0,0' x:points='0,0 1,1 2,0 3,1' marker-mid='url(#m)'/></svg>";
        var options = new OfficeSvgDrawingReaderOptions { MaximumElements = 4 };

        Assert.False(OfficeSvgDrawingReader.IsWithinSafetyLimits(Encoding.UTF8.GetBytes(svg), options));
    }

    [Theory]
    [InlineData("matrix(1024 0 0 1 0 0)", true)]
    [InlineData("matrix(1025 0 0 1 0 0)", false)]
    [InlineData("matrix(1 0 0 1 1000000 0)", true)]
    [InlineData("matrix(1 0 0 1 1000001 0)", false)]
    public void SvgSafetyPredicateEnforcesTransformCeilings(string transform, bool expected) {
        string svg = "<svg xmlns='http://www.w3.org/2000/svg' width='16' height='8'>"
            + "<g transform='" + transform + "'><rect width='1' height='1'/></g></svg>";

        Assert.Equal(expected, OfficeSvgDrawingReader.IsWithinSafetyLimits(Encoding.UTF8.GetBytes(svg)));
    }

    [Fact]
    public void SvgSafetyPredicateEnforcesEffectiveReferencedTransformCeilings() {
        const string svg = "<svg xmlns='http://www.w3.org/2000/svg' width='16' height='8'>"
            + "<defs><g id='scaled' transform='scale(33)'><rect width='1' height='1'/></g></defs>"
            + "<use href='#scaled' transform='scale(33)'/></svg>";

        Assert.False(OfficeSvgDrawingReader.IsWithinSafetyLimits(Encoding.UTF8.GetBytes(svg)));
    }

    [Fact]
    public void SvgSafetyPredicateChecksRasterizerTransformIdentityAndOnlyRenderedDefinitions() {
        const string namespaceCollision = "<svg xmlns='http://www.w3.org/2000/svg' xmlns:x='urn:test' width='16' height='8'>"
            + "<g transform='scale(1)' x:transform='scale(1025)'><rect width='1' height='1'/></g></svg>";
        const string unusedDefinition = "<svg xmlns='http://www.w3.org/2000/svg' width='16' height='8'>"
            + "<defs><g id='unused' transform='scale(1025)'><rect width='1' height='1'/></g></defs>"
            + "<rect width='1' height='1'/></svg>";

        Assert.False(OfficeSvgDrawingReader.IsWithinSafetyLimits(Encoding.UTF8.GetBytes(namespaceCollision)));
        Assert.True(OfficeSvgDrawingReader.IsWithinSafetyLimits(Encoding.UTF8.GetBytes(unusedDefinition)));
    }

    [Theory]
    [InlineData("cl\\69p-path")]
    [InlineData("\\63 lip-path")]
    [InlineData("clip-\\70 ath")]
    [InlineData("CL\\49P-PATH")]
    public void SvgSafetyPredicateDecodesEscapedInlinePropertyNames(string propertyName) {
        string svg = "<svg xmlns='http://www.w3.org/2000/svg' width='16' height='8'>"
            + "<defs><clipPath id='c'><rect width='1' height='1'/><rect x='2' width='1' height='1'/></clipPath></defs>"
            + "<rect width='4' height='4' style='" + propertyName + ":url(#c)'/>"
            + "<rect x='5' width='4' height='4' style='" + propertyName + ":url(#c)'/></svg>";
        var options = new OfficeSvgDrawingReaderOptions { MaximumElements = 7 };

        Assert.False(OfficeSvgDrawingReader.IsWithinSafetyLimits(Encoding.UTF8.GetBytes(svg), options));
    }

    [Fact]
    public void SvgSafetyPredicateBoundsAggregateRasterOverdraw() {
        Assert.True(IsSafe(256));
        Assert.False(IsSafe(257));

        static bool IsSafe(int repaints) {
            var svg = new StringBuilder("<svg xmlns='http://www.w3.org/2000/svg' width='4096' height='4096'>");
            for (int index = 0; index < repaints; index++) {
                svg.Append("<rect width='4096' height='4096' fill='rgba(0,0,0,.5)'/>");
            }
            svg.Append("</svg>");
            return OfficeSvgDrawingReader.IsWithinSafetyLimits(Encoding.UTF8.GetBytes(svg.ToString()));
        }
    }

    [Theory]
    [InlineData("1", true)]
    [InlineData("1000000", false)]
    public void SvgSafetyPredicateBoundsRasterFilterWork(string standardDeviation, bool expected) {
        string svg = "<svg xmlns='http://www.w3.org/2000/svg' width='256' height='256'>"
            + "<defs><filter id='f'><feGaussianBlur stdDeviation='" + standardDeviation + "'/></filter></defs>"
            + "<rect width='256' height='256' filter='url(#f)'/></svg>";

        Assert.Equal(expected, OfficeSvgDrawingReader.IsWithinSafetyLimits(Encoding.UTF8.GetBytes(svg)));
    }

    [Theory]
    [InlineData("<feMorphology radius='1000000'/>")]
    [InlineData("<feConvolveMatrix order='1000000'/>")]
    [InlineData("<feTurbulence numOctaves='1000000'/>")]
    public void SvgSafetyPredicateBoundsSiblingFilterParameters(string primitive) {
        string svg = "<svg xmlns='http://www.w3.org/2000/svg' width='256' height='256'>"
            + "<defs><filter id='f'>" + primitive + "</filter></defs>"
            + "<rect width='256' height='256' filter='url(#f)'/></svg>";

        Assert.False(OfficeSvgDrawingReader.IsWithinSafetyLimits(Encoding.UTF8.GetBytes(svg)));
    }

    [Fact]
    public void SvgSafetyPredicateSharesFilterWorkAcrossConsumers() {
        var svg = new StringBuilder("<svg xmlns='http://www.w3.org/2000/svg' width='256' height='256'>")
            .Append("<defs><filter id='f'><feGaussianBlur stdDeviation='1'/></filter></defs>");
        for (int index = 0; index < 20; index++) {
            svg.Append("<rect width='256' height='256' filter='url(#f)'/>");
        }
        svg.Append("</svg>");

        Assert.False(OfficeSvgDrawingReader.IsWithinSafetyLimits(Encoding.UTF8.GetBytes(svg.ToString())));
    }

    [Theory]
    [InlineData("href='#expensive'")]
    [InlineData("x:href='#expensive'")]
    public void SvgSafetyPredicateChargesInheritedFilterWork(string hrefAttribute) {
        string svg = "<svg xmlns='http://www.w3.org/2000/svg' xmlns:x='urn:test' width='256' height='256'>"
            + "<defs><filter id='expensive'><feGaussianBlur stdDeviation='1000000'/></filter>"
            + "<filter id='f' " + hrefAttribute + "/></defs>"
            + "<rect width='256' height='256' filter='url(#f)'/></svg>";

        Assert.False(OfficeSvgDrawingReader.IsWithinSafetyLimits(Encoding.UTF8.GetBytes(svg)));
    }

    [Theory]
    [InlineData(16, true)]
    [InlineData(17, false)]
    public void SvgSafetyPredicateBoundsInheritedFilterDepth(int filterCount, bool expected) {
        var svg = new StringBuilder("<svg xmlns='http://www.w3.org/2000/svg' width='256' height='256'><defs>");
        for (int index = 0; index < filterCount; index++) {
            svg.Append("<filter id='f").Append(index).Append("'");
            if (index + 1 < filterCount) svg.Append(" href='#f").Append(index + 1).Append("'");
            svg.Append(index + 1 == filterCount ? "><feGaussianBlur stdDeviation='0'/></filter>" : "/>");
        }
        svg.Append("</defs><rect width='256' height='256' filter='url(#f0)'/></svg>");

        Assert.Equal(expected, OfficeSvgDrawingReader.IsWithinSafetyLimits(Encoding.UTF8.GetBytes(svg.ToString())));
    }

    [Fact]
    public void SvgSafetyPredicateRejectsInheritedFilterCycle() {
        string svg = "<svg xmlns='http://www.w3.org/2000/svg' width='256' height='256'>"
            + "<defs><filter id='f0' href='#f1'/><filter id='f1' href='#f0'/></defs>"
            + "<rect width='256' height='256' filter='url(#f0)'/></svg>";

        Assert.False(OfficeSvgDrawingReader.IsWithinSafetyLimits(Encoding.UTF8.GetBytes(svg)));
    }

    [Theory]
    [InlineData("clip-path", "clipPath")]
    [InlineData("mask", "mask")]
    [InlineData("filter", "filter")]
    public void SvgSafetyPredicateChargesRootEffectReferences(string propertyName, string definitionName) {
        var svg = new StringBuilder("<svg xmlns='http://www.w3.org/2000/svg' width='64' height='64' ")
            .Append(propertyName).Append("='url(#effect)'><defs><").Append(definitionName).Append(" id='effect'>");
        for (int index = 0; index < 257; index++) svg.Append("<rect width='64' height='64'/>");
        svg.Append("</").Append(definitionName).Append("></defs><rect width='64' height='64'/></svg>");

        Assert.False(OfficeSvgDrawingReader.IsWithinSafetyLimits(Encoding.UTF8.GetBytes(svg.ToString())));
    }

    [Fact]
    public void SvgSafetyPredicateChargesUsePlacement() {
        var svg = new StringBuilder("<svg xmlns='http://www.w3.org/2000/svg' width='64' height='64'>")
            .Append("<defs><rect id='tile' x='10000' width='64' height='64'/></defs>");
        for (int index = 0; index < 257; index++) svg.Append("<use href='#tile' x='-10000'/>");
        svg.Append("</svg>");

        Assert.False(OfficeSvgDrawingReader.IsWithinSafetyLimits(Encoding.UTF8.GetBytes(svg.ToString())));
    }

    [Fact]
    public void SvgSafetyPredicateChargesSymbolViewportScaling() {
        var svg = new StringBuilder("<svg xmlns='http://www.w3.org/2000/svg' width='64' height='64'>")
            .Append("<defs><symbol id='tile' viewBox='0 0 1 1'><rect width='1' height='1'/></symbol></defs>");
        for (int index = 0; index < 257; index++) svg.Append("<use href='#tile' width='64' height='64'/>");
        svg.Append("</svg>");

        Assert.False(OfficeSvgDrawingReader.IsWithinSafetyLimits(Encoding.UTF8.GetBytes(svg.ToString())));
    }

    [Fact]
    public void SvgSafetyPredicateChargesSliceViewportScaling() {
        var svg = new StringBuilder("<svg xmlns='http://www.w3.org/2000/svg' width='4096' height='4096' viewBox='0 0 4096 1' preserveAspectRatio='xMidYMid slice'>");
        for (int index = 0; index < 257; index++) svg.Append("<rect width='1' height='1'/>");
        svg.Append("</svg>");

        Assert.False(OfficeSvgDrawingReader.IsWithinSafetyLimits(Encoding.UTF8.GetBytes(svg.ToString())));
    }

    [Theory]
    [InlineData("x:viewBox='0 0 1 1'")]
    [InlineData("viewBox='0 0 4096 4096' x:viewBox='0 0 1 1'")]
    public void SvgSafetyPredicateUsesProjectedRootViewBox(string viewBoxAttributes) {
        var svg = new StringBuilder("<svg xmlns='http://www.w3.org/2000/svg' xmlns:x='urn:test' width='4096' height='4096' ")
            .Append(viewBoxAttributes).Append(">");
        for (int index = 0; index < 257; index++) svg.Append("<rect width='1' height='1'/>");
        svg.Append("</svg>");

        Assert.False(OfficeSvgDrawingReader.IsWithinSafetyLimits(Encoding.UTF8.GetBytes(svg.ToString())));
    }

    [Fact]
    public void SvgSafetyPredicateUsesProjectedRootPreserveAspectRatio() {
        var svg = new StringBuilder("<svg xmlns='http://www.w3.org/2000/svg' xmlns:x='urn:test' width='4096' height='4096' viewBox='0 0 4096 1' x:preserveAspectRatio='xMidYMid slice'>");
        for (int index = 0; index < 257; index++) svg.Append("<rect width='1' height='1'/>");
        svg.Append("</svg>");

        Assert.False(OfficeSvgDrawingReader.IsWithinSafetyLimits(Encoding.UTF8.GetBytes(svg.ToString())));
    }

    [Theory]
    [InlineData("stroke-width='20000'")]
    [InlineData("stroke-width='1' stroke-linejoin='miter' stroke-miterlimit='20000'")]
    public void SvgSafetyPredicateChargesInheritedStrokeExtents(string strokeAttributes) {
        Assert.True(IsSafe(257, strokeAttributes: null));
        Assert.False(IsSafe(257, strokeAttributes));

        static bool IsSafe(int strokes, string? strokeAttributes) {
            var svg = new StringBuilder("<svg xmlns='http://www.w3.org/2000/svg' width='4096' height='4096'><g");
            if (strokeAttributes != null) svg.Append(" stroke='black' ").Append(strokeAttributes);
            svg.Append(">");
            for (int index = 0; index < strokes; index++) svg.Append("<rect x='10000' width='1' height='1'/>");
            svg.Append("</g></svg>");
            return OfficeSvgDrawingReader.IsWithinSafetyLimits(Encoding.UTF8.GetBytes(svg.ToString()));
        }
    }

    [Fact]
    public void SvgSafetyPredicateChargesNestedSvgViewportTransforms() {
        Assert.True(IsSafe(1, "x='0' y='0' width='1024' height='1024' viewBox='100 0 1 1'", "x='100'"));
        Assert.False(IsSafe(4097, "x='0' y='0' width='1024' height='1024' viewBox='100 0 1 1'", "x='100'"));
        Assert.False(IsSafe(4097, "x='10000' y='0'", "x='-10000'"));

        static bool IsSafe(int viewports, string viewportAttributes, string rectangleAttributes) {
            var svg = new StringBuilder("<svg xmlns='http://www.w3.org/2000/svg' width='4096' height='4096'>");
            for (int index = 0; index < viewports; index++) {
                svg.Append("<svg ").Append(viewportAttributes).Append("><rect ")
                    .Append(rectangleAttributes).Append(" width='1' height='1'/></svg>");
            }
            svg.Append("</svg>");
            return OfficeSvgDrawingReader.IsWithinSafetyLimits(Encoding.UTF8.GetBytes(svg.ToString()));
        }
    }

    [Fact]
    public void SvgSafetyPredicateChargesMarkerApplicationsConservatively() {
        Assert.True(IsSafe(1));
        Assert.False(IsSafe(257));

        static bool IsSafe(int applications) {
            var svg = new StringBuilder("<svg xmlns='http://www.w3.org/2000/svg' width='4096' height='4096'>")
                .Append("<defs><marker id='m' markerWidth='4096' markerHeight='4096' viewBox='0 0 1 1'>")
                .Append("<rect width='1' height='1'/></marker></defs><polyline marker-mid='url(#m)' points='");
            for (int index = 0; index < applications + 2; index++) svg.Append(index).Append(",0 ");
            svg.Append("'/></svg>");
            return OfficeSvgDrawingReader.IsWithinSafetyLimits(Encoding.UTF8.GetBytes(svg.ToString()));
        }
    }

    [Fact]
    public void SvgSafetyPredicateChargesMarkerDescendantsConservatively() {
        var svg = new StringBuilder("<svg xmlns='http://www.w3.org/2000/svg' width='4096' height='4096'>")
            .Append("<defs><marker id='m' markerWidth='4096' markerHeight='4096' viewBox='10000 0 1 1' refX='10000'>");
        for (int index = 0; index < 257; index++) svg.Append("<rect x='10000' width='1' height='1'/>");
        svg.Append("</marker></defs><line marker-start='url(#m)'/></svg>");

        Assert.False(OfficeSvgDrawingReader.IsWithinSafetyLimits(Encoding.UTF8.GetBytes(svg.ToString())));
    }

    [Theory]
    [InlineData("stroke-width")]
    [InlineData("str\\6f ke-width")]
    public void SvgSafetyPredicateRejectsStylesheetGeometryDeclarations(string propertyName) {
        const string paintOnly = "<svg xmlns='http://www.w3.org/2000/svg' width='16' height='8'>"
            + "<style>.painted{fill:red;content:'stroke-width:4096'}</style><rect class='painted' width='1' height='1'/></svg>";
        string geometry = "<svg xmlns='http://www.w3.org/2000/svg' width='4096' height='4096'>"
            + "<style>line{stroke:black;" + propertyName + ":4096}</style><line x1='0' x2='4096'/></svg>";

        Assert.True(OfficeSvgDrawingReader.IsWithinSafetyLimits(Encoding.UTF8.GetBytes(paintOnly)));
        Assert.False(OfficeSvgDrawingReader.IsWithinSafetyLimits(Encoding.UTF8.GetBytes(geometry)));
    }

    [Theory]
    [InlineData("transform:scale(4096)")]
    [InlineData("tr\\61nsform:scale(4096)")]
    [InlineData("letter-spacing:4096")]
    [InlineData("word-spacing:4096")]
    [InlineData("baseline-shift:4096")]
    [InlineData("writing-mode:vertical-rl")]
    public void SvgSafetyPredicateRejectsStylesheetLayoutGeometry(string declaration) {
        string svg = "<svg xmlns='http://www.w3.org/2000/svg' width='4096' height='4096'>"
            + "<style>rect{" + declaration + "}</style><rect width='1' height='1'/></svg>";

        Assert.False(OfficeSvgDrawingReader.IsWithinSafetyLimits(Encoding.UTF8.GetBytes(svg)));
    }

    [Fact]
    public void SvgSafetyPredicateChargesStylesheetNonScalingStrokePerPaintedElement() {
        const string safe = "<svg xmlns='http://www.w3.org/2000/svg' width='16' height='16'>"
            + "<style>.line{vector-effect:non-scaling-stroke}</style>"
            + "<rect class='line' width='1' height='1' fill='none' stroke='black'/></svg>";
        var amplified = new StringBuilder("<svg xmlns='http://www.w3.org/2000/svg' width='16' height='16'>"
            + "<style>.line{vector-effect:non-scaling-stroke}</style>");
        for (int index = 0; index < 257; index++) {
            amplified.Append("<rect class='line' width='1' height='1' fill='none' stroke='black'/>");
        }
        amplified.Append("</svg>");

        Assert.True(OfficeSvgDrawingReader.IsWithinSafetyLimits(Encoding.UTF8.GetBytes(safe)));
        Assert.False(OfficeSvgDrawingReader.IsWithinSafetyLimits(Encoding.UTF8.GetBytes(amplified.ToString())));
    }

    [Theory]
    [InlineData("transform:scale(4096)")]
    [InlineData("tr\\61nsform:scale(4096)")]
    [InlineData("transform-origin:50% 50%")]
    [InlineData("transform-box:fill-box")]
    public void SvgSafetyPredicateRejectsUnmodeledInlineTransforms(string declaration) {
        string svg = "<svg xmlns='http://www.w3.org/2000/svg' width='4096' height='4096'>"
            + "<rect style='" + declaration + "' width='1' height='1'/></svg>";

        Assert.False(OfficeSvgDrawingReader.IsWithinSafetyLimits(Encoding.UTF8.GetBytes(svg)));
    }

    [Theory]
    [InlineData("stroke-dasharray='1e-300'")]
    [InlineData("style='stroke-dasharray:1e-300'")]
    public void SvgSafetyPredicateRejectsUnboundedDashWork(string dashAttribute) {
        string svg = "<svg xmlns='http://www.w3.org/2000/svg' width='16' height='8'>"
            + "<path stroke='black' " + dashAttribute + " d='M0 0L1 0'/></svg>";

        Assert.False(OfficeSvgDrawingReader.IsWithinSafetyLimits(Encoding.UTF8.GetBytes(svg)));
    }

    [Fact]
    public void SvgSafetyPredicateChargesOffCanvasTextIntermediates() {
        var svg = new StringBuilder("<svg xmlns='http://www.w3.org/2000/svg' width='4096' height='4096'><g font-size='4096'>");
        for (int index = 0; index < 600; index++) svg.Append("<text x='10000'>X</text>");
        svg.Append("</g></svg>");

        Assert.False(OfficeSvgDrawingReader.IsWithinSafetyLimits(Encoding.UTF8.GetBytes(svg.ToString())));
    }

    [Fact]
    public void SvgSafetyPredicateChargesInheritedAnchoredTextBounds() {
        Assert.True(IsSafe(1, anchored: false));
        Assert.True(IsSafe(1, anchored: true));
        Assert.False(IsSafe(257, anchored: true));

        static bool IsSafe(int runs, bool anchored) {
            var svg = new StringBuilder("<svg xmlns='http://www.w3.org/2000/svg' width='4096' height='4096'>")
                .Append("<g font-size='4096'");
            if (anchored) svg.Append(" text-anchor='end'");
            svg.Append(">");
            for (int index = 0; index < runs; index++) svg.Append("<text x='4096' y='4096'>X</text>");
            svg.Append("</g></svg>");
            return OfficeSvgDrawingReader.IsWithinSafetyLimits(Encoding.UTF8.GetBytes(svg.ToString()));
        }
    }

    [Fact]
    public void SvgSafetyPredicateRejectsAmbiguousTextPositioning() {
        var svg = new StringBuilder("<svg xmlns='http://www.w3.org/2000/svg' width='4096' height='4096'>");
        for (int index = 0; index < 257; index++) {
            svg.Append("<text x='10000' dx='-10000' font-size='4096'>X</text>");
        }
        svg.Append("</svg>");

        Assert.False(OfficeSvgDrawingReader.IsWithinSafetyLimits(Encoding.UTF8.GetBytes(svg.ToString())));
    }

    private static string CreateEmbeddedImageSvg(byte[] png) =>
        "<svg xmlns='http://www.w3.org/2000/svg' width='16' height='8'><image width='1' height='1' href='data:image/png;base64,"
        + Convert.ToBase64String(png)
        + "'/></svg>";

    private static void WriteBigEndian(byte[] bytes, int offset, int value) {
        bytes[offset] = (byte)(value >> 24);
        bytes[offset + 1] = (byte)(value >> 16);
        bytes[offset + 2] = (byte)(value >> 8);
        bytes[offset + 3] = (byte)value;
    }
}
