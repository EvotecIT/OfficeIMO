using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Tests;

public partial class DrawingTests {
    [Fact]
    public void NestedEffectTextSharesTheRenderWideIntermediatePixelBudget() {
        var inner = new OfficeDrawing(20, 20).AddPositionedText("A", 0, 0, 10, 10,
            new OfficeImageFrameTransform(90, 5, 5), new OfficeFontInfo("Arial", 12), textAdvanceWidth: 8);
        var drawing = new OfficeDrawing(20, 20).AddEffectDrawing(inner, OfficeTransform.Identity);
        var budget = new OfficeRasterTransformedTextBudget { Pixels = 64_000_000L };

        Assert.Throws<OfficeImageExportLimitException>(() => OfficeDrawingRasterRenderer.Render(drawing,
            new OfficeDrawingRasterRenderOptions { TransformedTextBudget = budget }));
    }

    [Fact]
    public void FaxFillScanRejectsMissingEndOfLineWithoutScanningTheWholePayload() {
        byte[] encoded = new byte[8 * 1024 * 1024];

        Assert.Throws<InvalidDataException>(() => OfficeFaxDecoder.Decode(
            encoded, columns: 1, rows: 1, k: 0, endOfLine: true, byteAligned: false,
            blackIsOne: true, endOfBlock: false, maximumBytes: 1, CancellationToken.None));
    }

    [Fact]
    public void ReferencedSvgImagesShareOneValidatedEncodedPayload() {
        byte[] png = OfficePngWriter.Encode(new OfficeRasterImage(1, 1, OfficeColor.Red));
        string image = "data:image/png;base64," + Convert.ToBase64String(png);
        string svg = "<svg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 10 10'><defs>" +
            "<image id='i' href='" + image + "' width='1' height='1'/></defs>" +
            "<use href='#i' x='0'/><use href='#i' x='2'/><use href='#i' x='4'/><use href='#i' x='6'/></svg>";

        Assert.True(OfficeSvgDrawingReader.TryRead(Encoding.UTF8.GetBytes(svg), out OfficeDrawing? drawing, out int unsupported));
        Assert.Equal(0, unsupported);
        OfficeDrawingImage[] images = EnumerateDrawingImages(drawing!).ToArray();
        Assert.Equal(4, images.Length);
        Assert.All(images, item => Assert.Same(images[0].EncodedBytes, item.EncodedBytes));
    }

    [Fact]
    public void RepeatedForeignObjectReferencesInvokeTheRendererOnce() {
        string svg = "<svg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 100 10'><defs>" +
            "<foreignObject id='f' width='2' height='2'><div xmlns='http://www.w3.org/1999/xhtml'>Text</div></foreignObject>" +
            "</defs>" + string.Concat(Enumerable.Repeat("<use href='#f'/>", 30)) + "</svg>";
        int calls = 0;
        var options = new OfficeSvgDrawingReaderOptions {
            ForeignObjectRenderer = context => {
                calls++;
                return new OfficeDrawing(context.Width, context.Height);
            }
        };

        Assert.True(OfficeSvgDrawingReader.TryRead(Encoding.UTF8.GetBytes(svg), options, out _, out int unsupported));
        Assert.Equal(1, calls);
        Assert.Equal(0, unsupported);
    }

    [Fact]
    public void CachedForeignObjectPlacementsChargeTheirFullEffectSurface() {
        string svg = "<svg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 4096 4096'><defs>" +
            "<foreignObject id='f' width='1' height='1'><div xmlns='http://www.w3.org/1999/xhtml'>Text</div></foreignObject>" +
            "</defs>" + string.Concat(Enumerable.Repeat("<use href='#f'/>", 5)) + "</svg>";
        int calls = 0;
        var options = new OfficeSvgDrawingReaderOptions {
            ForeignObjectRenderer = context => {
                calls++;
                return new OfficeDrawing(context.Width, context.Height);
            }
        };

        Assert.True(OfficeSvgDrawingReader.TryRead(Encoding.UTF8.GetBytes(svg), options, out _, out int unsupported));
        Assert.Equal(1, calls);
        Assert.True(unsupported > 0);
    }

    [Fact]
    public void EmptyForeignObjectsDoNotExhaustRendererCalls() {
        string empty = string.Concat(Enumerable.Repeat("<foreignObject width='1' height='1'/>", 129));
        string svg = "<svg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 10 10'>" + empty +
            "<foreignObject width='1' height='1'><div xmlns='http://www.w3.org/1999/xhtml'>Text</div></foreignObject></svg>";
        int calls = 0;
        var options = new OfficeSvgDrawingReaderOptions {
            ForeignObjectRenderer = context => {
                calls++;
                return new OfficeDrawing(context.Width, context.Height);
            }
        };

        Assert.True(OfficeSvgDrawingReader.TryRead(Encoding.UTF8.GetBytes(svg), options, out _, out int unsupported));
        Assert.Equal(0, unsupported);
        Assert.Equal(1, calls);
    }

    [Fact]
    public void NestedFullSizeViewportsUseAnAggregateIntermediateBudget() {
        string svg = "<svg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 4096 4096'>" +
            "<svg><svg><svg><rect width='1' height='1'/></svg></svg></svg></svg>";

        Assert.True(OfficeSvgDrawingReader.TryRead(Encoding.UTF8.GetBytes(svg), out _, out int unsupported));
        Assert.True(unsupported > 0);
    }

    [Fact]
    public void TinyNestedViewportsChargeTheirLargeViewBoxScenes() {
        string nested = "<svg width='1' height='1' viewBox='0 0 4096 4096'><rect width='1' height='1'/></svg>";
        string svg = "<svg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 4096 4096'>" +
            string.Concat(Enumerable.Repeat(nested, 5)) + "</svg>";

        Assert.True(OfficeSvgDrawingReader.TryRead(Encoding.UTF8.GetBytes(svg), out _, out int unsupported));
        Assert.True(unsupported > 0);
    }

    [Fact]
    public void RepeatedSymbolsChargeTheirLargeViewBoxScenes() {
        string svg = "<svg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 4096 4096'><defs>" +
            "<symbol id='s' viewBox='0 0 4096 4096'><rect width='1' height='1'/></symbol></defs>" +
            string.Concat(Enumerable.Repeat("<use href='#s' width='1' height='1'/>", 5)) + "</svg>";

        Assert.True(OfficeSvgDrawingReader.TryRead(Encoding.UTF8.GetBytes(svg), out _, out int unsupported));
        Assert.True(unsupported > 0);
    }

    [Fact]
    public void RepeatedEmbeddedImagesChargeFullCanvasEffectSurfaces() {
        byte[] png = OfficePngWriter.Encode(new OfficeRasterImage(1, 1, OfficeColor.Red));
        string image = "data:image/png;base64," + Convert.ToBase64String(png);
        string svg = "<svg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 4096 4096'><defs>" +
            "<image id='i' href='" + image + "' width='1' height='1'/></defs>" +
            string.Concat(Enumerable.Repeat("<use href='#i'/>", 5)) + "</svg>";

        Assert.True(OfficeSvgDrawingReader.TryRead(Encoding.UTF8.GetBytes(svg), out _, out int unsupported));
        Assert.True(unsupported > 0);
    }

    [Fact]
    public void BlendedGroupsShareTheFullCanvasEffectSurfaceBudget() {
        string svg = "<svg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 4096 4096'>" +
            string.Concat(Enumerable.Repeat(
                "<g style='mix-blend-mode:multiply'><rect width='1' height='1'/></g>", 5)) + "</svg>";

        Assert.True(OfficeSvgDrawingReader.TryRead(Encoding.UTF8.GetBytes(svg), out _, out int unsupported));
        Assert.True(unsupported > 0);
    }

    [Fact]
    public void BlurSamplesChargeTheirFullCanvasIntermediateSurfaces() {
        string svg = "<svg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 4096 4096'><defs>" +
            "<filter id='blur'><feGaussianBlur stdDeviation='2'/></filter></defs>" +
            "<rect width='1' height='1' filter='url(#blur)'/></svg>";

        Assert.True(OfficeSvgDrawingReader.TryRead(Encoding.UTF8.GetBytes(svg), out _, out int unsupported));
        Assert.True(unsupported > 0);
    }

    [Fact]
    public void MarkerScenesChargeExpandedViewBoxPixelsAcrossPlacements() {
        string svg = "<svg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 100 100'><defs>" +
            "<marker id='m' markerWidth='1' markerHeight='1' viewBox='0 0 4000 4000'>" +
            "<rect width='1' height='1'/></marker></defs>" +
            "<polyline points='0,0 1,1 2,2 3,3 4,4 5,5 6,6 7,7 8,8 9,9' marker-mid='url(#m)'/>" +
            "</svg>";

        Assert.True(OfficeSvgDrawingReader.TryRead(Encoding.UTF8.GetBytes(svg), out _, out int unsupported));
        Assert.True(unsupported > 0);
    }

    [Fact]
    public void EmptyMarkersDoNotConsumeTheScenePixelBudget() {
        string svg = "<svg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 100 100'><defs>" +
            "<marker id='empty' markerWidth='1' markerHeight='1' viewBox='0 0 4000 4000'/>" +
            "<marker id='visible' markerWidth='1' markerHeight='1' viewBox='0 0 4000 4000'>" +
            "<rect width='1' height='1'/></marker></defs>" +
            "<polyline points='0,0 1,1 2,2 3,3 4,4 5,5 6,6 7,7 8,8' marker-mid='url(#empty)'/>" +
            "<polyline points='10,10 11,11 12,12' marker-mid='url(#visible)'/></svg>";

        Assert.True(OfficeSvgDrawingReader.TryRead(Encoding.UTF8.GetBytes(svg), out OfficeDrawing? drawing, out _));
        Assert.Contains(drawing!.Elements, element => element is OfficeDrawingEffectGroup);
    }

    [Fact]
    public void InvalidNestedViewportsDoNotConsumeThePixelBudget() {
        string invalid = "<svg width='4000' height='4000' viewBox='bad'><rect width='1' height='1'/></svg>";
        string svg = "<svg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 4000 4000'>" +
            string.Concat(Enumerable.Repeat(invalid, 4)) +
            "<svg width='4000' height='4000'><rect width='1' height='1'/></svg></svg>";

        Assert.True(OfficeSvgDrawingReader.TryRead(Encoding.UTF8.GetBytes(svg), out OfficeDrawing? drawing, out _));
        Assert.NotEmpty(drawing!.Elements);
    }

    [Fact]
    public void OversizedSvgTextPathIsOmittedBeforeGlyphSplitting() {
        string svg = "<svg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 100 20'><defs>" +
            "<path id='p' d='M0 10 H100'/></defs><text><textPath href='#p'>" +
            new string('a', 5000) + "</textPath></text></svg>";

        Assert.True(OfficeSvgDrawingReader.TryRead(Encoding.UTF8.GetBytes(svg), out _, out int unsupported));
        Assert.True(unsupported > 0);
    }

    [Fact]
    public void PaintedTextMeasuresTheNormalizedRunBeforeApplyingTheOutlineLimit() {
        string svg = "<svg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 100 20'>" +
            "<text x='1' y='12' fill='none' stroke='black'>A" + new string(' ', 4096) +
            "<tspan>B</tspan></text></svg>";

        Assert.True(OfficeSvgDrawingReader.TryRead(Encoding.UTF8.GetBytes(svg), out OfficeDrawing? drawing, out int unsupported));
        Assert.Equal(0, unsupported);
        Assert.NotEmpty(drawing!.Elements);
    }

    [Fact]
    public void SvgTextPathCountsUnicodeElementsRatherThanUtf16Units() {
        string svg = "<svg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 1000 20'><defs>" +
            "<path id='p' d='M0 10 H1000'/></defs><text font-size='8'><textPath href='#p'>" +
            string.Concat(Enumerable.Repeat("a\u0301", 2500)) + "</textPath></text></svg>";

        Assert.True(OfficeSvgDrawingReader.TryRead(Encoding.UTF8.GetBytes(svg), out OfficeDrawing? drawing, out int unsupported));
        Assert.Equal(0, unsupported);
        Assert.Contains(drawing!.Elements.OfType<OfficeDrawingText>(), text => text.Text.Contains("a\u0301", StringComparison.Ordinal));
    }

    [Fact]
    public void RepeatedTrefExpansionUsesOneDocumentWideTextBudget() {
        string svg = "<svg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 100 20'><defs>" +
            "<text id='t'>" + new string('a', 5000) + "</text></defs><text>" +
            string.Concat(Enumerable.Repeat("<tref href='#t'/>", 40)) + "</text></svg>";

        Assert.True(OfficeSvgDrawingReader.TryRead(Encoding.UTF8.GetBytes(svg), out _, out int unsupported));
        Assert.True(unsupported > 0);
    }

    [Fact]
    public void TinyPatternedDashesCannotKeepExpandingWithoutProgress() {
        string svg = "<svg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 100 20'><defs>" +
            "<pattern id='p' patternUnits='userSpaceOnUse' width='1' height='1'>" +
            "<rect width='1' height='1'/></pattern></defs>" +
            "<path d='M0 10 H100' fill='none' stroke='url(#p)' stroke-width='1' " +
            "stroke-dasharray='0.0000001 0.0000001'/></svg>";

        Assert.True(OfficeSvgDrawingReader.TryRead(Encoding.UTF8.GetBytes(svg), out _, out int unsupported));
        Assert.True(unsupported > 0);
    }

    [Fact]
    public void ShortPatternedStrokeAcceptsSmallUserUnitDashes() {
        string svg = "<svg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 0.1 0.1'><defs>" +
            "<pattern id='p' patternUnits='userSpaceOnUse' width='0.01' height='0.01'>" +
            "<rect width='0.01' height='0.01'/></pattern></defs>" +
            "<path d='M0 0.05 H0.05' fill='none' stroke='url(#p)' stroke-width='0.01' " +
            "stroke-dasharray='0.005 0.005'/></svg>";

        Assert.True(OfficeSvgDrawingReader.TryRead(Encoding.UTF8.GetBytes(svg), out OfficeDrawing? drawing, out int unsupported));
        Assert.Equal(0, unsupported);
        Assert.NotEmpty(drawing!.Elements);
    }

    [Fact]
    public void PatternedStrokeAcceptsZeroLengthDashArrayEntries() {
        string svg = "<svg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 100 20'><defs>" +
            "<pattern id='p' patternUnits='userSpaceOnUse' width='5' height='5'>" +
            "<rect width='5' height='5'/></pattern></defs>" +
            "<path d='M0 10 H100' fill='none' stroke='url(#p)' stroke-width='1' " +
            "stroke-dasharray='5 0'/></svg>";

        Assert.True(OfficeSvgDrawingReader.TryRead(Encoding.UTF8.GetBytes(svg), out OfficeDrawing? drawing, out int unsupported));
        Assert.Equal(0, unsupported);
        Assert.NotEmpty(drawing!.Elements);
    }

    private static IEnumerable<OfficeDrawingImage> EnumerateDrawingImages(OfficeDrawing drawing) {
        foreach (OfficeDrawingElement element in drawing.Elements) {
            if (element is OfficeDrawingImage drawingImage) yield return drawingImage;
            if (element is OfficeDrawingGroup group) {
                foreach (OfficeDrawingImage nestedImage in EnumerateDrawingImages(group.InnerDrawing)) yield return nestedImage;
            }
            if (element is OfficeDrawingEffectGroup effect) {
                foreach (OfficeDrawingImage nestedImage in EnumerateDrawingImages(effect.InnerDrawing)) yield return nestedImage;
            }
        }
    }
}
