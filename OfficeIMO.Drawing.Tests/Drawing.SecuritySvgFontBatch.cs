using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using OfficeIMO.Drawing;
using OfficeIMO.TestAssets;
using Xunit;

namespace OfficeIMO.Tests;

public partial class DrawingTests {
    [Fact]
    public void RasterEffectsShareActualScaledIntermediatePixelBudget() {
        var inner = new OfficeDrawing(1, 1).AddText("A", 0, 0, 1, 1, new OfficeFontInfo("Arial", 1));
        var drawing = new OfficeDrawing(1, 1)
            .AddEffectDrawing(inner, OfficeTransform.Identity)
            .AddEffectDrawing(inner, OfficeTransform.Identity);
        var budget = new OfficeRasterTransformedTextBudget { IntermediatePixels = 64_000_000L - 4L };

        Assert.Throws<OfficeImageExportLimitException>(() => OfficeDrawingRasterRenderer.Render(drawing,
            new OfficeDrawingRasterRenderOptions { Scale = 2, TransformedTextBudget = budget }));
    }

    [Fact]
    public void RejectedPatternDoesNotConsumeTheRetainedSurfaceBudget() {
        byte[] png = OfficePngWriter.Encode(new OfficeRasterImage(1, 1, OfficeColor.Red));
        string svg = "<svg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 4096 4096'><defs>" +
            "<pattern id='p' patternUnits='userSpaceOnUse' width='1' height='1'>" +
            "<rect width='1' height='1'/></pattern></defs>" +
            "<rect width='1' height='1' fill='url(#p)'/>" +
            "<image href='data:image/png;base64," + Convert.ToBase64String(png) + "' width='1' height='1'/></svg>";

        Assert.True(OfficeSvgDrawingReader.TryRead(Encoding.UTF8.GetBytes(svg), out OfficeDrawing? drawing, out int unsupported));
        Assert.True(unsupported > 0);
        Assert.Single(EnumerateDrawingImages(drawing!));
    }

    [Fact]
    public void EmptyMarkersDoNotConsumeFullCanvasSurfaceBudget() {
        byte[] png = OfficePngWriter.Encode(new OfficeRasterImage(1, 1, OfficeColor.Red));
        string svg = "<svg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 4000 4000'><defs>" +
            "<marker id='empty' markerWidth='1' markerHeight='1'/></defs>" +
            string.Concat(Enumerable.Repeat(
                "<path d='M0 0 L1 1' marker-end='url(#empty)'/>", 4)) +
            "<image href='data:image/png;base64," + Convert.ToBase64String(png) + "' width='1' height='1'/></svg>";

        Assert.True(OfficeSvgDrawingReader.TryRead(Encoding.UTF8.GetBytes(svg), out OfficeDrawing? drawing, out _));
        Assert.Single(EnumerateDrawingImages(drawing!));
    }

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
    public void FaxFillScanRejectsMissingEndOfLineWithinBoundedWork() {
        byte[] encoded = new byte[8 * 1024 * 1024];

        Assert.Throws<InvalidDataException>(() => OfficeFaxDecoder.Decode(
            encoded, columns: 1, rows: 1, k: 0, endOfLine: true, byteAligned: false,
            blackIsOne: true, endOfBlock: false, maximumBytes: 1, CancellationToken.None));
    }

    [Fact]
    public void FaxFillScanAcceptsLongValidGroupThreeFill() {
        string bits = new string('0', 4096) + "000000000001" + "000111";
        bits = bits.PadRight((bits.Length + 7) / 8 * 8, '0');
        byte[] encoded = Enumerable.Range(0, bits.Length / 8)
            .Select(index => Convert.ToByte(bits.Substring(index * 8, 8), 2)).ToArray();

        byte[] decoded = OfficeFaxDecoder.Decode(encoded, columns: 1, rows: 1, k: 0,
            endOfLine: true, byteAligned: false, blackIsOne: true, endOfBlock: false,
            maximumBytes: 1, CancellationToken.None);

        Assert.Equal(new byte[] { 0 }, decoded);
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
    public void ForeignObjectCallbackNestedEffectsShareTheRasterBudget() {
        string svg = "<svg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 10 10'>" +
            "<foreignObject width='10' height='10'><div xmlns='http://www.w3.org/1999/xhtml'>Text</div></foreignObject></svg>";
        var options = new OfficeSvgDrawingReaderOptions {
            ForeignObjectRenderer = context => {
                var content = new OfficeDrawing(context.Width, context.Height);
                var large = new OfficeDrawing(4000, 4000);
                for (int i = 0; i < 5; i++) content.AddEffectDrawing(large, OfficeTransform.Identity);
                return content;
            }
        };

        Assert.True(OfficeSvgDrawingReader.TryRead(Encoding.UTF8.GetBytes(svg), options, out OfficeDrawing? drawing, out int unsupported));
        Assert.True(unsupported > 0);
        Assert.Empty(drawing!.Elements);
    }

    [Theory]
    [InlineData("effect")]
    [InlineData("tile")]
    [InlineData("image")]
    public void CachedForeignObjectNestedSurfacesAreChargedForEachPlacement(string kind) {
        string svg = "<svg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 10 10'><defs>" +
            "<foreignObject id='f' width='1' height='1'><div xmlns='http://www.w3.org/1999/xhtml'>Text</div></foreignObject>" +
            "</defs>" + string.Concat(Enumerable.Repeat("<use href='#f'/>", 70)) + "</svg>";
        byte[]? png = kind == "image" ? OfficePngWriter.Encode(new OfficeRasterImage(1000, 1000, OfficeColor.Red)) : null;
        int calls = 0;
        var options = new OfficeSvgDrawingReaderOptions {
            ForeignObjectRenderer = context => {
                calls++;
                var content = new OfficeDrawing(context.Width, context.Height);
                if (kind == "effect") content.AddEffectDrawing(new OfficeDrawing(1000, 1000), OfficeTransform.Identity);
                else if (kind == "tile") content.AddTilingPattern(new OfficeDrawing(1000, 1000),
                    new OfficeImagePlacement(0, 0, 1, 1), 1000, 1000);
                else content.AddImage(png!, "image/png", new OfficeImageProjection(new OfficeImagePlacement(0, 0, 1, 1)));
                return content;
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
    public void CommentOnlyForeignObjectsDoNotExhaustRendererCalls() {
        string empty = string.Concat(Enumerable.Repeat("<foreignObject width='1' height='1'><!-- comment --></foreignObject>", 129));
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
    public void DirectRenderedClipsDoNotConsumeEffectSurfaces() {
        string svg = "<svg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 4096 4096'>" +
            "<defs><clipPath id='c'><rect width='5' height='5'/></clipPath></defs>" +
            string.Concat(Enumerable.Repeat("<rect width='5' height='5' clip-path='url(#c)'/>", 4)) + "</svg>";

        Assert.True(OfficeSvgDrawingReader.TryRead(Encoding.UTF8.GetBytes(svg), out OfficeDrawing? drawing, out int unsupported));
        Assert.Equal(0, unsupported);
        Assert.Equal(4, drawing!.Elements.OfType<OfficeDrawingGroup>().Count());
    }

    [Fact]
    public void PlainLinksDoNotConsumeEffectSurfaces() {
        string svg = "<svg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 4096 4096'>" +
            string.Concat(Enumerable.Repeat("<a href='https://example.test/'><rect width='5' height='5'/></a>", 4)) + "</svg>";

        Assert.True(OfficeSvgDrawingReader.TryRead(Encoding.UTF8.GetBytes(svg), out OfficeDrawing? drawing, out int unsupported));
        Assert.Equal(0, unsupported);
        Assert.Equal(4, drawing!.Elements.OfType<OfficeDrawingLink>().Count());
    }

    [Fact]
    public void PlainOutlinedTextDoesNotConsumeEffectSurfaces() {
        string svg = "<svg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 4096 4096'>" +
            string.Concat(Enumerable.Repeat("<text x='1' y='15' font-family='Painted' font-size='12' fill='none' stroke='black'>A</text>", 4)) + "</svg>";
        var options = new OfficeSvgDrawingReaderOptions();
        options.Fonts.Add("Painted", ManagedTextShapingTestAssets.CreateFont('A', 'B'));

        Assert.True(OfficeSvgDrawingReader.TryRead(Encoding.UTF8.GetBytes(svg), options, out OfficeDrawing? drawing, out int unsupported));
        Assert.Equal(0, unsupported);
        Assert.Equal(4, drawing!.Elements.OfType<OfficeDrawingGroup>().Count(group => group.ActualText == "A"));
    }

    [Fact]
    public void NestedFullSizeViewportsUseAnAggregateIntermediateBudget() {
        string svg = "<svg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 4096 4096'>" +
            "<svg><svg><svg><rect width='1' height='1'/></svg></svg></svg></svg>";

        Assert.True(OfficeSvgDrawingReader.TryRead(Encoding.UTF8.GetBytes(svg), out _, out int unsupported));
        Assert.True(unsupported > 0);
    }

    [Fact]
    public void NestedViewportsRetainBothScenesWithinTheIntermediateBudget() {
        string nested = "<svg width='4000' height='4000' viewBox='0 0 4000 4000'><rect width='1' height='1'/></svg>";
        string svg = "<svg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 4000 4000'>" + nested + nested + "</svg>";

        Assert.True(OfficeSvgDrawingReader.TryRead(Encoding.UTF8.GetBytes(svg), out OfficeDrawing? drawing, out int unsupported));
        Assert.Equal(0, unsupported);
        Assert.Equal(2, drawing!.Elements.Count);
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
    public void RepeatedSymbolsRetainBothScenesWithinTheIntermediateBudget() {
        string svg = "<svg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 4000 4000'><defs>" +
            "<symbol id='s' viewBox='0 0 4000 4000'><rect width='1' height='1'/></symbol></defs>" +
            "<use href='#s' width='4000' height='4000'/><use href='#s' width='4000' height='4000'/></svg>";

        Assert.True(OfficeSvgDrawingReader.TryRead(Encoding.UTF8.GetBytes(svg), out OfficeDrawing? drawing, out int unsupported));
        Assert.Equal(0, unsupported);
        Assert.Equal(2, drawing!.Elements.Count);
    }

    [Fact]
    public void OffsetFiltersChargeOnlyTheAdditionalIntermediateSurface() {
        string svg = "<svg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 1024 1024'><defs>" +
            "<filter id='offset'><feOffset dx='1' dy='1'/></filter></defs>" +
            string.Concat(Enumerable.Repeat("<rect width='1' height='1' filter='url(#offset)'/>", 24)) + "</svg>";

        Assert.True(OfficeSvgDrawingReader.TryRead(Encoding.UTF8.GetBytes(svg), out OfficeDrawing? drawing, out int unsupported));
        Assert.Equal(0, unsupported);
        Assert.Equal(24, drawing!.Elements.Count);
    }

    [Theory]
    [InlineData("feGaussianBlur stdDeviation='1'", 950)]
    [InlineData("feDropShadow dx='1' dy='1' stdDeviation='1'", 920)]
    public void BlurredFiltersChargeOnlyTheirAdditionalIntermediateSurfaces(string primitive, int size) {
        string svg = "<svg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 " + size + " " + size + "'><defs>" +
            "<filter id='effect'><" + primitive + "/></filter></defs>" +
            string.Concat(Enumerable.Repeat("<rect width='1' height='1' filter='url(#effect)'/>", 6)) + "</svg>";

        Assert.True(OfficeSvgDrawingReader.TryRead(Encoding.UTF8.GetBytes(svg), out OfficeDrawing? drawing, out int unsupported));
        Assert.Equal(0, unsupported);
        Assert.Equal(6, drawing!.Elements.Count);
    }

    [Theory]
    [InlineData("fill")]
    [InlineData("stroke")]
    public void PatternLayersChargeOnlyTheAdditionalIntermediateSurfaces(string paint) {
        string svg = "<svg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 1000 1000'><defs>" +
            "<pattern id='p' patternUnits='userSpaceOnUse' width='1000' height='1000'>" +
            "<rect width='1' height='1'/></pattern></defs>" +
            string.Concat(Enumerable.Repeat(
                "<rect width='1' height='1' " + paint + "='url(#p)' " +
                (paint == "stroke" ? "fill='none' stroke-width='1'" : string.Empty) + "/>", 16)) + "</svg>";

        Assert.True(OfficeSvgDrawingReader.TryRead(Encoding.UTF8.GetBytes(svg), out OfficeDrawing? drawing, out int unsupported));
        Assert.Equal(0, unsupported);
        Assert.Equal(16, drawing!.Elements.Count);
    }

    [Theory]
    [InlineData("fill")]
    [InlineData("stroke")]
    public void PatternTileAndCanvasSurfacesShareOneBudget(string paint) {
        string svg = "<svg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 1000 1000'><defs>" +
            "<pattern id='p' patternUnits='userSpaceOnUse' width='4000' height='4000'>" +
            "<rect width='1' height='1'/></pattern></defs>" +
            string.Concat(Enumerable.Repeat("<rect width='1' height='1' " + paint + "='url(#p)' " +
                (paint == "stroke" ? "fill='none' stroke-width='1'" : string.Empty) + "/>", 4)) + "</svg>";

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

    [Theory]
    [InlineData(1.0, 63)]
    [InlineData(0.5, 31)]
    public void EmbeddedImagePlacementsChargeDecodedAndOpacitySurfaces(double opacity, int expectedRetained) {
        byte[] png = OfficePngWriter.Encode(new OfficeRasterImage(1000, 1000, OfficeColor.Red));
        string image = "data:image/png;base64," + Convert.ToBase64String(png);
        string placement = "<image href='" + image + "' width='1' height='1' opacity='" +
            opacity.ToString(System.Globalization.CultureInfo.InvariantCulture) + "'/>";
        string svg = "<svg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 10 10'>" +
            string.Concat(Enumerable.Repeat(placement, 70)) + "</svg>";

        Assert.True(OfficeSvgDrawingReader.TryRead(Encoding.UTF8.GetBytes(svg), out OfficeDrawing? drawing, out int unsupported));
        Assert.Equal(70 - expectedRetained, unsupported);
        Assert.Equal(expectedRetained, drawing!.Elements.Count);
    }

    [Fact]
    public void MaskedGroupsReserveAllRasterSurfacesBeforeRetainingTheNextGroup() {
        string group = "<g mask='url(#m)'><rect width='1' height='1'/></g>";
        string svg = "<svg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 3000 3000'><defs>" +
            "<mask id='m' maskUnits='userSpaceOnUse' x='0' y='0' width='3000' height='3000'>" +
            "<rect width='3000' height='3000' fill='white'/></mask></defs>" + group + group + "</svg>";

        Assert.True(OfficeSvgDrawingReader.TryRead(Encoding.UTF8.GetBytes(svg), out OfficeDrawing? drawing, out int unsupported));
        Assert.True(unsupported > 0);
        Assert.Single(drawing!.Elements);
    }

    [Fact]
    public void RootMaskBeyondTheRasterBudgetCannotExposeItsHiddenChildren() {
        string svg = "<svg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 4096 4096' mask='url(#m)'><defs>" +
            "<mask id='m' maskUnits='userSpaceOnUse' x='0' y='0' width='4096' height='4096'>" +
            "<rect width='4096' height='4096' fill='black'/></mask></defs>" +
            "<rect width='4096' height='4096' fill='red'/></svg>";

        Assert.False(OfficeSvgDrawingReader.TryRead(Encoding.UTF8.GetBytes(svg), out OfficeDrawing? drawing, out _));
        Assert.Null(drawing);
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
    public void MarkerScenesAndTheirPlacementLayersShareTheRasterBudget() {
        string svg = "<svg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 100 100'><defs>" +
            "<marker id='m' markerWidth='1' markerHeight='1' viewBox='0 0 4000 4000'>" +
            "<rect width='1' height='1'/></marker></defs>" +
            string.Concat(Enumerable.Repeat("<line x1='0' y1='0' x2='1' y2='1' marker-end='url(#m)'/>", 4)) +
            "</svg>";

        Assert.True(OfficeSvgDrawingReader.TryRead(Encoding.UTF8.GetBytes(svg), out _, out int unsupported));
        Assert.True(unsupported > 0);
    }

    [Fact]
    public void DiscardedMarkerLayerReleasesItsSceneReservations() {
        string svg = "<svg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 100 100'><defs>" +
            "<marker id='m' markerWidth='1' markerHeight='1' viewBox='0 0 4000 4000'>" +
            "<rect width='1' height='1'/></marker></defs>" +
            "<polyline points='0,0 1,1 2,2 3,3 4,4 5,5' marker-mid='url(#m)'/>" +
            "<g style='mix-blend-mode:multiply'><rect width='1' height='1'/></g></svg>";

        Assert.True(OfficeSvgDrawingReader.TryRead(Encoding.UTF8.GetBytes(svg), out OfficeDrawing? drawing, out int unsupported));
        Assert.True(unsupported > 0);
        Assert.Contains(drawing!.Elements, element => element is OfficeDrawingEffectGroup);
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
            "<text x='1' y='12' font-family='Scoped' fill='none' stroke='black'>A" + new string(' ', 4096) +
            "<tspan>B</tspan></text></svg>";
        var options = new OfficeSvgDrawingReaderOptions();
        options.Fonts.Add("Scoped", ManagedTextShapingTestAssets.CreateFontWithDistinctGlyphs('A', 'B'));

        Assert.True(OfficeSvgDrawingReader.TryRead(Encoding.UTF8.GetBytes(svg), options, out OfficeDrawing? drawing, out int unsupported));
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
    public void PatternedStrokeRetainsPositiveSubNanounitDashesWithinTheWorkBudget() {
        string svg = "<svg xmlns='http://www.w3.org/2000/svg' width='40' height='40' viewBox='0 0 0.000000004 0.000000004'><defs>" +
            "<pattern id='p' patternUnits='userSpaceOnUse' width='0.000000004' height='0.000000004'>" +
            "<rect width='0.000000004' height='0.000000004'/></pattern></defs>" +
            "<path d='M0 0.000000002 H0.000000004' fill='none' stroke='url(#p)' " +
            "stroke-width='0.0000000005' stroke-dasharray='0.0000000005 0.0000000005'/></svg>";

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
