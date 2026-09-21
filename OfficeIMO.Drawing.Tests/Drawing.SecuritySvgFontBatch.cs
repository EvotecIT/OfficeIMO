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
    public void NestedFullSizeViewportsUseAnAggregateIntermediateBudget() {
        string svg = "<svg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 4096 4096'>" +
            "<svg><svg><svg><rect width='1' height='1'/></svg></svg></svg></svg>";

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
    public void OversizedSvgTextPathIsOmittedBeforeGlyphSplitting() {
        string svg = "<svg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 100 20'><defs>" +
            "<path id='p' d='M0 10 H100'/></defs><text><textPath href='#p'>" +
            new string('a', 5000) + "</textPath></text></svg>";

        Assert.True(OfficeSvgDrawingReader.TryRead(Encoding.UTF8.GetBytes(svg), out _, out int unsupported));
        Assert.True(unsupported > 0);
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
