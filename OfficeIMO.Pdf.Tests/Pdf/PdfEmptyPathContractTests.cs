using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using System.Text;
using System.Text.RegularExpressions;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class PdfEmptyPathContractTests {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void DeclaredMoveOnlyPathsReachPublicPdfShapeSurfaces(bool close) {
        var commands = close ? new[] { OfficePathCommand.MoveTo(5, 5), OfficePathCommand.Close() }
            : new[] { OfficePathCommand.MoveTo(5, 5) };
        var shape = OfficeShape.Path(20, 20, commands);
        shape.FillColor = OfficeColor.Red; shape.StrokeColor = OfficeColor.Blue;
        shape.StrokeLineCap = OfficeStrokeLineCap.Round;
        var options = new PdfOptions { CompressContentStreams = false };
        byte[][] documents = {
            PdfDocument.Create(options).Shape(shape).ToBytes(),
            PdfDocument.Create(options).Drawing(new OfficeDrawing(20, 20).AddShape(shape, 0, 0)).ToBytes(),
            PdfDocument.Create(options).Compose(b => b.Page(p => p.Canvas(c => c.Shape(shape, 0, 0)))).ToBytes(),
            PdfDocument.Create(options).Compose(b => b.Page(p => { p.Header(h => h.Shape(shape)); p.Content(c => c.Spacer(10)); })).ToBytes(),
            PdfDocument.Create(options).Compose(b => b.Page(p => { p.Footer(h => h.Shape(shape)); p.Content(c => c.Spacer(10)); })).ToBytes()
        };
        Assert.All(documents, bytes => {
            Assert.Single(PdfReadDocument.Open(bytes).Pages);
            // PDF strokes a single-point closed subpath with a round cap as a dot.
            Assert.DoesNotMatch(@"\bm\s+h\b", Encoding.ASCII.GetString(bytes));
        });
    }

    [Theory]
    [InlineData(0)]
    [InlineData(1)]
    [InlineData(2)]
    [InlineData(3)]
    public void EmptyContoursDoNotCloseButActualZeroLengthLinesRetainCaps(int placement) {
        var shape = OfficeShape.Path(20, 20,
            OfficePathCommand.MoveTo(5, 5), OfficePathCommand.Close(),
            OfficePathCommand.MoveTo(10, 10), OfficePathCommand.LineTo(10, 10), OfficePathCommand.Close(),
            OfficePathCommand.MoveTo(15, 15), OfficePathCommand.Close());
        shape.StrokeColor = OfficeColor.Blue; shape.StrokeWidth = 4;
        shape.StrokeLineCap = OfficeStrokeLineCap.Round;
        var document = PdfDocument.Create(new PdfOptions { CompressContentStreams = false });
        if (placement == 0) document.Shape(shape);
        else if (placement == 1) document.Compose(b => b.Page(p => p.Canvas(c => c.Shape(shape, 0, 0))));
        else document.Compose(b => b.Page(p => {
            if (placement == 2) p.Header(h => h.Shape(shape));
            else p.Footer(h => h.Shape(shape));
            p.Content(c => c.Spacer(10));
        }));
        string pdf = Encoding.ASCII.GetString(document.ToBytes());
        Assert.DoesNotMatch(@"\bm\s+h\b", pdf);
        Assert.Matches(@"\bl\s+h\b", pdf);
        Assert.Contains("1 J", pdf);
    }

    [Fact]
    public void DeclaredMoveOnlyClipsReachPublicShapeAndImageSurfaces() {
        var clip = OfficeClipPath.Path(20, 20, new[] { OfficePathCommand.MoveTo(5, 5) });
        var shape = OfficeShape.Rectangle(20, 20); shape.FillColor = OfficeColor.Red; shape.ClipPath = clip;
        byte[] image = OfficePngWriter.Encode(new OfficeRasterImage(20, 20, OfficeColor.Red));
        byte[][] documents = { PdfDocument.Create().Shape(shape).ToBytes(), PdfDocument.Create().Image(image, 20, 20, clipPath: clip).ToBytes() };
        Assert.All(documents, bytes => Assert.Single(PdfReadDocument.Open(bytes).Pages));
    }

    [Fact]
    public void EmptyPathAcceptanceKeepsMissingCommandAndCoordinateGuards() {
        var missing = new OfficeShape { Kind = OfficeShapeKind.Path, Width = 20, Height = 20 };
        Assert.Throws<System.ArgumentException>(() => PdfDocument.Create().Shape(missing));
        var outside = OfficeShape.Path(20, 20, OfficePathCommand.MoveTo(21, 21));
        Assert.Throws<System.ArgumentOutOfRangeException>(() => PdfDocument.Create().Shape(outside));
        var clip = OfficeClipPath.Path(20, 20, new[] { OfficePathCommand.MoveTo(21, 21) });
        byte[] image = OfficePngWriter.Encode(new OfficeRasterImage(20, 20, OfficeColor.Red));
        Assert.Throws<System.ArgumentOutOfRangeException>(() => PdfDocument.Create().Image(image, 20, 20, clipPath: clip));
    }
}
