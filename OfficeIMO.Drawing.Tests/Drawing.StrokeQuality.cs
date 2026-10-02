using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class DrawingStrokeQualityTests {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void SplittingAStrokePreservesItsPixelsAndOpacity(bool dashed) {
        var straight = new OfficeRasterImage(120, 90);
        var split = new OfficeRasterImage(120, 90);
        var color = OfficeColor.FromRgba(30, 90, 200, 128);
        var a = new[] { new OfficePoint(20, 30), new OfficePoint(100, 30) };
        var b = new[] { a[0], new OfficePoint(60, 30), a[1] };
        if (dashed) {
            new OfficeRasterCanvas(straight).DrawPatternedPolyline(a, color, 8, new[] { 15D, 7D });
            new OfficeRasterCanvas(split).DrawPatternedPolyline(b, color, 8, new[] { 15D, 7D });
        } else {
            new OfficeRasterCanvas(straight).DrawPolyline(a, color, 8);
            new OfficeRasterCanvas(split).DrawPolyline(b, color, 8);
        }
        Assert.Equal(straight.GetPixels(), split.GetPixels());
        Assert.InRange(MaxAlpha(split), 120, 128);
    }

    [Fact]
    public void IntersectingSubpathsAndClosedCornersArePaintedOnce() {
        var image = Svg("<path d='M20 30 H100 M60 10 V70 M20 50 H100 V70 H20 Z' fill='none' stroke='blue' stroke-width='8' stroke-opacity='.5'/>");
        Assert.InRange(image.GetPixel(60, 30).A, 127, 128);
        Assert.InRange(MaxAlpha(image), 127, 128);
        var rectangle = new OfficeRasterImage(120, 90);
        new OfficeRasterCanvas(rectangle).DrawRectangle(20, 20, 60, 40, OfficeColor.FromRgba(0, 0, 0, 128), 8);
        Assert.InRange(MaxAlpha(rectangle), 127, 128);
    }

    [Fact]
    public void FractionalStrokeWidthControlsAreaInsteadOfClampingToOnePixel() {
        int Coverage(double width) {
            var image = new OfficeRasterImage(100, 40);
            new OfficeRasterCanvas(image).DrawLine(10, 20, 90, 20, OfficeColor.Black, width);
            return Enumerable.Range(0, 40).Sum(y => image.GetPixel(50, y).A);
        }
        Assert.InRange(Coverage(.25), 62, 66);
        Assert.InRange(Coverage(1), 253, 257);
        Assert.InRange(Coverage(2), 508, 512);
    }

    [Fact]
    public void SvgCapsAndJoinsUseTheirRequestedGeometry() {
        OfficeRasterImage Cap(string cap) => Svg($"<path d='M20 30 H60 V70' fill='none' stroke='black' stroke-width='8' stroke-linecap='{cap}'/>");
        Assert.Equal(0, Cap("butt").GetPixel(17, 30).A);
        Assert.True(Cap("round").GetPixel(17, 30).A > 0);
        Assert.Equal(0, Cap("round").GetPixel(16, 26).A);
        Assert.Equal(255, Cap("square").GetPixel(16, 26).A);
        OfficeRasterImage Join(string join, double limit = 10) => Svg($"<path d='M20 70 L60 20 L100 70' fill='none' stroke='black' stroke-width='12' stroke-linejoin='{join}' stroke-miterlimit='{limit}'/>");
        Assert.True(Join("miter").GetPixel(60, 11).A > 0);
        Assert.Equal(0, Join("round").GetPixel(60, 11).A);
        Assert.Equal(0, Join("bevel").GetPixel(60, 14).A);
        Assert.True(Join("round").GetPixel(60, 14).A > 0);
        Assert.Equal(Join("bevel").GetPixels(), Join("miter", 1).GetPixels());
    }

    [Fact]
    public void DashedGradientStrokesRetainColorAndSinglePaintOpacity() {
        var image = Svg("<defs><linearGradient id='g'><stop stop-color='red' stop-opacity='.5'/><stop offset='1' stop-color='blue' stop-opacity='.5'/></linearGradient></defs><path d='M10 30 H60 H110' fill='none' stroke='url(#g)' stroke-width='8' stroke-dasharray='20 10' stroke-linecap='butt'/>");
        Assert.True(image.GetPixel(15, 30).R > image.GetPixel(15, 30).B);
        Assert.True(image.GetPixel(105, 30).B > image.GetPixel(105, 30).R);
        Assert.Equal(0, image.GetPixel(35, 30).A);
        Assert.InRange(MaxAlpha(image), 127, 128);
    }

    [Fact]
    public void ZeroGapsAreContinuousAndZeroDashesHaveRoundDots() {
        var solid = new OfficeRasterImage(80, 40);
        var dashed = new OfficeRasterImage(80, 40);
        new OfficeRasterCanvas(solid).DrawLine(10, 20, 70, 20, OfficeColor.Black, 2);
        new OfficeRasterCanvas(dashed).DrawDashedLine(10, 20, 70, 20, OfficeColor.Black, 2, 5, 0);
        Assert.Equal(solid.GetPixels(), dashed.GetPixels());
        var dots = Svg("<path d='M20 30 H100' stroke='black' stroke-width='4' stroke-dasharray='0 10' stroke-linecap='round'/>");
        Assert.True(dots.GetPixel(20, 30).A > 0);
        Assert.True(dots.GetPixel(30, 30).A > 0);
        Assert.Equal(0, dots.GetPixel(25, 30).A);
    }

    [Theory]
    [InlineData(.25)]
    [InlineData(.45)]
    [InlineData(.75)]
    public void ThinFillDetailsKeepTheirAreaAtDifferentSamplePhases(double phase) {
        var image = new OfficeRasterImage(100, 40);
        new OfficeRasterCanvas(image).FillPolygon(new[] { new OfficePoint(10, 20 + phase), new OfficePoint(90, 20 + phase), new OfficePoint(90, 20 + phase + .1), new OfficePoint(10, 20 + phase + .1) }, OfficeColor.Black);
        Assert.InRange(image.GetPixel(50, 20).A, 25, 26);
    }

    [Theory]
    [InlineData(1D)]
    [InlineData(4D)]
    [InlineData(16D)]
    public void QuadraticFlatteningTracksOutputResolution(double scale) {
        var start = new OfficePoint(10, 280);
        var control = new OfficePoint(150, -100);
        var end = new OfficePoint(290, 280);
        var commands = new[] { OfficePathCommand.MoveTo(start.X, start.Y), OfficePathCommand.QuadraticBezierTo(control.X, control.Y, end.X, end.Y) };
        var points = OfficePathFlattener.Flatten(commands, 0, 0, scale)[0].Points;
        int segments = points.Count - 1;
        for (int i = 0; i < segments; i++) {
            double t = (i + .5) / segments, inverse = 1 - t;
            double x = scale * (inverse * inverse * start.X + 2 * inverse * t * control.X + t * t * end.X);
            double y = scale * (inverse * inverse * start.Y + 2 * inverse * t * control.Y + t * t * end.Y);
            double chordX = (points[i].X + points[i + 1].X) / 2, chordY = (points[i].Y + points[i + 1].Y) / 2;
            Assert.InRange(Math.Sqrt((x - chordX) * (x - chordX) + (y - chordY) * (y - chordY)), 0, .05);
        }
    }

    private static int MaxAlpha(OfficeRasterImage image) => Enumerable.Range(0, image.Width).SelectMany(x => Enumerable.Range(0, image.Height).Select(y => (int)image.GetPixel(x, y).A)).Max();

    private static OfficeRasterImage Svg(string body) {
        Assert.True(OfficeSvgDrawingReader.TryRead(Encoding.UTF8.GetBytes($"<svg xmlns='http://www.w3.org/2000/svg' width='120' height='90'>{body}</svg>"), out var drawing, out int unsupported));
        Assert.Equal(0, unsupported);
        return OfficeDrawingRasterRenderer.Render(drawing!, new OfficeDrawingRasterRenderOptions { Background = OfficeColor.Transparent });
    }
}
