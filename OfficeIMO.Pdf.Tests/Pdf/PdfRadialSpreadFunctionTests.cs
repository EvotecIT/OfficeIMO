using System;
using System.Text;
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfRadialSpreadFunctionTests {
    [Fact]
    public void MaximumStopFieldKeepsPortableCalculatorProcedureDepth() {
        var stops = new OfficeGradientStop[1024];
        for (int i = 0; i < stops.Length; i++) stops[i] = new OfficeGradientStop(i / 1023D,
            OfficeColor.FromRgba((byte)(i % 256), (byte)((i * 7) % 256), (byte)((i * 17) % 256), (byte)(64 + i % 192)));
        string source = PdfRadialSpreadFunction.Build(1, 0, 0, 0, 0, 1, stops,
            OfficeGradientSpreadMode.Reflect, OfficeGradientColorInterpolation.Srgb, null, false, false, colorChannel: 0);
        int depth = 0, maximumDepth = 0;
        foreach (char token in source) {
            if (token == '{') maximumDepth = Math.Max(maximumDepth, ++depth);
            if (token == '}') depth--;
        }
        Assert.InRange(maximumDepth, 1, 10);
        Assert.True(PdfCalculatorProgram.TryParse(Encoding.ASCII.GetBytes(source), out var program));
        Assert.NotNull(program.Evaluate(new[] { .5D, .25D }, 1));
    }

    [Theory]
    [InlineData(0, false, 0)]
    [InlineData(.4, false, .25)]
    [InlineData(.4, true, .25)]
    [InlineData(1, false, 0)]
    [InlineData(1.5, false, 0)]
    [InlineData(1, true, 0)]
    [InlineData(1.5, true, 0)]
    public void CalculatorPreservesPeriodicCircleField(double focus, bool reverse, double smallRadius) {
        var stops = new[] { new OfficeGradientStop(0, new OfficeColor(255, 0, 0, 64)),
            new OfficeGradientStop(.3, new OfficeColor(0, 255, 0, 192)),
            new OfficeGradientStop(1, new OfficeColor(0, 0, 255, 128)) };
        foreach (var spread in new[] { OfficeGradientSpreadMode.Repeat, OfficeGradientSpreadMode.Reflect }) {
            var field = (reverse ? new OfficeRadialGradient(0, 0, 1, focus, 0, smallRadius, stops)
                : new OfficeRadialGradient(focus, 0, smallRadius, 0, 0, 1, stops)).WithSpreadMode(spread);
            foreach (bool alpha in new[] { false, true }) {
                string source = PdfRadialSpreadFunction.Build(field.StartX, field.StartY, field.StartRadius,
                    field.EndX, field.EndY, field.EndRadius, field.Stops, spread, field.ColorInterpolation,
                    null, false, alpha);
                Assert.True(PdfCalculatorProgram.TryParse(Encoding.ASCII.GetBytes(source), out var program));
                for (int y = -12; y <= 12; y++) for (int x = -12; x <= 12; x++) {
                    double px = x / 4D, py = y / 4D;
                    double ratio = field.SampleRatio(px, py);
                    var expected = double.IsNaN(ratio) ? OfficeColor.Transparent : ratio <= .3 ? OfficeGradientColors.Interpolate(stops[0].Color, stops[1].Color, ratio / .3, field.ColorInterpolation) : OfficeGradientColors.Interpolate(stops[1].Color, stops[2].Color, (ratio - .3) / .7, field.ColorInterpolation);
                    var actual = program.Evaluate(new[] { px, py }, alpha ? 1 : 3);
                    Assert.NotNull(actual);
                    double[] channels = alpha ? new[] { expected.A / 255D } : new[] { expected.R / 255D, expected.G / 255D, expected.B / 255D };
                    for (int c = 0; c < channels.Length; c++) Assert.InRange(Math.Abs(actual![c] - channels[c]), 0, 1D / 255D);
                }
            }
        }
    }
}
