using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class DrawingJpegForwardTransformTests {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void QuantizedCoefficientsMatchIndependentTwoDimensionalDct(bool scalar) {
        var random = new Random(104729);
        var input = new int[64];
        var quantization = new int[64];
        var actual = new int[64];
        var workspace = new double[64];
        for (int block = 0; block < 24; block++) {
            for (int index = 0; index < 64; index++) {
                input[index] = random.Next(-128, 128);
                quantization[index] = block < 8 ? 1 : random.Next(1, 256);
            }
            if (scalar) OfficeJpegForwardTransform.QuantizeScalar(input, quantization, actual, workspace);
            else OfficeJpegForwardTransform.Quantize(input, quantization, actual, workspace);
            for (int v = 0; v < 8; v++) {
                for (int u = 0; u < 8; u++) {
                    double sum = 0;
                    for (int y = 0; y < 8; y++) {
                        for (int x = 0; x < 8; x++) {
                            sum += input[y * 8 + x] *
                                Math.Cos((2 * x + 1) * u * Math.PI / 16) *
                                Math.Cos((2 * y + 1) * v * Math.PI / 16);
                        }
                    }
                    double coefficient = sum * (u == 0 ? 1 / Math.Sqrt(2) : 1) *
                        (v == 0 ? 1 / Math.Sqrt(2) : 1) / 4;
                    int expected = (int)Math.Round(coefficient / quantization[v * 8 + u]);
                    // Algebraically equivalent summation can cross a half-integer
                    // rounding boundary by one quantized unit.
                    Assert.InRange(Math.Abs(actual[v * 8 + u] - expected), 0, 1);
                    if (actual[v * 8 + u] != expected) {
                        double distanceToHalf = Math.Abs(Math.Abs(coefficient / quantization[v * 8 + u] - expected) - 0.5D);
                        Assert.InRange(distanceToHalf, 0D, 0.0001D);
                    }
                }
            }
        }
    }

    [Theory]
    [InlineData(-128)]
    [InlineData(0)]
    [InlineData(127)]
    public void ConstantBlockHasOnlyTheExpectedDcCoefficient(int level) {
        var input = Enumerable.Repeat(level, 64).ToArray();
        var quantization = Enumerable.Repeat(1, 64).ToArray();
        var result = new int[64];
        OfficeJpegForwardTransform.Quantize(input, quantization, result, new double[64]);
        Assert.Equal(level * 8, result[0]);
        Assert.All(result.Skip(1), coefficient => Assert.Equal(0, coefficient));
    }
}
