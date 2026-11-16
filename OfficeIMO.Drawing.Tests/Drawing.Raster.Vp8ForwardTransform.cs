using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class DrawingVp8ForwardTransformTests {
    [Fact]
    public void ForwardTransformIsInvertedByTheRfc6386InverseTransform() {
        var residual = new int[16];
        var rows = new double[16];
        var coefficients = new double[16];
        var quantized = new int[16];
        uint random = 0x6C8E9CF5U;
        int worst = 0;
        for (int block = 0; block < 20000; block++) {
            for (int index = 0; index < 16; index++) {
                random ^= random << 13; random ^= random >> 17; random ^= random << 5;
                residual[index] = (int)(random % 511U) - 255;
            }

            OfficeVp8Encoder.ForwardTransform(residual, rows, coefficients);
            for (int index = 0; index < 16; index++) {
                quantized[index] = (int)coefficients[index];
            }

            int[] reconstructed = OfficeVp8Transform.InverseTransform4x4(quantized);
            for (int index = 0; index < 16; index++) {
                worst = Math.Max(worst, Math.Abs(reconstructed[index] - residual[index]));
            }
        }

        Assert.True(worst <= 1, $"Largest reconstruction error was {worst}.");
    }

    [Fact]
    public void FlatBlockProducesOnlyTheDcCoefficient() {
        var residual = Enumerable.Repeat(-37, 16).ToArray();
        var coefficients = new double[16];

        OfficeVp8Encoder.ForwardTransform(residual, new double[16], coefficients);

        Assert.Equal(-37 * 8D, coefficients[0]);
        Assert.All(coefficients.Skip(1), value => Assert.Equal(0D, value));
    }
}
