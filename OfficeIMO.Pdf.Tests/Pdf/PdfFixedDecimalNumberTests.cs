using System.Globalization;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public class PdfFixedDecimalNumberTests {
    // Content-stream numbers must match the runtime's own "0.###"-style formatting, so generated PDFs stay
    // byte-identical whichever path writes them. Values cover exact short decimals (the integer path), values
    // with more digits than kept (rounding by the framework), signs, zero, large magnitudes and non-decimals.
    [Theory]
    [InlineData(3)]
    [InlineData(5)]
    [InlineData(6)]
    public void FixedDecimalsMatchesTheFrameworkFormat(int decimals) {
        string format = "0." + new string('#', decimals);
        var values = new List<double> {
            0D, 1D, -1D, 0.5D, -0.5D, 0.001D, -0.001D, 0.0005D, -0.0005D, 0.0004999D, 12.5D, 72D, 595.276D, 841.89D,
            1D / 3D, 2D / 3D, 0.1D + 0.2D, 1E-7D, 123456789.123D, 999999999999.999D, 1E12D, -1E12D, 1E15D, 1E20D,
            double.Epsilon, double.MaxValue, -double.MaxValue
        };
        var random = new Random(20261114);
        double scale = Math.Pow(10D, decimals);
        for (int index = 0; index < 20000; index++) {
            values.Add(random.Next(-2000000, 2000000) / scale);
            values.Add(Math.Round((random.NextDouble() - 0.5D) * 2000D, decimals, MidpointRounding.AwayFromZero));
            values.Add((random.NextDouble() - 0.5D) * 2000D);
            values.Add(random.Next(0, 256) / 255D);
        }

        foreach (double value in values) {
            Assert.Equal(value.ToString(format, CultureInfo.InvariantCulture), PdfNumberFormatter.FixedDecimals(value, decimals));
        }
    }
}
