#if NET8_0_OR_GREATER
using System.Globalization;
using OfficeIMO.CSV;
using Xunit;

namespace OfficeIMO.CSV.Tests;

/// <summary>Double text written by the CSV writers must equal the runtime's invariant text for every value.</summary>
public class CsvDoubleTextTests
{
    [Fact]
    public void InvariantDoubleTextMatchesTheRuntime()
    {
        var random = new Random(20261111);
        var values = new List<double> {
            0.0, -0.0, double.NaN, double.PositiveInfinity, double.NegativeInfinity, double.Epsilon, double.MaxValue, double.MinValue,
            0.1, 0.2 + 0.1, 0.001, 0.00099999, 0.0001, 1e-5, 999_999_999_999_999.0, 99_999_999_999_999.9, 1e15, 1e15 - 1, 1e14,
            123_456_789.1234567, 0.1234567, 0.12345678, 1.5, -1.5, 2.5, 1234.5678, -0.0123, 4_503_599_627_370_496.5, 9_007_199_254_740_993
        };
        for (var sample = 0; sample < 200_000; sample++)
        {
            int places = random.Next(0, 10);
            int digits = random.Next(1, 18);
            decimal mantissa = 0;
            for (var digit = 0; digit < digits; digit++) mantissa = mantissa * 10 + random.Next(10);
            double value = (double)mantissa / Math.Pow(10, places) * Math.Pow(10, random.Next(-6, 7));
            values.Add(random.Next(2) == 0 ? value : -value);
            values.Add(BitConverter.Int64BitsToDouble(((long)random.Next() << 32) | (uint)random.Next()));
        }

        Span<char> text = stackalloc char[64];
        foreach (double value in values)
        {
            Assert.True(CsvWriter.TryFormatInvariant(value, text, out int written));
            Assert.Equal(value.ToString(CultureInfo.InvariantCulture), text[..written].ToString());
        }
    }
}
#endif
