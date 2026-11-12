#if NET8_0_OR_GREATER
using System.Globalization;
using OfficeIMO.CSV;
using Xunit;

namespace OfficeIMO.CSV.Tests;

/// <summary>Number and date text written by the CSV writers must equal the runtime's text for every value.</summary>
public class CsvValueTextTests
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
        for (int power = -1074; power <= 1023; power++)
        {
            double two = Math.ScaleB(1, power);
            values.Add(two);
            values.Add(Math.BitIncrement(two));
            values.Add(Math.BitDecrement(two));
        }
        for (int power = -323; power <= 308; power++)
        {
            double ten = double.Parse("1e" + power, CultureInfo.InvariantCulture);
            values.Add(ten);
            values.Add(-Math.BitIncrement(ten));
            values.Add(Math.BitDecrement(ten));
        }
        for (var sample = 0; sample < 200_000; sample++)
        {
            values.Add(BitConverter.Int64BitsToDouble(((long)random.Next() << 32) | (uint)random.Next()) * (random.Next(2) == 0 ? 1 : -1));
            values.Add(BitConverter.Int64BitsToDouble(random.NextInt64(1, 1L << 52)));
            values.Add(random.NextDouble());
        }
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

    [Fact]
    public void InvariantDecimalTextMatchesTheRuntime()
    {
        var random = new Random(20261112);
        var values = new List<decimal> {
            0m, -0m, 0.000m, new decimal(0, 0, 0, true, 3), 1m, -1m, 0.5m, 1.50m, -0.0050m, decimal.MaxValue, decimal.MinValue,
            new decimal(-1, -1, 0, false, 0), new decimal(-1, -1, 0, true, 28), new decimal(1, 0, 0, false, 28), new decimal(0, 0, 1, false, 5),
            79_228_162_514_264_337_593_543_950_335m, 0.0000000000000000000000000001m, 12345678901234567890m, 1844674407370955161.5m
        };
        for (var sample = 0; sample < 200_000; sample++)
        {
            int low = random.Next(int.MinValue, int.MaxValue), middle = random.Next(int.MinValue, int.MaxValue);
            int high = random.Next(4) == 0 ? random.Next() : 0;
            if (random.Next(3) == 0) middle = 0;
            values.Add(new decimal(low, middle, high, random.Next(2) == 0, (byte)random.Next(0, 29)));
            values.Add(random.Next(0, 10_000_000) / 100m - 50_000m);
        }

        Span<char> text = stackalloc char[64];
        foreach (decimal value in values)
        {
            Assert.True(CsvWriter.TryFormatInvariant(value, text, out int written));
            Assert.Equal(value.ToString(CultureInfo.InvariantCulture), text[..written].ToString());
        }
    }
    [Fact]
    public void InvariantIntegerTextMatchesTheRuntime()
    {
        var values = new List<long> { 0, 1, -1, 9, -9, 10, -10, int.MaxValue, int.MinValue, long.MaxValue, long.MinValue, long.MinValue + 1 };
        for (long power = 1; power <= 1_000_000_000_000_000_000; power *= 10)
        {
            values.AddRange(new[] { power - 1, power, power + 1, -(power - 1), -power, -(power + 1) });
        }
        var random = new Random(20261114);
        for (var sample = 0; sample < 100_000; sample++)
        {
            values.Add(random.NextInt64(long.MinValue, long.MaxValue) >> random.Next(0, 63));
        }

        Span<char> text = stackalloc char[64];
        foreach (long value in values)
        {
            Assert.True(CsvWriter.TryFormatInvariant(value, text, out int written));
            Assert.Equal(value.ToString(CultureInfo.InvariantCulture), text[..written].ToString());
            if (value is >= int.MinValue and <= int.MaxValue)
            {
                Assert.True(CsvWriter.TryFormatInvariant((int)value, text, out written));
                Assert.Equal(((int)value).ToString(CultureInfo.InvariantCulture), text[..written].ToString());
            }
        }
    }

    [Theory]
    [InlineData("")]
    [InlineData("de-DE")]
    public void FormattedDateTextMatchesTheRuntime(string cultureName)
    {
        CultureInfo culture = cultureName.Length == 0 ? CultureInfo.InvariantCulture : CultureInfo.GetCultureInfo(cultureName);
        string[] formats = {
            "yyyy-MM-ddTHH:mm:ss", "yyyy-MM-dd HH:mm:ss.fff", "yyyy/MM/dd", "dd.MM.yyyy", "HH:mm", "yyyyMMdd_HHmmss", "yyyy-MM-ddTHH:mm:ss.fffffff",
            "ss.f", "mm:ss.ffffff", "--yyyy--", "yyyy-MM-dd'T'HH:mm:ss", "yyyy\\-MM\\Tdd", "\"at\" HH:mm", "'un''closed", "'a\\b'yyyy", "yyyy\\", "'unclosed", "yyyy-MM-ddTHH:mm:ssK", "yy-M-d", "yyyyy", "MMM dd", "hh:mm tt", "'T'yyyy", "dddd", "fffffffff", "s", "O", "d"
        };
        var random = new Random(20261113);
        var values = new List<DateTime> { DateTime.MinValue, DateTime.MaxValue, new DateTime(999, 1, 2, 3, 4, 5, 6, DateTimeKind.Utc), new DateTime(2024, 2, 29, 23, 59, 59, 999).AddTicks(9999) };
        for (int sample = 0; sample < 20_000; sample++)
        {
            values.Add(new DateTime(random.NextInt64(DateTime.MinValue.Ticks, DateTime.MaxValue.Ticks), (DateTimeKind)random.Next(3)));
        }

        Span<char> text = stackalloc char[128];
        foreach (string format in formats)
        {
            foreach (DateTime value in values)
            {
                string expected;
                try { expected = value.ToString(format, culture); }
                catch (FormatException) { Assert.Throws<FormatException>(() => CsvWriter.TryFormatDate(value, new char[128], out _, format, culture)); continue; }
                Assert.True(CsvWriter.TryFormatDate(value, text, out int written, format, culture));
                Assert.Equal(expected, text[..written].ToString());
            }
        }
    }
}
#endif