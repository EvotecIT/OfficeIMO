#if NET8_0_OR_GREATER
using System;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using OfficeIMO.CSV;
using Xunit;

namespace OfficeIMO.CSV.Tests;

public sealed class CsvDataReaderUtf8ScalarTests {
    [Theory]
    [InlineData("", "9223372036854775807", long.MaxValue)]
    [InlineData("", "-9223372036854775808", long.MinValue)]
    [InlineData("", "9007199254740993", 9007199254740993L)]
    [InlineData("", "9.007199254740993E15", 9007199254740993L)]
    [InlineData("", "(9,007,199,254,740,993)", -9007199254740993L)]
    [InlineData("", " 42 ", 42L)]
    [InlineData("de-DE", "9.007.199.254.740.993", 9007199254740993L)]
    [InlineData("fr-FR", "9\u202f007\u202f199\u202f254\u202f740\u202f993", 9007199254740993L)]
    public void Int64KeepsExactWidthAndCulture(string cultureName, string text, long expected) {
        var options = new CsvLoadOptions { Delimiter = ';', Culture = CultureInfo.GetCultureInfo(cultureName) };
        using var stream = new MemoryStream(Encoding.UTF8.GetBytes("Value\n" + text + "\n"));
        using var reader = CsvDocument.OpenDataReader(stream, options);

        Assert.True(reader.Read());
        Assert.Equal(expected, reader.GetInt64(0));
        Assert.Equal(text, reader.GetString(0));
        Assert.Equal(typeof(string), reader.GetFieldType(0));
    }

    [Theory]
    [InlineData("9223372036854775808")]
    [InlineData("-9223372036854775809")]
    [InlineData("1.5")]
    [InlineData("not-a-number")]
    public void FailedInt64ParsingPreservesTheGetterError(string text) {
        using var stream = new MemoryStream(Encoding.UTF8.GetBytes("Value\n" + text + "\n"));
        using var reader = CsvDocument.OpenDataReader(stream);
        Assert.True(reader.Read());
        Assert.Throws<InvalidCastException>(() => reader.GetInt64(0));
        Assert.Equal(text, reader.GetString(0));
    }

    [Theory]
    [InlineData("", "1,234.5000")]
    [InlineData("de-DE", "1.234,5000")]
    [InlineData("fr-FR", "1\u202f234,5000")]
    public void FloatingAndDecimalGettersRetainCultureAndDecimalScale(string cultureName, string text) {
        var options = new CsvLoadOptions { Delimiter = ';', Culture = CultureInfo.GetCultureInfo(cultureName) };
        using var stream = new MemoryStream(Encoding.UTF8.GetBytes("Value\n" + text + "\n"));
        using var reader = CsvDocument.OpenDataReader(stream, options);
        Assert.True(reader.Read());

        Assert.Equal(1234.5d, reader.GetDouble(0));
        Assert.Equal(1234.5f, reader.GetFloat(0));
        Assert.Equal(decimal.GetBits(1234.5000m), decimal.GetBits(reader.GetDecimal(0)));
        Assert.Equal(text, reader.GetString(0));
    }

    [Theory]
    [InlineData("79228162514264337593543950335")]
    [InlineData("-79228162514264337593543950335")]
    [InlineData("0.1234567890123456789012345678")]
    [InlineData("-0.0000")]
    public void DecimalGettersRetainFullPrecisionAndSignedZero(string text) {
        using var stream = new MemoryStream(Encoding.UTF8.GetBytes("Value\n" + text + "\n"));
        using var reader = CsvDocument.OpenDataReader(stream);
        Assert.True(reader.Read());
        decimal expected = decimal.Parse(text, NumberStyles.Any, CultureInfo.InvariantCulture);
        Assert.Equal(decimal.GetBits(expected), decimal.GetBits(reader.GetDecimal(0)));
    }

    [Fact]
    public void ScalarCharacterParsersKeepUnicodeFormatsAndLongTextFallback() {
        const string format = "dd MMMM yyyy HH:mm:ss.fffffff";
        DateTime expected = new DateTime(2026, 10, 9, 8, 9, 10).AddTicks(1234567);
        Guid identifier = Guid.Parse("e08e5e09-5d59-484a-9e5e-6888fc4d3e79");
        string longPrefix = new string(' ', 129);
        string csv = "Date;Active;Guid\n09 octobre 2026 08:09:10.1234567;TrUe;" + identifier +
            "\n" + longPrefix + "2026-10-09T08:09:10.1234567Z;" + longPrefix + "false;" + longPrefix + identifier + "\n";
        using var stream = new MemoryStream(Encoding.UTF8.GetBytes(csv));
        using var reader = CsvDocument.OpenDataReader(stream, new CsvLoadOptions {
            Delimiter = ';', Culture = CultureInfo.GetCultureInfo("fr-FR"), DateTimeFormats = new[] { format },
        });

        Assert.True(reader.Read());
        Assert.Equal(expected, reader.GetDateTime(0));
        Assert.Equal(DateTimeKind.Unspecified, reader.GetDateTime(0).Kind);
        Assert.True(reader.GetBoolean(1));
        Assert.Equal(identifier, reader.GetGuid(2));
        Assert.True(reader.Read());
        Assert.Equal(expected.Ticks, reader.GetDateTime(0).Ticks);
        Assert.Equal(DateTimeKind.Utc, reader.GetDateTime(0).Kind);
        Assert.False(reader.GetBoolean(1));
        Assert.Equal(identifier, reader.GetGuid(2));
    }

    [Fact]
    public void SchemaConvertersDefaultsAndNullTokensRemainAuthoritative() {
        CsvSchema schema = new CsvSchemaBuilder()
            .Column("Id").AsType(typeof(long)).ConvertUsing(value => long.Parse((string)value!, CultureInfo.InvariantCulture) + 40)
            .Column("Amount").AsType(typeof(double)).ConvertUsing(value => double.Parse((string)value!, CultureInfo.InvariantCulture) * 2)
            .Column("Optional").AsType(typeof(long)).WithDefault(7L)
            .Done().Build();
        using var stream = new MemoryStream(Encoding.UTF8.GetBytes("Id,Amount,Optional\n2,1.5,NULL\n"));
        using var reader = CsvDocument.OpenDataReader(stream, new CsvLoadOptions { NullValue = "NULL" },
            new CsvDataReaderOptions { Schema = schema });

        Assert.True(reader.Read());
        Assert.False(reader.TryGetUtf8Text(0, out _));
        Assert.Equal(42L, reader.GetInt64(0));
        Assert.Equal(3d, reader.GetDouble(1));
        Assert.False(reader.IsDBNull(2));
        Assert.Equal(7L, reader.GetInt64(2));
    }

    [Fact]
    public void EmptyMissingMalformedAndCursorFailuresKeepExistingContracts() {
        byte[] bytes = Encoding.UTF8.GetBytes("Id,Amount\n1,\n2\n")
            .Concat(new byte[] { 0xC3, 0x28, (byte)',', (byte)'3', (byte)'\n' }).ToArray();
        using var stream = new MemoryStream(bytes);
        using var reader = CsvDocument.OpenDataReader(stream);

        Assert.Throws<InvalidOperationException>(() => reader.GetInt64(0));
        Assert.Throws<InvalidOperationException>(() => reader.GetDouble(1));
        Assert.True(reader.Read());
        Assert.Equal(1L, reader.GetInt64(0));
        Assert.False(reader.IsDBNull(1));
        Assert.Throws<InvalidCastException>(() => reader.GetDouble(1));
        Assert.True(reader.Read());
        Assert.Equal(2L, reader.GetInt64(0));
        Assert.False(reader.TryGetUtf8Text(1, out _));
        Assert.Equal(string.Empty, reader.GetString(1));
        Assert.Throws<InvalidCastException>(() => reader.GetDouble(1));
        Assert.True(reader.Read());
        Assert.False(reader.TryGetUtf8Text(0, out _));
        Assert.Equal("�(", reader.GetString(0));
        Assert.Throws<InvalidCastException>(() => reader.GetInt64(0));
        Assert.Equal(3d, reader.GetDouble(1));
        Assert.Throws<IndexOutOfRangeException>(() => reader.GetInt64(-1));
        Assert.Throws<IndexOutOfRangeException>(() => reader.GetDouble(2));
        Assert.False(reader.Read());
        Assert.Throws<InvalidOperationException>(() => reader.GetInt64(0));
        reader.Close();
        Assert.Throws<InvalidOperationException>(() => reader.GetDouble(0));
    }
}
#endif
