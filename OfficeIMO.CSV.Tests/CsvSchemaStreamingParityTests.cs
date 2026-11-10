using System.Globalization;
using System.Text;
using OfficeIMO.CSV;
using Xunit;

namespace OfficeIMO.CSV.Tests;

/// <summary>
/// A schema read from a byte stream returns the same typed values, missing values and conversion
/// failures as the same schema over decoded text.
/// </summary>
public sealed class CsvSchemaStreamingParityTests {
    private const string Input =
        "Id,Name,Price,When,Active,Note\n" +
        "1,Ada,12.50,2026-01-02,true,plain\n" +
        "2,\"Bob, Jr.\",\"7\",2026-02-03,false,\"say \"\"hi\"\"\"\n" +
        "x,Cy,abc,not-a-date,maybe,\n" +
        ",,,,,\n" +
        "5,Zoë,1e3,2026-03-04T05:06:07,1\n" +
        "6,Eve,3.25,2026-12-31,0,\"two\nlines\"\n";

    [Theory]
    [InlineData(DataMappingErrorValuePolicy.Include, false, false)]
    [InlineData(DataMappingErrorValuePolicy.Include, true, false)]
    [InlineData(DataMappingErrorValuePolicy.Redact, false, false)]
    [InlineData(DataMappingErrorValuePolicy.Include, false, true)]
    [InlineData(DataMappingErrorValuePolicy.Redact, true, true)]
    public void SchemaValuesMatchDecodedText(DataMappingErrorValuePolicy policy, bool trim, bool generalColumn) {
        Assert.Equal(Observe(reader: null, policy, trim, generalColumn), Observe(reader: Input, policy, trim, generalColumn));
    }

    [Theory]
    [InlineData(1000, null)]
    [InlineData(2, null)]
    [InlineData(1000, "x")]
    public void InferredSchemaMatchesDecodedText(int sampleRows, string? nullValue) {
        Assert.Equal(Observe(reader: null, sampleRows, nullValue), Observe(reader: Input, sampleRows, nullValue));
    }

    // Numbers parse from UTF-8 bytes on the stream and from characters on a text reader; culture
    // symbols outside ASCII (group spaces, U+2212 minus, Arabic signs) must parse alike.
    [Theory]
    [InlineData("")]
    [InlineData("en-US")]
    [InlineData("de-DE")]
    [InlineData("fr-FR")]
    [InlineData("sv-SE")]
    [InlineData("ar-SA")]
    public void CultureNumbersMatchDecodedText(string cultureName) {
        CultureInfo culture = CultureInfo.GetCultureInfo(cultureName);
        NumberFormatInfo format = culture.NumberFormat;
        string[] values = {
            "12", "-12", "+12", " 12 ", "1,234.5", "1.234,5", "1\u00a0234,5", "1\u202f234,5", "\u221212", "(12)", "12-",
            "1e3", "-0", "0x1F", "$12", "12\u00a0\u20ac", "NaN", "\u221e", "-Infinity", "300", "70000", "3000000000",
            "9999999999999999999999999999.9", "\u0661\u0662", format.NegativeSign + "5", "5" + format.NumberDecimalSeparator + "25",
            "1" + format.NumberGroupSeparator + "234", format.CurrencySymbol + "7", format.PercentSymbol + "3", "\"4\""
        };
        var builder = new CsvSchemaBuilder();
        string[] kinds = { "Int32", "Int64", "Int16", "Byte", "Double", "Decimal", "Single" };
        builder.Column(kinds[0]).AsInt32();
        builder.Column(kinds[1]).AsType(typeof(long));
        builder.Column(kinds[2]).AsType(typeof(short));
        builder.Column(kinds[3]).AsType(typeof(byte));
        builder.Column(kinds[4]).AsType(typeof(double));
        builder.Column(kinds[5]).AsType(typeof(decimal));
        builder.Column(kinds[6]).AsType(typeof(float));
        var text = new StringBuilder(string.Join(";", kinds)).Append('\n');
        foreach (string value in values) text.Append(string.Join(";", Enumerable.Repeat(value, kinds.Length))).Append('\n');
        string input = text.ToString();
        var options = new CsvReaderOptions { Schema = builder.Build(), Culture = culture, Delimiter = ';' };
        Assert.Equal(Read(input, fromBytes: false, options), Read(input, fromBytes: true, options));
    }

    private static string Observe(string? reader, int sampleRows, string? nullValue) =>
        Read(reader, new CsvReaderOptions { InferSchema = true, SchemaSampleRows = sampleRows, NullValue = nullValue });

    // A general column (a decimal with a default) converts through values; the others convert text directly.
    private static string Observe(string? reader, DataMappingErrorValuePolicy policy, bool trim, bool generalColumn) {
        var builder = new CsvSchemaBuilder();
        builder.Column("Id").AsInt32().Required();
        builder.Column("Name").AsString().Required();
        if (generalColumn) builder.Column("Price").AsType(typeof(decimal)).WithDefault(-1m);
        else builder.Column("Price").AsInt32();
        builder.Column("When").AsDateTime();
        builder.Column("Active").AsBoolean();
        builder.Column("Note").AsString();
        return Read(reader, new CsvReaderOptions {
            Schema = builder.Build(),
            MappingErrorValuePolicy = policy,
            TrimWhitespace = trim,
            Culture = CultureInfo.InvariantCulture
        });
    }

    // A null reader selects the UTF-8 byte stream; otherwise the decoded text overload.
    private static string Read(string? reader, CsvReaderOptions options) => Read(reader ?? Input, fromBytes: reader is null, options);

    private static string Read(string input, bool fromBytes, CsvReaderOptions options) {
        var result = new StringBuilder();
        try {
            using CsvDataReader data = fromBytes
                ? CsvDocument.OpenReader(new MemoryStream(Encoding.UTF8.GetBytes(input)), options)
                : CsvDocument.OpenReader(new StringReader(input), options);
            for (int column = 0; column < data.FieldCount; column++) {
                result.Append(data.GetName(column)).Append('=').Append(data.GetFieldType(column).Name).Append(' ');
            }
            result.Append('\n');
            while (true) {
                bool read;
                try {
                    read = data.Read();
                } catch (Exception exception) {
                    result.Append("read error ").Append(exception.GetType().Name).Append('\n');
                    break;
                }
                if (!read) break;
                for (int column = 0; column < data.FieldCount; column++) {
                    try {
                        object value = data.GetValue(column);
                        result.Append(value is DBNull ? "<null>" : Convert.ToString(value, CultureInfo.InvariantCulture))
                            .Append(':').Append(value.GetType().Name);
                    } catch (Exception exception) {
                        result.Append("error ").Append(exception.GetType().Name);
                    }
                    result.Append(" | ");
                }
                result.Append('\n');
            }
        } catch (Exception exception) {
            result.Append("open error ").Append(exception.GetType().Name);
        }
        return result.ToString();
    }
}
