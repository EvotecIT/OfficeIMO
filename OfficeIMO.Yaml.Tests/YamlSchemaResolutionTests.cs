using System.Globalization;
using Xunit;

namespace OfficeIMO.Yaml.Tests;

// The resolution tables of YAML 1.2.2 section 10 and https://yaml.org/type/. The suite checks events, not types,
// so nothing else pins which spellings resolve to which type under each schema.
public class YamlSchemaResolutionTests {
    private static YamlReader ScalarReader(string yaml, YamlSchema? schema = null) {
        var reader = new YamlReader(yaml, new YamlReaderOptions { Schema = schema });
        while (reader.Read() && reader.TokenType != YamlTokenType.Scalar) { }
        return reader;
    }

    private static string Kind(string scalar, YamlSchema? schema) {
        using YamlReader reader = ScalarReader(scalar, schema);
        if (reader.IsNull()) return "null";
        if (reader.TryGetBoolean(out bool b)) return "bool:" + (b ? "true" : "false");
        if (reader.TryGetInt64(out long i)) return "int:" + i.ToString(CultureInfo.InvariantCulture);
        if (reader.TryGetDouble(out double d)) return "float:" + d.ToString("R", CultureInfo.InvariantCulture);
        return "str";
    }

    [Theory]
    // Core schema (the default).
    [InlineData("null", null, "null")]
    [InlineData("Null", null, "null")]
    [InlineData("NULL", null, "null")]
    [InlineData("~", null, "null")]
    [InlineData("nULL", null, "str")]
    [InlineData("True", null, "bool:true")]
    [InlineData("FALSE", null, "bool:false")]
    [InlineData("yes", null, "str")]
    [InlineData("on", null, "str")]
    [InlineData("+12", null, "int:12")]
    [InlineData("-0", null, "int:0")]
    [InlineData("0o17", null, "int:15")]
    [InlineData("0x1F", null, "int:31")]
    [InlineData("0b11", null, "str")]
    [InlineData("012", null, "int:12")]
    [InlineData("1_000", null, "str")]
    [InlineData("1.5", null, "float:1.5")]
    [InlineData(".5", null, "float:0.5")]
    [InlineData("1.", null, "float:1")]
    [InlineData("1e3", null, "float:1000")]
    [InlineData("-.inf", null, "float:-Infinity")]
    [InlineData(".NaN", null, "float:NaN")]
    [InlineData(".nan.", null, "str")]
    [InlineData("1:20", null, "str")]
    // JSON schema.
    [InlineData("null", YamlSchema.Json, "null")]
    [InlineData("~", YamlSchema.Json, "str")]
    [InlineData("True", YamlSchema.Json, "str")]
    [InlineData("-12", YamlSchema.Json, "int:-12")]
    [InlineData("+12", YamlSchema.Json, "str")]
    [InlineData("012", YamlSchema.Json, "str")]
    [InlineData("1.5e-3", YamlSchema.Json, "float:0.0015")]
    [InlineData(".inf", YamlSchema.Json, "str")]
    // Failsafe schema.
    [InlineData("null", YamlSchema.Failsafe, "str")]
    [InlineData("true", YamlSchema.Failsafe, "str")]
    [InlineData("12", YamlSchema.Failsafe, "str")]
    // YAML 1.1 types.
    [InlineData("y", YamlSchema.Yaml11, "bool:true")]
    [InlineData("NO", YamlSchema.Yaml11, "bool:false")]
    [InlineData("On", YamlSchema.Yaml11, "bool:true")]
    [InlineData("off", YamlSchema.Yaml11, "bool:false")]
    [InlineData("oN", YamlSchema.Yaml11, "str")]
    [InlineData("0b1010", YamlSchema.Yaml11, "int:10")]
    [InlineData("017", YamlSchema.Yaml11, "int:15")]
    [InlineData("08", YamlSchema.Yaml11, "str")]
    [InlineData("-0x_1F", YamlSchema.Yaml11, "int:-31")]
    [InlineData("1_000", YamlSchema.Yaml11, "int:1000")]
    [InlineData("190:20:30", YamlSchema.Yaml11, "int:685230")]
    [InlineData("1:60", YamlSchema.Yaml11, "str")]
    [InlineData("0o17", YamlSchema.Yaml11, "str")]
    [InlineData("1e3", YamlSchema.Yaml11, "str")]
    [InlineData("1.0e+3", YamlSchema.Yaml11, "float:1000")]
    [InlineData("1.0e3", YamlSchema.Yaml11, "str")]
    [InlineData("190:20:30.15", YamlSchema.Yaml11, "float:685230.15")]
    [InlineData("1.29.2", YamlSchema.Yaml11, "str")]
    [InlineData("2001-12-14", YamlSchema.Yaml11, "str")]
    public void PlainScalarsResolvePerSchema(string scalar, YamlSchema? schema, string expected) {
        Assert.Equal(expected, Kind(scalar, schema));
    }

    [Fact]
    public void Yaml11DirectiveSelectsTheYaml11SchemaUnlessForced() {
        Assert.Equal("bool:true", Kind("%YAML 1.1\n--- yes", null));
        Assert.Equal("str", Kind("%YAML 1.1\n--- yes", YamlSchema.Core));
        Assert.Equal("str", Kind("%YAML 1.2\n--- yes", null));
    }

    [Theory]
    [InlineData("'12'", "str")]
    [InlineData("\"true\"", "str")]
    [InlineData("! 12", "str")]
    [InlineData("!!int '12'", "int:12")]
    [InlineData("!!float 12", "float:12")]
    [InlineData("!!str 12", "str")]
    [InlineData("!custom 12", "str")]
    [InlineData("!!null ''", "null")]
    public void TagsAndQuotingOverrideImplicitResolution(string scalar, string expected) {
        Assert.Equal(expected, Kind(scalar, null));
    }

    [Fact]
    public void IntegersOutsideTheInt64RangeAreDoublesAndDecimalsOnly() {
        using YamlReader reader = ScalarReader("99999999999999999999");
        Assert.False(reader.TryGetInt64(out _));
        Assert.True(reader.TryGetDouble(out double d));
        Assert.Equal(1e20, d);
        Assert.True(reader.TryGetDecimal(out decimal m));
        Assert.Equal(99999999999999999999m, m);
    }

    [Fact]
    public void DecimalsKeepFractionDigitsAndRejectSpecialValues() {
        using (YamlReader reader = ScalarReader("0.1")) {
            Assert.True(reader.TryGetDecimal(out decimal m));
            Assert.Equal(0.1m, m);
        }
        using (YamlReader reader = ScalarReader(".inf")) Assert.False(reader.TryGetDecimal(out _));
    }

    // Short decimals skip decimal.Parse. Every spelling of up to five digits, points and signs, and long digit runs at
    // the length limit, must give the parser's exact bits (value, sign and scale) or defer to it.
    [Fact]
    public void ShortDecimalPathMatchesDecimalParse() {
        const string alphabet = "0123456789.-+";
        var buffer = new char[24];
        int matched = 0;
        void Check(int length) {
            string text = new string(buffer, 0, length);
            if (!YamlSchemaResolver.TryParseShortDecimal(text.AsSpan(), out decimal fast)) return;
            Assert.True(decimal.TryParse(text, NumberStyles.Float, CultureInfo.InvariantCulture, out decimal expected), text);
            Assert.Equal(decimal.GetBits(expected), decimal.GetBits(fast));
            matched++;
        }
        void Sweep(int length, int position) {
            if (position == length) {
                Check(length);
                return;
            }
            foreach (char c in alphabet) {
                buffer[position] = c;
                Sweep(length, position + 1);
            }
        }
        for (int length = 1; length <= 5; length++) Sweep(length, 0);
        var random = new Random(20261110);
        for (int i = 0; i < 20_000; i++) {
            int length = random.Next(15, 21);
            for (int j = 0; j < length; j++) buffer[j] = (char)('0' + random.Next(10));
            if (random.Next(2) == 0) buffer[random.Next(1, length - 1)] = '.';
            if (random.Next(3) == 0) buffer[0] = '-';
            Check(length);
        }
        Assert.True(matched > 100_000);
    }

    [Fact]
    public void Yaml11TimestampsAreNotNumbers() {
        using YamlReader reader = ScalarReader("2001-12-14 21:59:43.10 -5", YamlSchema.Yaml11);
        Assert.False(reader.TryGetDouble(out _));
        Assert.False(reader.IsNull());
        Assert.Equal("2001-12-14 21:59:43.10 -5", reader.GetString());
    }
}
