using Xunit;

namespace OfficeIMO.Yaml.Tests;

// Default-style strings must stay strings for YAML 1.2 core and YAML 1.1 readers, and typed scalars must read back as
// the same value under both. The suite replay requests explicit styles, so it never exercises the default choice.
public class YamlWriterQuotingTests {
    [Theory]
    [InlineData("yes")] [InlineData("No")] [InlineData("on")] [InlineData("OFF")] [InlineData("y")] [InlineData("n")]
    [InlineData("~")] [InlineData("null")] [InlineData("")] [InlineData("1e3")] [InlineData("0o17")] [InlineData("012")]
    [InlineData("0x1F")] [InlineData("1:20")] [InlineData(".inf")] [InlineData("-.NaN")] [InlineData("2001-12-14")]
    [InlineData("<<")] [InlineData("=")] [InlineData(" padded ")] [InlineData("a: b")] [InlineData("#tag")] [InlineData("- item")]
    [InlineData("1_000")] [InlineData("+12")] [InlineData("true")]
    public void DefaultStyleStringsReadAsStringsUnderCoreAndYaml11(string value) {
        string text = Write(w => {
            w.WriteStartMapping();
            w.WriteKey(value);
            w.WriteScalar(value);
            w.WriteEndMapping();
        });
        foreach (YamlSchema schema in new[] { YamlSchema.Core, YamlSchema.Yaml11 }) {
            YamlMappingEntry entry = YamlDocument.Parse(text, new YamlLoadOptions { Schema = schema }).Root!.AsMapping()[0];
            Assert.Equal(YamlScalarKind.String, entry.Key.AsScalar().Kind);
            Assert.Equal(YamlScalarKind.String, entry.Value.AsScalar().Kind);
            Assert.Equal(value, entry.Value.AsString());
        }
    }

    [Fact]
    public void PlainSafeStringsStayPlain() {
        Assert.Equal("name: web-01\n", Write(w => {
            w.WriteStartMapping();
            w.WriteKey("name");
            w.WriteScalar("web-01");
            w.WriteEndMapping();
        }));
    }

    [Theory]
    [InlineData(1.0)] [InlineData(-0.5)] [InlineData(1e20)] [InlineData(2.5e-7)] [InlineData(double.PositiveInfinity)] [InlineData(double.NegativeInfinity)]
    public void FloatsReadBackAsTheSameFloatUnderCoreAndYaml11(double value) {
        string text = Write(w => w.WriteScalar(value));
        foreach (YamlSchema schema in new[] { YamlSchema.Core, YamlSchema.Yaml11 }) {
            YamlScalar scalar = YamlDocument.Parse(text, new YamlLoadOptions { Schema = schema }).Root!.AsScalar();
            Assert.Equal(YamlScalarKind.Float, scalar.Kind);
            Assert.Equal(value, scalar.GetValue<double>());
        }
    }

    private static string Write(Action<YamlWriter> write) {
        var text = new StringWriter();
        using (var writer = new YamlWriter(text)) write(writer);
        return text.ToString();
    }
}
