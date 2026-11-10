using System.Text;
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

    [Theory]
    [InlineData("1.29.2", "tag: 1.29.2\n")] [InlineData("10.0.0.1", "tag: 10.0.0.1\n")]
    [InlineData("1.5", "tag: '1.5'\n")] [InlineData("yes", "tag: 'yes'\n")] [InlineData("on", "tag: 'on'\n")]
    [InlineData("012", "tag: '012'\n")] [InlineData("1:20", "tag: '1:20'\n")] [InlineData("1,000", "tag: '1,000'\n")]
    public void AssignedStringsAreQuotedOnlyWhenAYaml11ReaderInUseWouldReadAnotherType(string value, string expected) {
        // PyYAML, ruamel.yaml in 1.1 mode, Ruby Psych and go-yaml v2 read versions and addresses as strings.
        YamlDocument document = YamlDocument.Parse("tag: 1.29.1\n");
        document.Root!.AsMapping()["tag"] = value;
        Assert.Equal(expected, document.ToString());
    }

    // The writer decides quoting with shortcuts for text that starts like a word; this sweeps the shapes the shortcuts
    // separate (letters, signs, digits, dots, the null and boolean words in every casing) through both schemas.
    [Fact]
    public void ShortStringsOfEveryLeadingShapeReadBackAsStrings() {
        string[] parts = { "", "n", "N", "y", "Y", "o", "O", "t", "T", "f", "F", "a", "-", "+", ".", "0", "1", "e", "_", ":", "~", "<", "=" };
        var values = new List<string>();
        foreach (string a in parts) {
            foreach (string b in parts) {
                foreach (string c in parts) values.Add(a + b + c);
            }
        }
        foreach (string word in new[] { "null", "true", "false", "yes", "no", "on", "off", "nan", "inf" }) {
            values.Add(word);
            values.Add(char.ToUpperInvariant(word[0]) + word.Substring(1));
            values.Add(word.ToUpperInvariant());
            values.Add(word + "s");
            values.Add("-" + word);
            values.Add("-." + word);
            values.Add("." + word);
        }
        var text = new StringBuilder();
        using (var writer = new YamlWriter(new StringWriter(text))) {
            writer.WriteStartSequence();
            foreach (string value in values) writer.WriteScalar(value);
            writer.WriteEndSequence();
        }
        foreach (YamlSchema schema in new[] { YamlSchema.Core, YamlSchema.Yaml11 }) {
            YamlSequence items = YamlDocument.Parse(text.ToString(), new YamlLoadOptions { Schema = schema }).Root!.AsSequence();
            for (int i = 0; i < values.Count; i++) {
                Assert.True(items[i].AsScalar().Kind == YamlScalarKind.String && items[i].AsString() == values[i], "'" + values[i] + "' under " + schema);
            }
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
