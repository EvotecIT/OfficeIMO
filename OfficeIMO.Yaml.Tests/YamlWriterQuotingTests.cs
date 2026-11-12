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

    // The quoting decision skips the resolver for text that starts like a word or like a number. Every string of up to
    // three characters over the characters those shortcuts examine, every four-character string over the ones that
    // shape numbers, and each one-character change of longer numbers and timestamps must get the full resolution's answer.
    [Fact]
    public void QuotingShortcutsAgreeWithFullResolution() {
        const string all = "0123456789+-._,: \tabcdefABCDEFxXoOtTZzinINmMyYsSlLuUrRhH~<=#!";
        const string numeric = "019+-._,: xobeEZTinm";
        var buffer = new char[32];
        var failures = new List<string>();
        int checkedCount = 0;
        void Check(int length) {
            ReadOnlySpan<char> text = buffer.AsSpan(0, length);
            bool expected = YamlScalarFormat.IsSchemaAmbiguousResolved(text);
            if (YamlScalarFormat.IsSchemaAmbiguous(text) != expected && failures.Count < 20) {
                failures.Add("'" + text.ToString() + "' should " + (expected ? "" : "not ") + "need quotes.");
            }
            checkedCount++;
        }
        void Sweep(string alphabet, int length, int position) {
            if (position == length) {
                Check(length);
                return;
            }
            foreach (char c in alphabet) {
                buffer[position] = c;
                Sweep(alphabet, length, position + 1);
            }
        }
        for (int length = 1; length <= 3; length++) Sweep(all, length, 0);
        Sweep(numeric, 4, 0);
        string[] longer = {
            "1_000", "0x1F_FF", "0o17", "0b1010_0111", "1,000.5", "+.inf", "-.NaN", "1.29.2", "100m", "128Mi", "1e+10", "1.5E-3",
            "190:20:30.15", "-1:20", ".5", "+12.5e-07", "2001-12-14", "2001-12-14t21:59:43.10-05:00", "2001-12-14 21:59:43.10 Z",
            "2001-12-14T21:59:43Z", "10.0.0.1", "8080", "007", "1__2", "+0x_1", "1:2:3:4"
        };
        foreach (string value in longer) {
            for (int i = 0; i <= value.Length; i++) {
                foreach (char c in all) {
                    value.CopyTo(0, buffer, 0, value.Length);
                    if (i < value.Length) {
                        buffer[i] = c;
                        Check(value.Length);
                    } else {
                        buffer[i] = c;
                        Check(value.Length + 1);
                    }
                }
            }
        }
        Assert.True(failures.Count == 0, string.Join(Environment.NewLine, failures));
        Assert.True(checkedCount > 400_000);
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
