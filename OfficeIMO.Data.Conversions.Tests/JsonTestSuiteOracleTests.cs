using System.Text.Json;
using OfficeIMO.Data.Conversions;
using OfficeIMO.Tests;
using Xunit;

namespace OfficeIMO.Data.Conversions.Tests;

// Cross-format oracle: JSONTestSuite accept cases pass through the event pipeline to the same data (numbers keep their
// exact text), and reject cases fail. Implementation-defined (i_) cases only must not crash or hang.
public class JsonTestSuiteOracleTests {
    private static readonly IReadOnlyDictionary<string, string> Exceptions = new Dictionary<string, string>(StringComparer.Ordinal) {
        ["y_object_duplicated_key.json"] = "Duplicate keys are rejected by design; YAML, TOML and XML targets cannot hold them.",
        ["y_object_duplicated_key_and_value.json"] = "Duplicate keys are rejected by design; YAML, TOML and XML targets cannot hold them."
    };

    private static string Directory => Path.Combine(RepositoryTestPaths.Find(), "OfficeIMO.TestAssets", "Json", "JSONTestSuite", "test_parsing");

    public static TheoryData<string> Cases() {
        var data = new TheoryData<string>();
        foreach (string path in System.IO.Directory.GetFiles(Directory, "*.json").OrderBy(p => p, StringComparer.Ordinal)) data.Add(Path.GetFileName(path));
        return data;
    }

    [Theory]
    [MemberData(nameof(Cases))]
    public void AcceptsValidAndRejectsInvalidJson(string name) {
        byte[] input = File.ReadAllBytes(Path.Combine(Directory, name));
        string? failure = null;
        byte[]? output = null;
        try {
            using var buffer = new MemoryStream();
            DataConversionEngine.Convert(new MemoryStream(input), DataFormat.Json, buffer, DataFormat.Json, null);
            output = buffer.ToArray();
        } catch (InvalidDataException error) {
            failure = error.Message;
        }

        if (name.StartsWith("i_", StringComparison.Ordinal)) return;
        bool shouldAccept = name.StartsWith("y_", StringComparison.Ordinal);
        if (Exceptions.TryGetValue(name, out string? reason)) {
            Assert.True(shouldAccept == (failure != null), name + " now behaves as the suite expects; remove it from the exceptions (" + reason + ").");
            return;
        }

        if (!shouldAccept) {
            Assert.True(failure != null, name + " was accepted but must be rejected.");
            return;
        }

        Assert.True(failure == null, name + ": " + failure);
        using JsonDocument expected = JsonDocument.Parse(input);
        using JsonDocument actual = JsonDocument.Parse(output);
        Assert.True(Equal(expected.RootElement, actual.RootElement), name + " changed: " + System.Text.Encoding.UTF8.GetString(output!));
    }

    // Every accepted document converts to XML that an independent parser accepts.
    [Theory]
    [MemberData(nameof(Cases))]
    public void AcceptedJsonBecomesWellFormedXml(string name) {
        if (!name.StartsWith("y_", StringComparison.Ordinal) || Exceptions.ContainsKey(name)) return;
        using var output = new MemoryStream();
        DataConversionEngine.Convert(new MemoryStream(File.ReadAllBytes(Path.Combine(Directory, name))), DataFormat.Json, output, DataFormat.Xml, null);
        output.Position = 0;
        System.Xml.Linq.XDocument.Load(output);
    }

    // Every accepted document survives JSON -> YAML -> JSON with equal data; numbers compare by value.
    [Theory]
    [MemberData(nameof(Cases))]
    public void AcceptedJsonRoundTripsThroughYaml(string name) {
        if (!name.StartsWith("y_", StringComparison.Ordinal) || Exceptions.ContainsKey(name)) return;
        byte[] input = File.ReadAllBytes(Path.Combine(Directory, name));
        using var yaml = new MemoryStream();
        DataConversionEngine.Convert(new MemoryStream(input), DataFormat.Json, yaml, DataFormat.Yaml, null);
        using var back = new MemoryStream();
        DataConversionEngine.Convert(new MemoryStream(yaml.ToArray()), DataFormat.Yaml, back, DataFormat.Json, null);
        using JsonDocument expected = JsonDocument.Parse(input);
        using JsonDocument actual = JsonDocument.Parse(back.ToArray());
        Assert.True(Equal(expected.RootElement, actual.RootElement, numbersByValue: true), name + ":\n" + System.Text.Encoding.UTF8.GetString(yaml.ToArray()));
    }

    internal static bool Equal(JsonElement left, JsonElement right, bool numbersByValue = false) {
        if (left.ValueKind != right.ValueKind) return false;
        switch (left.ValueKind) {
            case JsonValueKind.Object:
                List<JsonProperty> a = left.EnumerateObject().ToList(), b = right.EnumerateObject().ToList();
                if (a.Count != b.Count) return false;
                for (int i = 0; i < a.Count; i++) {
                    if (a[i].Name != b[i].Name || !Equal(a[i].Value, b[i].Value, numbersByValue)) return false;
                }

                return true;
            case JsonValueKind.Array:
                List<JsonElement> x = left.EnumerateArray().ToList(), y = right.EnumerateArray().ToList();
                return x.Count == y.Count && x.Zip(y, (l, r) => Equal(l, r, numbersByValue)).All(e => e);
            case JsonValueKind.String:
                return left.GetString() == right.GetString();
            case JsonValueKind.Number:
                if (left.GetRawText() == right.GetRawText()) return true;
                return numbersByValue && double.Parse(left.GetRawText(), System.Globalization.CultureInfo.InvariantCulture)
                    == double.Parse(right.GetRawText(), System.Globalization.CultureInfo.InvariantCulture);
            default:
                return true;
        }
    }
}
