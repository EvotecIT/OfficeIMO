using System.Globalization;
using System.Text.Json;
using OfficeIMO.Data.Conversions;
using OfficeIMO.Toml;
using OfficeIMO.Toml.Tests.Conformance;
using Xunit;
using TaggedJson = OfficeIMO.Toml.Tests.Conformance.JsonNode;

namespace OfficeIMO.Data.Conversions.Tests;

// toml-test is the oracle in both directions: every valid case converts TOML -> JSON to the case's tagged JSON, and
// TOML -> TOML through the TOML writer re-reads to the same tagged JSON, so the sink keeps every type, including the four
// date-time kinds. Known gaps belong in Exceptions with a reason; the list only shrinks.
public class TomlOracleTests {
    private static readonly IReadOnlyDictionary<string, string> Exceptions = new Dictionary<string, string>(StringComparer.Ordinal);

    public static TheoryData<string> ValidCases() {
        var data = new TheoryData<string>();
        foreach (TomlTestCase testCase in TomlTestSuite.Valid) data.Add(testCase.Name);
        return data;
    }

    [Theory]
    [MemberData(nameof(ValidCases))]
    public void TomlToJsonMatchesTheTaggedJson(string name) {
        TomlTestCase testCase = TomlTestSuite.Valid.Single(c => c.Name == name);
        TaggedJson expected = JsonText.Parse(testCase.Json!);
        (string json, _) = Convert.Run(testCase.Toml, DataFormat.Toml, DataFormat.Json);
        using JsonDocument actual = JsonDocument.Parse(json);
        Check(name, Compare(actual.RootElement, expected, "$"), json);
    }

    [Theory]
    [MemberData(nameof(ValidCases))]
    public void TomlWrittenByTheSinkReadsBackToTheTaggedJson(string name) {
        TomlTestCase testCase = TomlTestSuite.Valid.Single(c => c.Name == name);
        TaggedJson expected = JsonText.Parse(testCase.Json!);
        (string toml, OfficeConversionReport report) = Convert.Run(testCase.Toml, DataFormat.Toml, DataFormat.Toml);
        string? failure;
        try {
            using TomlDocument back = TomlDocument.Parse(toml);
            failure = TomlTestJson.Compare(back.Root, expected);
        } catch (TomlException error) {
            failure = "the output does not parse: " + error.Message;
        }

        // Moving values ahead of sub-tables is the only loss a TOML-to-TOML copy may report.
        string[] losses = report.Diagnostics.Where(d => d.LossKind != OfficeConversionLossKind.None && d.Code != OfficeDiagnosticCodes.KeyOrderChanged
            && d.Code != OfficeDiagnosticCodes.CommentsOmitted).Select(d => d.Code).ToArray();
        if (failure == null && losses.Length > 0) failure = "reported " + string.Join(", ", losses);
        Check(name, failure, toml);
    }

    private static void Check(string name, string? failure, string output) {
        if (Exceptions.TryGetValue(name, out string? reason)) {
            Assert.True(failure != null, name + " now passes; remove it from the exception list (" + reason + ").");
            return;
        }

        Assert.True(failure == null, name + ": " + failure + "\n" + output);
    }

    // JSON has no date-time type, so date-times arrive as strings; they must parse back to the same TOML value and kind.
    private static string? Compare(JsonElement actual, TaggedJson expected, string path) {
        if (IsLeaf(expected)) return CompareValue(actual, expected.GetString("type"), expected.GetString("value"), path);
        if (expected.Array != null) {
            if (actual.ValueKind != JsonValueKind.Array) return path + ": expected an array but found " + actual.ValueKind;
            if (actual.GetArrayLength() != expected.Array.Count) return path + ": expected " + expected.Array.Count + " items but found " + actual.GetArrayLength();
            int i = 0;
            foreach (JsonElement item in actual.EnumerateArray()) {
                string? difference = Compare(item, expected.Array[i], path + "[" + i + "]");
                if (difference != null) return difference;
                i++;
            }

            return null;
        }

        if (actual.ValueKind != JsonValueKind.Object) return path + ": expected an object but found " + actual.ValueKind;
        int count = 0;
        foreach (JsonProperty property in actual.EnumerateObject()) {
            count++;
            if (!expected.Object!.TryGetValue(property.Name, out TaggedJson? value)) return path + ": unexpected key '" + property.Name + "'";
            string? difference = Compare(property.Value, value, path + "." + property.Name);
            if (difference != null) return difference;
        }

        return count == expected.Object!.Count ? null : path + ": expected " + expected.Object.Count + " keys but found " + count;
    }

    private static bool IsLeaf(TaggedJson node) =>
        node.Object != null && node.Object.Count == 2 && node.Get("type")?.String != null && node.Get("value")?.String != null;

    private static string? CompareValue(JsonElement actual, string type, string expected, string path) {
        bool equal = type switch {
            "string" => actual.ValueKind == JsonValueKind.String && actual.GetString() == expected,
            "integer" => actual.ValueKind == JsonValueKind.Number && actual.GetRawText() == long.Parse(expected, CultureInfo.InvariantCulture).ToString(CultureInfo.InvariantCulture),
            "float" => FloatEquals(actual, expected),
            "bool" => actual.ValueKind is JsonValueKind.True or JsonValueKind.False && (actual.ValueKind == JsonValueKind.True ? "true" : "false") == expected,
            _ => actual.ValueKind == JsonValueKind.String && SameDateTime(actual.GetString()!, expected)
        };
        return equal ? null : path + ": expected " + type + " '" + expected + "' but found " + actual.GetRawText();
    }

    // Non-finite floats are the strings System.Text.Json's named floating-point literals read back.
    private static bool FloatEquals(JsonElement actual, string expected) {
        string text = expected.TrimStart('+');
        if (text is "nan" or "-nan") return actual.ValueKind == JsonValueKind.String && actual.GetString() == "NaN";
        if (text == "inf") return actual.ValueKind == JsonValueKind.String && actual.GetString() == "Infinity";
        if (text == "-inf") return actual.ValueKind == JsonValueKind.String && actual.GetString() == "-Infinity";
        return actual.ValueKind == JsonValueKind.Number
            && double.Parse(actual.GetRawText(), NumberStyles.Float, CultureInfo.InvariantCulture).Equals(double.Parse(text, NumberStyles.Float, CultureInfo.InvariantCulture));
    }

    private static bool SameDateTime(string actual, string expected) {
        using TomlDocument ours = TomlDocument.Parse("v = " + actual);
        using TomlDocument theirs = TomlDocument.Parse("v = " + expected);
        return ours.Root["v"].AsValue().Value is TomlDateTime left && theirs.Root["v"].AsValue().Value is TomlDateTime right
            && left.Kind == right.Kind && left.Equals(right);
    }
}
