using System.Globalization;

namespace OfficeIMO.Yaml.Tests.Conformance;

/// <summary>
/// Compares loaded YAML documents with the suite's <c>in.json</c> using JSON data-model equality:
/// objects are unordered, numbers compare by value whatever their spelling, and integer versus float is not distinguished.
/// Non-string YAML keys are compared through their JSON key text: scalar text for numbers and booleans, <c>null</c> for null.
/// </summary>
public static class YamlJsonEquivalence {
    /// <summary>Compares every document. Returns null when equivalent, or a description of the first difference.</summary>
    public static string? FindDifference(IReadOnlyList<YamlConformanceNode> expected, IReadOnlyList<YamlConformanceNode> actual) {
        if (expected == null) throw new ArgumentNullException(nameof(expected));
        if (actual == null) throw new ArgumentNullException(nameof(actual));
        if (expected.Count != actual.Count) return "expected " + expected.Count + " document(s) but loaded " + actual.Count;
        for (int i = 0; i < expected.Count; i++) {
            string? difference = FindDifference(expected[i], actual[i], "$" + (expected.Count > 1 ? "[doc " + i + "]" : string.Empty));
            if (difference != null) return difference;
        }

        return null;
    }

    private static string? FindDifference(YamlConformanceNode expected, YamlConformanceNode actual, string path) {
        switch (expected.Kind) {
            case YamlConformanceNodeKind.Null:
            case YamlConformanceNodeKind.Boolean:
            case YamlConformanceNodeKind.String:
                return actual.Kind == expected.Kind && string.Equals(expected.Scalar, actual.Scalar, StringComparison.Ordinal)
                    ? null
                    : Mismatch(path, expected, actual);
            case YamlConformanceNodeKind.Integer:
            case YamlConformanceNodeKind.Float:
                return IsNumber(actual) && NumbersEqual(expected.Scalar!, actual.Scalar!) ? null : Mismatch(path, expected, actual);
            case YamlConformanceNodeKind.Sequence:
                if (actual.Kind != YamlConformanceNodeKind.Sequence) return Mismatch(path, expected, actual);
                if (expected.Items.Count != actual.Items.Count) return path + ": expected " + expected.Items.Count + " items but found " + actual.Items.Count;
                for (int i = 0; i < expected.Items.Count; i++) {
                    string? difference = FindDifference(expected.Items[i], actual.Items[i], path + "[" + i + "]");
                    if (difference != null) return difference;
                }

                return null;
            case YamlConformanceNodeKind.Mapping:
                return FindMappingDifference(expected, actual, path);
            default:
                throw new ArgumentOutOfRangeException(nameof(expected), expected.Kind, "Unknown node kind.");
        }
    }

    private static string? FindMappingDifference(YamlConformanceNode expected, YamlConformanceNode actual, string path) {
        if (actual.Kind != YamlConformanceNodeKind.Mapping) return Mismatch(path, expected, actual);
        var loaded = new Dictionary<string, YamlConformanceNode>(StringComparer.Ordinal);
        foreach (KeyValuePair<YamlConformanceNode, YamlConformanceNode> entry in actual.Entries) {
            string? key = JsonKey(entry.Key);
            if (key == null) return path + ": loaded a " + entry.Key.Kind + " key that JSON cannot represent";
            if (loaded.ContainsKey(key)) return path + ": loaded duplicate key \"" + key + "\"";
            loaded.Add(key, entry.Value);
        }

        if (expected.Entries.Count != loaded.Count) return path + ": expected " + expected.Entries.Count + " keys but found " + loaded.Count;
        foreach (KeyValuePair<YamlConformanceNode, YamlConformanceNode> entry in expected.Entries) {
            string key = JsonKey(entry.Key) ?? throw new ArgumentException(path + ": expected value has a " + entry.Key.Kind + " key.", nameof(expected));
            if (!loaded.TryGetValue(key, out YamlConformanceNode? value)) return path + ": missing key \"" + key + "\"";
            string? difference = FindDifference(entry.Value, value, path + "." + key);
            if (difference != null) return difference;
        }

        return null;
    }

    private static string? JsonKey(YamlConformanceNode key) => key.Kind switch {
        YamlConformanceNodeKind.String or YamlConformanceNodeKind.Boolean or YamlConformanceNodeKind.Integer or YamlConformanceNodeKind.Float => key.Scalar,
        YamlConformanceNodeKind.Null => "null",
        _ => null
    };

    private static bool IsNumber(YamlConformanceNode node) =>
        node.Kind == YamlConformanceNodeKind.Integer || node.Kind == YamlConformanceNodeKind.Float;

    private static bool NumbersEqual(string expected, string actual) {
        const NumberStyles style = NumberStyles.Float;
        if (decimal.TryParse(expected, style, CultureInfo.InvariantCulture, out decimal left)
            && decimal.TryParse(actual, style, CultureInfo.InvariantCulture, out decimal right)) {
            return left == right;
        }

        if (double.TryParse(expected, style, CultureInfo.InvariantCulture, out double a)
            && double.TryParse(actual, style, CultureInfo.InvariantCulture, out double b)) {
            return a.Equals(b);
        }

        return string.Equals(expected, actual, StringComparison.Ordinal);
    }

    private static string Mismatch(string path, YamlConformanceNode expected, YamlConformanceNode actual) =>
        path + ": expected " + expected.Kind + " " + expected + " but loaded " + actual.Kind + " " + actual;
}
