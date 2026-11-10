using System.Globalization;
using OfficeIMO.Yaml.Tests.Conformance;

namespace OfficeIMO.Yaml.Benchmarks.LargeCorpus;

/// <summary>How two loaded values differ.</summary>
internal enum YamlDataDifferenceKind {
    /// <summary>Different number of documents.</summary>
    DocumentCount,
    /// <summary>A collection on one side and something else on the other, or a different size.</summary>
    Structure,
    /// <summary>Two mappings whose keys differ at the same position.</summary>
    Key,
    /// <summary>Two scalars resolved to different kinds, for example string and boolean.</summary>
    ScalarKind,
    /// <summary>Two scalars of the same kind with different values.</summary>
    ScalarValue
}

/// <summary>One difference between two loaded values.</summary>
internal sealed record YamlDataDifference(YamlDataDifferenceKind Kind, string Path, YamlConformanceNode? Left, YamlConformanceNode? Right) {
    public string Describe() => Kind + " at " + Path + ": " + Show(Left) + " vs " + Show(Right);

    /// <summary>Short scalar text for reports.</summary>
    public static string Show(YamlConformanceNode? node) {
        if (node == null) return "<none>";
        string text = node.Kind + " " + node;
        return text.Length > 60 ? text.Substring(0, 57) + "..." : text;
    }
}

/// <summary>
/// Ordered comparison of loaded documents: mapping entries in source order, numbers by value, so integers wider than
/// 64 bits (which the OfficeIMO adapter reads as doubles) still equal the same integer read exactly.
/// </summary>
internal static class YamlDataComparison {
    /// <summary>Returns the first difference, or null when equal.</summary>
    public static YamlDataDifference? First(IReadOnlyList<YamlConformanceNode> left, IReadOnlyList<YamlConformanceNode> right) {
        var found = new List<YamlDataDifference>(1);
        Collect(left, right, found, 1);
        return found.Count == 0 ? null : found[0];
    }

    /// <summary>Collects up to <paramref name="limit"/> differences, continuing past scalar differences.</summary>
    public static IReadOnlyList<YamlDataDifference> All(IReadOnlyList<YamlConformanceNode> left, IReadOnlyList<YamlConformanceNode> right, int limit) {
        var found = new List<YamlDataDifference>();
        Collect(left, right, found, limit);
        return found;
    }

    private static void Collect(IReadOnlyList<YamlConformanceNode> left, IReadOnlyList<YamlConformanceNode> right, List<YamlDataDifference> found, int limit) {
        if (left.Count != right.Count) {
            found.Add(new YamlDataDifference(YamlDataDifferenceKind.DocumentCount, "$", YamlConformanceNode.Integer(left.Count.ToString(CultureInfo.InvariantCulture)), YamlConformanceNode.Integer(right.Count.ToString(CultureInfo.InvariantCulture))));
            return;
        }

        for (int i = 0; i < left.Count && found.Count < limit; i++) Compare(left[i], right[i], left.Count > 1 ? "$[doc " + i + "]" : "$", found, limit);
    }

    private static void Compare(YamlConformanceNode left, YamlConformanceNode right, string path, List<YamlDataDifference> found, int limit) {
        if (found.Count >= limit) return;
        bool leftCollection = left.Kind is YamlConformanceNodeKind.Sequence or YamlConformanceNodeKind.Mapping;
        bool rightCollection = right.Kind is YamlConformanceNodeKind.Sequence or YamlConformanceNodeKind.Mapping;
        if (leftCollection || rightCollection) {
            if (left.Kind != right.Kind) {
                found.Add(new YamlDataDifference(YamlDataDifferenceKind.Structure, path, left, right));
                return;
            }

            if (left.Kind == YamlConformanceNodeKind.Sequence) {
                if (left.Items.Count != right.Items.Count) {
                    found.Add(new YamlDataDifference(YamlDataDifferenceKind.Structure, path, left, right));
                    return;
                }

                for (int i = 0; i < left.Items.Count; i++) Compare(left.Items[i], right.Items[i], path + "[" + i + "]", found, limit);
                return;
            }

            if (left.Entries.Count != right.Entries.Count) {
                found.Add(new YamlDataDifference(YamlDataDifferenceKind.Structure, path, left, right));
                return;
            }

            for (int i = 0; i < left.Entries.Count; i++) {
                KeyValuePair<YamlConformanceNode, YamlConformanceNode> l = left.Entries[i];
                KeyValuePair<YamlConformanceNode, YamlConformanceNode> r = right.Entries[i];
                if (!ScalarEquals(l.Key, r.Key) && !(l.Key.Kind is YamlConformanceNodeKind.Mapping or YamlConformanceNodeKind.Sequence)) {
                    found.Add(new YamlDataDifference(YamlDataDifferenceKind.Key, path + "{" + i + "}", l.Key, r.Key));
                    if (found.Count >= limit) return;
                    continue;
                }

                Compare(l.Value, r.Value, path + "." + KeyText(l.Key), found, limit);
            }

            return;
        }

        if (ScalarEquals(left, right)) return;
        bool sameKind = left.Kind == right.Kind || (IsNumber(left) && IsNumber(right));
        found.Add(new YamlDataDifference(sameKind ? YamlDataDifferenceKind.ScalarValue : YamlDataDifferenceKind.ScalarKind, path, left, right));
    }

    public static bool ScalarEquals(YamlConformanceNode left, YamlConformanceNode right) {
        if (IsNumber(left) && IsNumber(right)) return NumbersEqual(left, right);
        return left.Kind == right.Kind && string.Equals(left.Scalar, right.Scalar, StringComparison.Ordinal);
    }

    private static bool IsNumber(YamlConformanceNode node) => node.Kind is YamlConformanceNodeKind.Integer or YamlConformanceNodeKind.Float;

    private static bool NumbersEqual(YamlConformanceNode left, YamlConformanceNode right) {
        if (left.Kind == YamlConformanceNodeKind.Integer && right.Kind == YamlConformanceNodeKind.Integer) return string.Equals(left.Scalar, right.Scalar, StringComparison.Ordinal);
        // Integer versus float only matches for integers beyond 64 bits, which one side may hold as a double.
        if (left.Kind != right.Kind && !IsWideInteger(left) && !IsWideInteger(right)) return false;
        double a = double.Parse(left.Scalar!, NumberStyles.Float, CultureInfo.InvariantCulture);
        double b = double.Parse(right.Scalar!, NumberStyles.Float, CultureInfo.InvariantCulture);
        return a.Equals(b);
    }

    private static bool IsWideInteger(YamlConformanceNode node) =>
        node.Kind == YamlConformanceNodeKind.Integer && !long.TryParse(node.Scalar, NumberStyles.AllowLeadingSign, CultureInfo.InvariantCulture, out _);

    private static string KeyText(YamlConformanceNode key) {
        string text = key.Scalar ?? key.ToString();
        return text.Length > 40 ? text.Substring(0, 37) + "..." : text;
    }
}
