using System.Text;
using OfficeIMO.Yaml.Benchmarks.Adapters;

namespace OfficeIMO.Yaml.Tests.Conformance;

/// <summary>
/// Qualifies one merge of an overlay into a base file:
/// <list type="bullet">
/// <item><description>minimal diff: every line that differs from the base belongs to a path the merge report names;</description></item>
/// <item><description>data: the merged file loads, in YamlDotNet, to what <see cref="YamlHelmMerge"/> computes from the
/// YamlDotNet-loaded inputs, and SharpYaml and OfficeIMO.Yaml load it to the same data;</description></item>
/// <item><description>determinism: merging again produces the same bytes.</description></item>
/// </list>
/// </summary>
public static class YamlMergeChecks {
    private static readonly UTF8Encoding Utf8 = new(false);

    /// <summary>Runs every check and returns the failures; an empty list passes.</summary>
    public static IReadOnlyList<string> Run(byte[] baseBytes, byte[] overlayBytes, string sequences, bool nullRemovesKey, out byte[] merged) {
        var failures = new List<string>();
        YamlMergeOptions options = Options(sequences, nullRemovesKey, out YamlHelmSequenceRule rule, out string? itemKey);
        merged = Merge(baseBytes, overlayBytes, options, out YamlMergeReport report);
        byte[] again = Merge(baseBytes, overlayBytes, options, out _);
        if (!merged.AsSpan().SequenceEqual(again)) failures.Add("determinism: a second merge wrote different bytes");

        string? diff = CheckDiff(baseBytes, merged, report);
        if (diff != null) failures.Add("minimal diff: " + diff);

        var yamlDotNet = new YamlDotNetConformanceEngine();
        IReadOnlyList<YamlConformanceNode> expected;
        try {
            expected = YamlHelmMerge.Merge(yamlDotNet.Load(baseBytes), yamlDotNet.Load(overlayBytes), rule, itemKey, nullRemovesKey);
        } catch (Exception exception) {
            failures.Add("reference: YamlDotNet cannot load the inputs: " + exception.Message);
            return failures;
        }

        IYamlConformanceEngine[] readers = { yamlDotNet, new SharpYamlConformanceEngine(), new OfficeIMOYamlConformanceEngine() };
        foreach (IYamlConformanceEngine reader in readers) {
            string? difference;
            try {
                difference = YamlJsonEquivalence.FindDifference(expected, reader.Load(merged));
            } catch (Exception exception) {
                difference = exception.GetType().Name + ": " + exception.Message;
            }

            if (difference != null) failures.Add("data (" + reader.Name + "): " + difference);
        }

        return failures;
    }

    /// <summary>Reads a manifest sequence rule: <c>replace</c>, <c>append</c> or <c>mergeByKey:name</c>.</summary>
    public static YamlMergeOptions Options(string sequences, bool nullRemovesKey, out YamlHelmSequenceRule rule, out string? itemKey) {
        itemKey = null;
        var options = new YamlMergeOptions { NullRemovesKey = nullRemovesKey };
        if (sequences == "append") {
            rule = YamlHelmSequenceRule.Append;
            options.Sequences = YamlSequenceMerge.Append;
        } else if (sequences.StartsWith("mergeByKey:", StringComparison.Ordinal)) {
            rule = YamlHelmSequenceRule.MergeByKey;
            itemKey = sequences.Substring("mergeByKey:".Length);
            options.Sequences = YamlSequenceMerge.MergeByKey(itemKey);
        } else if (sequences == "replace") {
            rule = YamlHelmSequenceRule.Replace;
        } else {
            throw new InvalidDataException("Unknown sequence rule '" + sequences + "'.");
        }

        return options;
    }

    private static byte[] Merge(byte[] baseBytes, byte[] overlayBytes, YamlMergeOptions options, out YamlMergeReport report) {
        using YamlDocument document = YamlDocument.Load(new MemoryStream(baseBytes, writable: false));
        using YamlDocument overlay = YamlDocument.Load(new MemoryStream(overlayBytes, writable: false));
        report = document.Merge(overlay, options);
        var output = new MemoryStream();
        document.Save(output);
        return output.ToArray();
    }

    /// <summary>
    /// Returns null when every changed line of the merged text lies in a region the report accounts for: the lines of a
    /// replaced or removed node with its key line and the comment and blank lines around it, or the end of the collection
    /// that received an added path.
    /// </summary>
    private static string? CheckDiff(byte[] baseBytes, byte[] mergedBytes, YamlMergeReport report) {
        string baseText = Decode(baseBytes);
        string[] before = SplitLines(baseText), after = SplitLines(Decode(mergedBytes));
        using YamlDocument document = YamlDocument.Load(new MemoryStream(baseBytes, writable: false));
        var touched = new bool[before.Length + 1];
        foreach (string path in report.Replaced.Concat(report.Removed)) {
            if (Own(document, path) is { } node) Mark(touched, before, NodeFirstLine(node), NodeLastLine(node));
            else MarkEnd(touched, before, Ancestor(document, path));
        }

        foreach (string path in report.Added) MarkEnd(touched, before, Ancestor(document, path));

        foreach ((int start, int deleted, int inserted) in LineDiff.Hunks(before, after)) {
            bool ok = deleted == 0
                ? touched[start] || (start > 0 && touched[start - 1])
                : Enumerable.Range(start, deleted).All(line => touched[line]);
            if (!ok) {
                string sample = deleted > 0 ? before[start] : after.Length > 0 ? "(insertion before base line " + (start + 1) + ")" : string.Empty;
                return "line " + (start + 1) + " changed outside the reported paths (" + deleted + " deleted, " + inserted + " inserted): " + sample.Trim();
            }
        }

        return null;
    }

    private static YamlNode? Own(YamlDocument document, string path) {
        YamlNode? node = document.Select(path);
        if (node == null) return null;
        try {
            return node.GetPath() == path ? node : null;
        } catch (InvalidOperationException) {
            return null;
        }
    }

    /// <summary>The nearest node above the path that the base has as its own, or the root.</summary>
    private static YamlNode Ancestor(YamlDocument document, string path) {
        string current = path;
        while (true) {
            int cut = LastStep(current);
            current = cut <= 0 ? "$" : current.Substring(0, cut);
            if (current == "$") return document.Root!;
            if (Own(document, current) is { } node) return node;
        }
    }

    // The start of the last step: the last '.' or '[' outside a quoted key.
    private static int LastStep(string path) {
        int last = -1;
        bool quoted = false;
        for (int i = 0; i < path.Length; i++) {
            char c = path[i];
            if (quoted) {
                if (c == '\\') i++;
                else if (c == '"') quoted = false;
            } else if (c == '"') {
                quoted = true;
            } else if (c == '.' || c == '[') {
                last = i;
            }
        }

        return last;
    }

    private static int NodeFirstLine(YamlNode node) {
        YamlNode first = node;
        if (node.Parent is YamlMapping mapping) {
            foreach (YamlMappingEntry entry in mapping) {
                if (ReferenceEquals(entry.Value, node)) first = entry.Key;
            }
        }

        return first.Start.Line - 1;
    }

    private static int NodeLastLine(YamlNode node) => node.End.Column == 1 && node.End.Line > node.Start.Line ? node.End.Line - 2 : node.End.Line - 1;

    private static void Mark(bool[] touched, string[] lines, int first, int last) {
        while (first > 0 && IsTrivia(lines[first - 1])) first--;
        while (last + 1 < lines.Length && IsTrivia(lines[last + 1])) last++;
        for (int i = first; i <= last && i < touched.Length; i++) touched[i] = true;
    }

    private static void MarkEnd(bool[] touched, string[] lines, YamlNode collection) {
        int last = NodeLastLine(collection);
        touched[Math.Min(last + 1, touched.Length - 1)] = true;
        Mark(touched, lines, last, last);
    }

    private static bool IsTrivia(string line) {
        string trimmed = line.Trim();
        return trimmed.Length == 0 || trimmed[0] == '#';
    }

    private static string Decode(byte[] bytes) => OfficeIMO.Yaml.Benchmarks.Adapters.YamlTextEncoding.Decode(bytes, out _);

    private static string[] SplitLines(string text) => text.Replace("\r\n", "\n").Split('\n');
}

/// <summary>A Myers line diff that reports hunks as (base line, deleted lines, inserted lines).</summary>
internal static class LineDiff {
    public static IEnumerable<(int Start, int Deleted, int Inserted)> Hunks(string[] a, string[] b) {
        int prefix = 0;
        while (prefix < a.Length && prefix < b.Length && a[prefix] == b[prefix]) prefix++;
        int suffix = 0;
        while (suffix < a.Length - prefix && suffix < b.Length - prefix && a[a.Length - 1 - suffix] == b[b.Length - 1 - suffix]) suffix++;
        int n = a.Length - prefix - suffix, m = b.Length - prefix - suffix;
        // ops: true = keep, for each step of the shortest edit script.
        List<char> script = Script(a, b, prefix, n, m);
        var hunks = new List<(int, int, int)>();
        int x = prefix, start = -1, deleted = 0, inserted = 0;
        foreach (char op in script) {
            if (op == '=') {
                if (start >= 0) hunks.Add((start, deleted, inserted));
                start = -1;
                deleted = inserted = 0;
                x++;
            } else {
                if (start < 0) start = x;
                if (op == '-') {
                    deleted++;
                    x++;
                } else {
                    inserted++;
                }
            }
        }

        if (start >= 0) hunks.Add((start, deleted, inserted));
        return hunks;
    }

    private static List<char> Script(string[] a, string[] b, int offset, int n, int m) {
        int max = n + m;
        var trace = new List<int[]>();
        var v = new int[2 * max + 2];
        int d;
        for (d = 0; d <= max; d++) {
            trace.Add((int[])v.Clone());
            bool done = false;
            for (int k = -d; k <= d; k += 2) {
                int x = k == -d || (k != d && v[max + k - 1] < v[max + k + 1]) ? v[max + k + 1] : v[max + k - 1] + 1;
                int y = x - k;
                while (x < n && y < m && a[offset + x] == b[offset + y]) {
                    x++;
                    y++;
                }

                v[max + k] = x;
                if (x >= n && y >= m) {
                    done = true;
                    break;
                }
            }

            if (done) break;
        }

        var ops = new List<char>();
        int cx = n, cy = m;
        for (int step = d; step > 0; step--) {
            int[] vv = trace[step];
            int k = cx - cy;
            int prevK = k == -step || (k != step && vv[max + k - 1] < vv[max + k + 1]) ? k + 1 : k - 1;
            int prevX = vv[max + prevK], prevY = prevX - prevK;
            while (cx > prevX && cy > prevY) {
                ops.Add('=');
                cx--;
                cy--;
            }

            ops.Add(cx == prevX ? '+' : '-');
            cx = prevX;
            cy = prevY;
        }

        while (cx > 0 && cy > 0) {
            ops.Add('=');
            cx--;
            cy--;
        }

        ops.Reverse();
        return ops;
    }
}
