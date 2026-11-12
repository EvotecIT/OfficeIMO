using System.Diagnostics;
using System.Globalization;
using System.Text;
using OfficeIMO.Yaml.Tests.Conformance;
using YamlDotNet.Core;
using YamlDotNet.RepresentationModel;

namespace OfficeIMO.Yaml.Benchmarks.LargeCorpus;

// Inside the namespace so YamlDotNet's tree types win over the OfficeIMO.Yaml model types of the same name.
using YamlDocument = YamlDotNet.RepresentationModel.YamlDocument;
using YamlNode = YamlDotNet.RepresentationModel.YamlNode;

/// <summary>
/// Merges a synthetic overlay into every Helm values file of the large corpus and runs the merge pair checks:
/// <c>--merge-corpus [--cache dir] [--output-dir dir] [--limit n]</c>. Each overlay is derived from its base with
/// YamlDotNet: every scalar leaf changes value (keeping its type), every mapping gets one new key, and sequences are
/// restated with their changed items. Writes <c>merge-corpus.md</c> and <c>merge-corpus-failures.tsv</c>.
/// </summary>
internal static class YamlMergeCorpusRun {
    private const string AddedKey = "officeimoMergeAdded";

    public static int Run(string[] args) {
        string root = YamlLargeCorpus.ResolveRoot(Option(args, "--cache"));
        string outputDirectory = Path.GetFullPath(Option(args, "--output-dir") ?? Path.Combine(root, "reports"));
        int limit = int.Parse(Option(args, "--limit") ?? int.MaxValue.ToString(CultureInfo.InvariantCulture), CultureInfo.InvariantCulture);
        List<YamlLargeCorpusFile> files = YamlLargeCorpus.Load(root).Files
            .Where(f => f.Category == "Helm" && Path.GetFileName(f.Path).StartsWith("values", StringComparison.Ordinal))
            .Take(limit).ToList();

        int merged = 0, skipped = 0, minimal = 0, reference = 0, deterministic = 0, readers = 0, crashed = 0;
        var failures = new List<string>();
        var watch = Stopwatch.StartNew();
        foreach (YamlLargeCorpusFile file in files) {
            byte[] baseBytes = file.ReadYaml();
            byte[]? overlay = Overlay(baseBytes);
            if (overlay == null) {
                // Bases YamlDotNet cannot load, or whose root is not a mapping, have no reference merge.
                skipped++;
                continue;
            }

            IReadOnlyList<string> found;
            try {
                found = YamlMergeChecks.Run(baseBytes, overlay, "replace", nullRemovesKey: false, out _);
            } catch (Exception exception) {
                crashed++;
                failures.Add(file.Name + "\tcrash\t" + OneLine(exception.GetType().Name + ": " + exception.Message));
                continue;
            }

            merged++;
            if (!found.Any(f => f.StartsWith("minimal diff", StringComparison.Ordinal))) minimal++;
            if (!found.Any(f => f.StartsWith("data (YamlDotNet)", StringComparison.Ordinal) || f.StartsWith("reference", StringComparison.Ordinal))) reference++;
            if (!found.Any(f => f.StartsWith("data (", StringComparison.Ordinal))) readers++;
            if (!found.Any(f => f.StartsWith("determinism", StringComparison.Ordinal))) deterministic++;
            foreach (string failure in found) failures.Add(file.Name + "\t" + OneLine(failure));
        }

        var report = new StringBuilder();
        report.Append("# Merge corpus\n\n");
        report.Append("Helm values files: ").Append(files.Count).Append(", merged: ").Append(merged).Append(", skipped (no reference): ").Append(skipped)
            .Append(", crashed: ").Append(crashed).Append(", ").Append(watch.Elapsed.TotalSeconds.ToString("F1", CultureInfo.InvariantCulture)).Append(" s\n\n");
        report.Append("| Check | Passed |\n| --- | ---: |\n");
        report.Append("| Minimal diff | ").Append(minimal).Append(" / ").Append(merged).Append(" |\n");
        report.Append("| Reference (Helm rules over YamlDotNet data) | ").Append(reference).Append(" / ").Append(merged).Append(" |\n");
        report.Append("| All readers (YamlDotNet, SharpYaml, OfficeIMO.Yaml) | ").Append(readers).Append(" / ").Append(merged).Append(" |\n");
        report.Append("| Determinism | ").Append(deterministic).Append(" / ").Append(merged).Append(" |\n");
        Directory.CreateDirectory(outputDirectory);
        File.WriteAllText(Path.Combine(outputDirectory, "merge-corpus.md"), report.ToString(), new UTF8Encoding(false));
        File.WriteAllLines(Path.Combine(outputDirectory, "merge-corpus-failures.tsv"), failures, new UTF8Encoding(false));
        Console.WriteLine(report);
        Console.Error.WriteLine("Failures: " + Path.Combine(outputDirectory, "merge-corpus-failures.tsv"));
        return 0;
    }

    /// <summary>Builds the overlay text, or null when YamlDotNet cannot load the base or its root is not a mapping.</summary>
    private static byte[]? Overlay(byte[] baseBytes) {
        var stream = new YamlStream();
        try {
            stream.Load(new StringReader(Adapters.YamlTextEncoding.Decode(baseBytes, out _)));
        } catch (YamlException) {
            return null;
        }

        if (stream.Documents.Count != 1 || stream.Documents[0].RootNode is not YamlMappingNode root) return null;
        var overlay = new YamlStream(new YamlDocument(Transform(root, 0)));
        var writer = new StringWriter();
        overlay.Save(writer, assignAnchors: false);
        return new UTF8Encoding(false).GetBytes(writer.ToString());
    }

    private static YamlNode Transform(YamlNode node, int depth) {
        if (depth > 200) throw new InvalidOperationException("The base is nested too deeply.");
        switch (node) {
            case YamlMappingNode mapping: {
                var result = new YamlMappingNode();
                foreach (KeyValuePair<YamlNode, YamlNode> entry in mapping.Children) {
                    if (entry.Key is YamlScalarNode { Value: "<<" }) continue;
                    result.Add(Copy(entry.Key), Transform(entry.Value, depth + 1));
                }

                if (!result.Children.ContainsKey(new YamlScalarNode(AddedKey))) result.Add(AddedKey, new YamlScalarNode("level " + depth) { Style = ScalarStyle.DoubleQuoted });
                return result;
            }
            case YamlSequenceNode sequence: {
                var result = new YamlSequenceNode();
                foreach (YamlNode item in sequence.Children) result.Add(Transform(item, depth + 1));
                return result;
            }
            default:
                return Change((YamlScalarNode)node);
        }
    }

    // A new value of the same type under the core schema, so the merge assigns rather than replaces.
    private static YamlScalarNode Change(YamlScalarNode scalar) {
        string value = scalar.Value ?? string.Empty;
        bool plain = scalar.Style is ScalarStyle.Plain or ScalarStyle.Any;
        YamlConformanceNode resolved = Adapters.YamlCoreSchema.Resolve(value, plain, scalar.Tag.IsEmpty ? null : scalar.Tag.Value);
        return resolved.Kind switch {
            YamlConformanceNodeKind.Null => new YamlScalarNode("merged") { Style = ScalarStyle.DoubleQuoted },
            YamlConformanceNodeKind.Boolean => new YamlScalarNode(resolved.Scalar == "true" ? "false" : "true") { Style = ScalarStyle.Plain },
            YamlConformanceNodeKind.Integer when long.TryParse(resolved.Scalar, NumberStyles.AllowLeadingSign, CultureInfo.InvariantCulture, out long number) && number < long.MaxValue
                => new YamlScalarNode((number + 1).ToString(CultureInfo.InvariantCulture)) { Style = ScalarStyle.Plain },
            YamlConformanceNodeKind.Float => new YamlScalarNode("2.5") { Style = ScalarStyle.Plain },
            _ => new YamlScalarNode(value + "-merged") { Style = ScalarStyle.DoubleQuoted }
        };
    }

    private static YamlNode Copy(YamlNode key) => key is YamlScalarNode scalar ? new YamlScalarNode(scalar.Value) { Style = scalar.Style, Tag = scalar.Tag } : key;

    private static string OneLine(string text) => text.Replace('\t', ' ').Replace('\r', ' ').Replace('\n', ' ');

    private static string? Option(string[] args, string name) {
        for (int i = 0; i < args.Length - 1; i++) {
            if (string.Equals(args[i], name, StringComparison.OrdinalIgnoreCase)) return args[i + 1];
        }

        return null;
    }
}
