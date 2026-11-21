using System.Text;
using BenchmarkDotNet.Attributes;
using OfficeIMO.Yaml.Benchmarks.Adapters;
using OfficeIMO.Yaml.Tests.Conformance;
using YamlDotNet.RepresentationModel;

namespace OfficeIMO.Yaml.Benchmarks;

using OfficeIMOYamlDocument = OfficeIMO.Yaml.YamlDocument;
// Inside the namespace so YamlDotNet's tree types win over the OfficeIMO.Yaml model types of the same name.
using YamlNode = YamlDotNet.RepresentationModel.YamlNode;

/// <summary>
/// Merge a staging overlay into an umbrella chart's values file and save it: OfficeIMO.Yaml's minimal-diff merge
/// against YamlDotNet loading both files, merging the representation trees by Helm's rules and serializing the result.
/// </summary>
/// <remarks>
/// The base nests the committed Bitnami nginx values under eight subchart keys (about 400 KB, as an umbrella chart's
/// values do) and the overlay repeats the committed staging overlay under each. Setup checks that both outputs load to
/// the data an independent Helm-rules merge computes.
/// </remarks>
[MemoryDiagnoser]
public class YamlMergeBenchmarks {
    private const int Subcharts = 8;
    private static readonly UTF8Encoding Utf8 = new(false);
    private string _base = string.Empty;
    private string _overlay = string.Empty;

    [GlobalSetup]
    public void Setup() {
        YamlMergePair pair = YamlMergeCorpus.Load().Single(p => p.Id == "helm-nginx-staging");
        _base = Umbrella(Utf8.GetString(pair.Base));
        _overlay = Umbrella(Utf8.GetString(pair.Overlay));
        Validate(MergedBytes(), "OfficeIMO.Yaml");
        Validate(Utf8.GetBytes(YamlDotNetText()), "YamlDotNet");
    }

    [Benchmark(Baseline = true)]
    public long YamlDotNet() => YamlDotNetText().Length;

    [Benchmark]
    public long OfficeIMOYaml() => MergedBytes().Length;

    private string YamlDotNetText() {
        var document = new YamlStream();
        document.Load(new StringReader(_base));
        var overlay = new YamlStream();
        overlay.Load(new StringReader(_overlay));
        Merge((YamlMappingNode)document.Documents[0].RootNode, (YamlMappingNode)overlay.Documents[0].RootNode);
        var writer = new StringWriter();
        document.Save(writer, assignAnchors: false);
        return writer.ToString();
    }

    // Helm's mergeMaps on YamlDotNet's tree: maps merge recursively, everything else replaces.
    private static void Merge(YamlMappingNode target, YamlMappingNode overlay) {
        foreach (KeyValuePair<YamlNode, YamlNode> entry in overlay.Children) {
            if (target.Children.TryGetValue(entry.Key, out YamlNode? current) && current is YamlMappingNode mapping && entry.Value is YamlMappingNode overlayMapping) Merge(mapping, overlayMapping);
            else target.Children[entry.Key] = entry.Value;
        }
    }

    private void Validate(byte[] output, string lane) {
        var reader = new YamlDotNetConformanceEngine();
        IReadOnlyList<YamlConformanceNode> expected = YamlHelmMerge.Merge(reader.Load(Utf8.GetBytes(_base)), reader.Load(Utf8.GetBytes(_overlay)), YamlHelmSequenceRule.Replace, null, nullRemovesKey: false);
        string? difference = YamlJsonEquivalence.FindDifference(expected, reader.Load(output));
        if (difference != null) throw new InvalidOperationException(lane + " merged to different data: " + difference);
    }

    private byte[] MergedBytes() {
        using OfficeIMOYamlDocument document = OfficeIMOYamlDocument.Parse(_base);
        using OfficeIMOYamlDocument overlay = OfficeIMOYamlDocument.Parse(_overlay);
        document.Merge(overlay);
        var output = new MemoryStream();
        document.Save(output);
        return output.ToArray();
    }

    // Nests a values file under one key per subchart, indenting every line by two spaces.
    private static string Umbrella(string values) {
        var text = new StringBuilder(values.Length * Subcharts + 64);
        string[] lines = values.Replace("\r\n", "\n").Split('\n');
        for (int i = 1; i <= Subcharts; i++) {
            text.Append("web").Append(i).Append(":\n");
            foreach (string line in lines) {
                if (line.Length > 0) text.Append("  ").Append(line);
                text.Append('\n');
            }
        }

        return text.ToString();
    }
}
