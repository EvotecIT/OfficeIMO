using OfficeIMO.Yaml.Tests.Conformance;
using SharpYaml.Syntax;

namespace OfficeIMO.Yaml.Benchmarks.Adapters;

/// <summary>
/// SharpYaml 3 syntax tree with trivia, which the library documents as reproducing the original text.
/// It offers no event or value view, so only the round-trip check applies.
/// </summary>
internal sealed class SharpYamlSyntaxConformanceEngine : IYamlConformanceEngine {
    private static readonly YamlSyntaxOptions Options = new() { IncludeTrivia = true };

    public string Name => "SharpYaml syntax tree";

    public YamlEngineFeatures Features => YamlEngineFeatures.RoundTrip | YamlEngineFeatures.Lossless;

    public IReadOnlyList<YamlTestEvent> Parse(byte[] yaml) => throw new NotSupportedException();

    public IReadOnlyList<YamlConformanceNode> Load(byte[] yaml) => throw new NotSupportedException();

    public byte[] RoundTrip(byte[] yaml) {
        string text = YamlTextEncoding.Decode(yaml, out YamlTextEncoding encoding);
        return encoding.Encode(YamlSyntaxTree.Parse(text, Options).ToFullString());
    }
}
