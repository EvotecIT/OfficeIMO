using OfficeIMO.Yaml.Tests.Conformance;
using YamlDotNet.Core;
using YamlDotNet.Core.Events;
using YamlDotNet.RepresentationModel;
using DotNetScalarStyle = YamlDotNet.Core.ScalarStyle;

namespace OfficeIMO.Yaml.Benchmarks.Adapters;

/// <summary>YamlDotNet through its parser and representation model (<see cref="YamlStream"/>).</summary>
internal sealed class YamlDotNetConformanceEngine : IYamlConformanceEngine {
    public string Name => "YamlDotNet";

    public YamlEngineFeatures Features => YamlEngineFeatures.Events | YamlEngineFeatures.Load | YamlEngineFeatures.RoundTrip;

    public IReadOnlyList<YamlTestEvent> Parse(byte[] yaml) {
        var parser = new Parser(new StringReader(YamlTextEncoding.Decode(yaml, out _)));
        var events = new List<YamlTestEvent>();
        while (parser.MoveNext()) {
            YamlTestEvent? converted = Convert(parser.Current!);
            if (converted != null) events.Add(converted);
        }

        return events;
    }

    public IReadOnlyList<YamlConformanceNode> Load(byte[] yaml) {
        var stream = new YamlStream();
        stream.Load(new StringReader(YamlTextEncoding.Decode(yaml, out _)));
        var documents = new List<YamlConformanceNode>();
        foreach (YamlDocument document in stream.Documents) documents.Add(ToNode(document.RootNode, 0));
        return documents;
    }

    public byte[] RoundTrip(byte[] yaml) {
        string text = YamlTextEncoding.Decode(yaml, out YamlTextEncoding encoding);
        var stream = new YamlStream();
        stream.Load(new StringReader(text));
        var output = new StringWriter();
        stream.Save(output, assignAnchors: false);
        return encoding.Encode(output.ToString());
    }

    private static YamlTestEvent? Convert(ParsingEvent value) => value switch {
        StreamStart => YamlTestEvent.StreamStart(),
        StreamEnd => YamlTestEvent.StreamEnd(),
        DocumentStart start => YamlTestEvent.DocumentStart(!start.IsImplicit),
        DocumentEnd end => YamlTestEvent.DocumentEnd(!end.IsImplicit),
        MappingStart map => YamlTestEvent.MappingStart(map.Style == MappingStyle.Flow, Anchor(map.Anchor), Tag(map.Tag)),
        MappingEnd => YamlTestEvent.MappingEnd(),
        SequenceStart sequence => YamlTestEvent.SequenceStart(sequence.Style == SequenceStyle.Flow, Anchor(sequence.Anchor), Tag(sequence.Tag)),
        SequenceEnd => YamlTestEvent.SequenceEnd(),
        Scalar scalar => YamlTestEvent.Scalar(scalar.Value, Style(scalar.Style), Anchor(scalar.Anchor), Tag(scalar.Tag)),
        AnchorAlias alias => YamlTestEvent.Alias(alias.Value.Value),
        Comment => null,
        _ => throw new NotSupportedException("Unknown YamlDotNet event " + value.GetType().Name + ".")
    };

    private static YamlConformanceNode ToNode(YamlNode node, int depth) {
        if (depth > 1000) throw new InvalidOperationException("Nesting is too deep or recursive.");
        switch (node) {
            case YamlScalarNode scalar:
                string? tag = Tag(scalar.Tag);
                return YamlCoreSchema.Resolve(scalar.Value ?? string.Empty, scalar.Style == DotNetScalarStyle.Plain || scalar.Style == DotNetScalarStyle.Any, tag);
            case YamlSequenceNode sequence:
                var items = new List<YamlConformanceNode>();
                foreach (YamlNode child in sequence.Children) items.Add(ToNode(child, depth + 1));
                return YamlConformanceNode.Sequence(items);
            case YamlMappingNode mapping:
                var entries = new List<KeyValuePair<YamlConformanceNode, YamlConformanceNode>>();
                foreach (KeyValuePair<YamlNode, YamlNode> entry in mapping.Children) {
                    entries.Add(new KeyValuePair<YamlConformanceNode, YamlConformanceNode>(ToNode(entry.Key, depth + 1), ToNode(entry.Value, depth + 1)));
                }

                return YamlConformanceNode.Mapping(entries);
            default:
                throw new NotSupportedException("Unresolved YamlDotNet node " + node.GetType().Name + ".");
        }
    }

    private static string? Anchor(AnchorName anchor) => anchor.IsEmpty ? null : anchor.Value;

    private static string? Tag(TagName tag) => tag.IsEmpty ? null : tag.Value;

    private static YamlTestScalarStyle Style(DotNetScalarStyle style) => style switch {
        DotNetScalarStyle.SingleQuoted => YamlTestScalarStyle.SingleQuoted,
        DotNetScalarStyle.DoubleQuoted => YamlTestScalarStyle.DoubleQuoted,
        DotNetScalarStyle.Literal => YamlTestScalarStyle.Literal,
        DotNetScalarStyle.Folded => YamlTestScalarStyle.Folded,
        _ => YamlTestScalarStyle.Plain
    };
}
