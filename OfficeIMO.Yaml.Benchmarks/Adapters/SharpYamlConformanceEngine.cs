using OfficeIMO.Yaml.Tests.Conformance;
using SharpYaml;
using SharpYaml.Events;
using SharpYaml.Model;
using SharpYamlScalar = SharpYaml.Events.Scalar;
using SharpYamlScalarStyle = SharpYaml.ScalarStyle;
using SharpYamlStyle = SharpYaml.YamlStyle;

namespace OfficeIMO.Yaml.Benchmarks.Adapters;

// Inside the namespace so these win over the OfficeIMO.Yaml model types of the same name.
using YamlDocument = SharpYaml.Model.YamlDocument;
using YamlMapping = SharpYaml.Model.YamlMapping;
using YamlSequence = SharpYaml.Model.YamlSequence;

/// <summary>SharpYaml 3 through its parser and editable model (<see cref="SharpYaml.Model.YamlStream"/>).</summary>
internal sealed class SharpYamlConformanceEngine : IYamlConformanceEngine {
    public string Name => "SharpYaml";

    public YamlEngineFeatures Features => YamlEngineFeatures.Events | YamlEngineFeatures.Load | YamlEngineFeatures.RoundTrip;

    public IReadOnlyList<YamlTestEvent> Parse(byte[] yaml) {
        IParser parser = Parser.CreateParser(new StringReader(YamlTextEncoding.Decode(yaml, out _)));
        var events = new List<YamlTestEvent>();
        while (parser.MoveNext()) {
            YamlTestEvent? converted = Convert(parser.Current!);
            if (converted != null) events.Add(converted);
        }

        return events;
    }

    public IReadOnlyList<YamlConformanceNode> Load(byte[] yaml) {
        SharpYaml.Model.YamlStream stream = SharpYaml.Model.YamlStream.Load(new StringReader(YamlTextEncoding.Decode(yaml, out _)));
        var documents = new List<YamlConformanceNode>();
        // The model loads an alias as the anchored element itself, so aliases resolve without a lookup here.
        foreach (YamlDocument document in stream) documents.Add(document.Contents == null ? YamlConformanceNode.Null : ToNode(document.Contents, 0));
        return documents;
    }

    public byte[] RoundTrip(byte[] yaml) {
        string text = YamlTextEncoding.Decode(yaml, out YamlTextEncoding encoding);
        SharpYaml.Model.YamlStream stream = SharpYaml.Model.YamlStream.Load(new StringReader(text));
        var output = new StringWriter();
        stream.WriteTo(output, false);
        return encoding.Encode(output.ToString());
    }

    private static YamlTestEvent? Convert(ParsingEvent value) => value switch {
        StreamStart => YamlTestEvent.StreamStart(),
        StreamEnd => YamlTestEvent.StreamEnd(),
        DocumentStart start => YamlTestEvent.DocumentStart(!start.IsImplicit),
        DocumentEnd end => YamlTestEvent.DocumentEnd(!end.IsImplicit),
        MappingStart map => YamlTestEvent.MappingStart(map.Style == SharpYamlStyle.Flow, map.Anchor, map.Tag),
        MappingEnd => YamlTestEvent.MappingEnd(),
        SequenceStart sequence => YamlTestEvent.SequenceStart(sequence.Style == SharpYamlStyle.Flow, sequence.Anchor, sequence.Tag),
        SequenceEnd => YamlTestEvent.SequenceEnd(),
        SharpYamlScalar scalar => YamlTestEvent.Scalar(scalar.Value, Style(scalar.Style), scalar.Anchor, scalar.Tag),
        AnchorAlias alias => YamlTestEvent.Alias(alias.Value),
        _ => throw new NotSupportedException("Unknown SharpYaml event " + value.GetType().Name + ".")
    };

    private static YamlConformanceNode ToNode(YamlElement element, int depth) {
        if (depth > 1000) throw new InvalidOperationException("Nesting is too deep or recursive.");
        switch (element) {
            case YamlValue scalar:
                return YamlCoreSchema.Resolve(scalar.Value ?? string.Empty, scalar.Style == SharpYamlScalarStyle.Plain || scalar.Style == SharpYamlScalarStyle.Any, scalar.Tag);
            case YamlSequence sequence:
                var items = new List<YamlConformanceNode>();
                foreach (YamlElement child in sequence) items.Add(ToNode(child, depth + 1));
                return YamlConformanceNode.Sequence(items);
            case YamlMapping mapping:
                var entries = new List<KeyValuePair<YamlConformanceNode, YamlConformanceNode>>();
                foreach (KeyValuePair<YamlElement, YamlElement?> entry in mapping) {
                    YamlConformanceNode value = entry.Value == null ? YamlConformanceNode.Null : ToNode(entry.Value, depth + 1);
                    entries.Add(new KeyValuePair<YamlConformanceNode, YamlConformanceNode>(ToNode(entry.Key, depth + 1), value));
                }

                return YamlConformanceNode.Mapping(entries);
            default:
                throw new NotSupportedException("Unknown SharpYaml element " + element.GetType().Name + ".");
        }
    }

    private static YamlTestScalarStyle Style(SharpYamlScalarStyle style) => style switch {
        SharpYamlScalarStyle.SingleQuoted => YamlTestScalarStyle.SingleQuoted,
        SharpYamlScalarStyle.DoubleQuoted => YamlTestScalarStyle.DoubleQuoted,
        SharpYamlScalarStyle.Literal => YamlTestScalarStyle.Literal,
        SharpYamlScalarStyle.Folded => YamlTestScalarStyle.Folded,
        _ => YamlTestScalarStyle.Plain
    };
}
