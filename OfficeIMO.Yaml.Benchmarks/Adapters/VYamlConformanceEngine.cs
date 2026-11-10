using System.Globalization;
using OfficeIMO.Yaml.Tests.Conformance;
using VYaml.Parser;
using VYaml.Serialization;

namespace OfficeIMO.Yaml.Benchmarks.Adapters;

/// <summary>
/// VYaml through its UTF-8 <see cref="YamlParser"/> and its dynamic deserialization. The parser's public surface has no
/// scalar style, flow marker or explicit document marker, so events are compared by content only, and loading uses
/// VYaml's own scalar resolution because the parser cannot tell a quoted scalar from a plain one. VYaml has no tree
/// model, so there is no round-trip. Depth is raised from VYaml's default of 64 to OfficeIMO.Yaml's 256 so real-world
/// files compare on content rather than on limit choices.
/// </summary>
internal sealed class VYamlConformanceEngine : IYamlConformanceEngine {
    internal const int MaxDepth = 256;

    /// <summary>The referenced package version; the VYaml assembly itself reports 1.0.0. Keep in step with the project file.</summary>
    internal const string PackageVersion = "1.6.1";
    private const string CoreTagPrefix = "tag:yaml.org,2002:";

    public string Name => "VYaml";

    public YamlEngineFeatures Features => YamlEngineFeatures.Events | YamlEngineFeatures.ContentEvents | YamlEngineFeatures.Load;

    public IReadOnlyList<YamlTestEvent> Parse(byte[] yaml) {
        var parser = YamlParser.FromBytes(yaml);
        parser.MaxDepth = MaxDepth;
        var events = new List<YamlTestEvent>();
        while (parser.Read()) events.Add(Convert(ref parser));
        return events;
    }

    public IReadOnlyList<YamlConformanceNode> Load(byte[] yaml) {
        var parser = YamlParser.FromBytes(yaml);
        parser.MaxDepth = MaxDepth;
        var documents = new List<YamlConformanceNode>();
        foreach (object? document in YamlSerializer.DeserializeMultipleDocuments<object?>(ref parser)) documents.Add(ToNode(document, 0));
        return documents;
    }

    public byte[] RoundTrip(byte[] yaml) => throw new NotSupportedException("VYaml has no document model to save.");

    /// <inheritdoc />
    public byte[] Emit(IReadOnlyList<YamlTestEvent> events) => throw new NotSupportedException();

    private static YamlTestEvent Convert(ref YamlParser parser) {
        switch (parser.CurrentEventType) {
            case ParseEventType.StreamStart: return YamlTestEvent.StreamStart();
            case ParseEventType.StreamEnd: return YamlTestEvent.StreamEnd();
            case ParseEventType.DocumentStart: return YamlTestEvent.DocumentStart(false);
            case ParseEventType.DocumentEnd: return YamlTestEvent.DocumentEnd(false);
            case ParseEventType.MappingStart: return YamlTestEvent.MappingStart(false, AnchorName(ref parser), TagName(ref parser));
            case ParseEventType.MappingEnd: return YamlTestEvent.MappingEnd();
            case ParseEventType.SequenceStart: return YamlTestEvent.SequenceStart(false, AnchorName(ref parser), TagName(ref parser));
            case ParseEventType.SequenceEnd: return YamlTestEvent.SequenceEnd();
            case ParseEventType.Scalar:
                return YamlTestEvent.Scalar(parser.GetScalarAsString() ?? string.Empty, YamlTestScalarStyle.Plain, AnchorName(ref parser), TagName(ref parser));
            case ParseEventType.Alias:
                return YamlTestEvent.Alias(AnchorName(ref parser) ?? throw new InvalidOperationException("VYaml reported an alias without a name."));
            default:
                throw new NotSupportedException("Unknown VYaml event " + parser.CurrentEventType + ".");
        }
    }

    private static string? AnchorName(ref YamlParser parser) => parser.TryGetCurrentAnchor(out Anchor anchor) ? anchor.Name : null;

    // VYaml keeps the handle and suffix apart and does not apply %TAG directives, so only the standard handles resolve.
    private static string? TagName(ref YamlParser parser) {
        if (!parser.TryGetCurrentTag(out Tag tag)) return null;
        return ResolveTag(tag.Handle, tag.Suffix);
    }

    internal static string ResolveTag(string handle, string suffix) => handle switch {
        "!!" => CoreTagPrefix + suffix,
        "" => suffix,
        _ => handle + suffix
    };

    private static YamlConformanceNode ToNode(object? value, int depth) {
        if (depth > 1000) throw new InvalidOperationException("Nesting is too deep or recursive.");
        switch (value) {
            case null: return YamlConformanceNode.Null;
            case bool flag: return YamlConformanceNode.Boolean(flag);
            case int number: return YamlConformanceNode.Integer(number.ToString(CultureInfo.InvariantCulture));
            case long number: return YamlConformanceNode.Integer(number.ToString(CultureInfo.InvariantCulture));
            case ulong number: return YamlConformanceNode.Integer(number.ToString(CultureInfo.InvariantCulture));
            case double number: return YamlConformanceNode.Float(FormatFloat(number));
            case float number: return YamlConformanceNode.Float(FormatFloat(number));
            case string text: return YamlConformanceNode.String(text);
            case IDictionary<object, object?> mapping: {
                var entries = new List<KeyValuePair<YamlConformanceNode, YamlConformanceNode>>(mapping.Count);
                foreach (KeyValuePair<object, object?> entry in mapping) entries.Add(new(ToNode(entry.Key, depth + 1), ToNode(entry.Value, depth + 1)));
                return YamlConformanceNode.Mapping(entries);
            }
            case IList<object?> sequence: {
                var items = new List<YamlConformanceNode>(sequence.Count);
                foreach (object? item in sequence) items.Add(ToNode(item, depth + 1));
                return YamlConformanceNode.Sequence(items);
            }
            default:
                throw new NotSupportedException("Unexpected VYaml value " + value.GetType().Name + ".");
        }
    }

    private static string FormatFloat(double value) =>
        double.IsPositiveInfinity(value) ? "Infinity" : double.IsNegativeInfinity(value) ? "-Infinity" : double.IsNaN(value) ? "NaN" : value.ToString("R", CultureInfo.InvariantCulture);
}
