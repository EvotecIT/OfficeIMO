using System.Reflection;
using OfficeIMO.Yaml.Tests.Conformance;
using One = OneYaml;

namespace OfficeIMO.Yaml.Benchmarks.Adapters;

/// <summary>
/// OneYaml through its lossless <see cref="One.YamlDocument"/>. OneYaml has no public event parser, so events are read
/// from the document's nodes: kinds, scalar text, anchors, tags and aliases, compared by content because the nodes do
/// not say whether a document marker was explicit. Loading resolves each node's text with the shared YAML 1.2
/// core-schema resolver, as the other adapters do. Duplicate keys are allowed, as OfficeIMO.Yaml's default mode allows
/// them, so files compare on structure rather than on that policy. OneYaml reads and writes bytes itself, detecting
/// UTF-8, UTF-16 and UTF-32.
/// </summary>
internal sealed class OneYamlConformanceEngine : IYamlConformanceEngine {
    private static readonly One.YamlDocumentOptions Options = new() { AllowDuplicateKeys = true };

    /// <summary>The referenced package version. Keep in step with the project file.</summary>
    internal static string PackageVersion { get; } =
        typeof(One.YamlDocument).Assembly.GetCustomAttribute<AssemblyInformationalVersionAttribute>()?.InformationalVersion.Split('+')[0] ?? "unknown";

    public string Name => "OneYaml";

    public YamlEngineFeatures Features => YamlEngineFeatures.Events | YamlEngineFeatures.ContentEvents | YamlEngineFeatures.Load | YamlEngineFeatures.RoundTrip | YamlEngineFeatures.Lossless;

    public IReadOnlyList<YamlTestEvent> Parse(byte[] yaml) {
        using One.YamlDocument document = One.YamlDocument.Parse(yaml, Options);
        var events = new List<YamlTestEvent> { YamlTestEvent.StreamStart() };
        foreach (One.YamlNode root in document.Documents) {
            events.Add(YamlTestEvent.DocumentStart(false));
            AddEvents(root, events, 0);
            events.Add(YamlTestEvent.DocumentEnd(false));
        }

        events.Add(YamlTestEvent.StreamEnd());
        return events;
    }

    public IReadOnlyList<YamlConformanceNode> Load(byte[] yaml) {
        using One.YamlDocument document = One.YamlDocument.Parse(yaml, Options);
        var documents = new List<YamlConformanceNode>();
        foreach (One.YamlNode root in document.Documents) documents.Add(ToNode(root, 0));
        return documents;
    }

    public byte[] RoundTrip(byte[] yaml) {
        using One.YamlDocument document = One.YamlDocument.Parse(yaml, Options);
        return document.ToBytes();
    }

    /// <inheritdoc />
    public byte[] Emit(IReadOnlyList<YamlTestEvent> events) => throw new NotSupportedException();

    private static void AddEvents(One.YamlNode node, List<YamlTestEvent> events, int depth) {
        if (depth > 1000) throw new InvalidOperationException("Nesting is too deep or recursive.");
        switch (node.Kind) {
            case One.YamlNodeKind.Alias:
                events.Add(YamlTestEvent.Alias(node.ToString().TrimStart('*').Trim()));
                return;
            case One.YamlNodeKind.Scalar:
                events.Add(YamlTestEvent.Scalar(ScalarText(node), Style(node.Style), node.Anchor, node.Tag));
                return;
            case One.YamlNodeKind.Mapping:
                events.Add(YamlTestEvent.MappingStart(node.Style == One.YamlNodeStyle.Flow, node.Anchor, node.Tag));
                foreach (KeyValuePair<One.YamlNode, One.YamlNode> entry in node.Entries) {
                    AddEvents(entry.Key, events, depth + 1);
                    AddEvents(entry.Value, events, depth + 1);
                }

                events.Add(YamlTestEvent.MappingEnd());
                return;
            default:
                events.Add(YamlTestEvent.SequenceStart(node.Style == One.YamlNodeStyle.Flow, node.Anchor, node.Tag));
                foreach (One.YamlNode item in node.Items) AddEvents(item, events, depth + 1);
                events.Add(YamlTestEvent.SequenceEnd());
                return;
        }
    }

    private static YamlConformanceNode ToNode(One.YamlNode node, int depth) {
        if (depth > 1000) throw new InvalidOperationException("Nesting is too deep or recursive.");
        if (node.Kind == One.YamlNodeKind.Alias) node = node.Target;
        switch (node.Kind) {
            case One.YamlNodeKind.Scalar:
                return YamlCoreSchema.Resolve(ScalarText(node), node.Style == One.YamlNodeStyle.Plain, node.Tag);
            case One.YamlNodeKind.Mapping: {
                var entries = new List<KeyValuePair<YamlConformanceNode, YamlConformanceNode>>();
                foreach (KeyValuePair<One.YamlNode, One.YamlNode> entry in node.Entries) entries.Add(new(ToNode(entry.Key, depth + 1), ToNode(entry.Value, depth + 1)));
                return YamlConformanceNode.Mapping(entries);
            }
            default: {
                var items = new List<YamlConformanceNode>();
                foreach (One.YamlNode item in node.Items) items.Add(ToNode(item, depth + 1));
                return YamlConformanceNode.Sequence(items);
            }
        }
    }

    // GetValue<string> gives the scalar's content, except for a plain scalar that resolves to null (empty, ~, null) or
    // carries a tag its text does not fit; those are single-line plain text, so the source without properties is the text.
    private static string ScalarText(One.YamlNode node) {
        string? text = null;
        try {
            text = node.GetValue<string>();
        } catch (One.YamlException) when (node.Style == One.YamlNodeStyle.Plain) {
        }

        if (text != null) return text;
        ReadOnlySpan<char> rest = node.ToString().AsSpan().Trim();
        while (rest.Length > 0 && (rest[0] == '&' || rest[0] == '!')) {
            int space = rest.IndexOfAny(" \t\r\n");
            if (space < 0) return string.Empty;
            rest = rest.Slice(space).TrimStart();
        }

        return rest.ToString();
    }

    private static YamlTestScalarStyle Style(One.YamlNodeStyle style) => style switch {
        One.YamlNodeStyle.SingleQuoted => YamlTestScalarStyle.SingleQuoted,
        One.YamlNodeStyle.DoubleQuoted => YamlTestScalarStyle.DoubleQuoted,
        One.YamlNodeStyle.Literal => YamlTestScalarStyle.Literal,
        One.YamlNodeStyle.Folded => YamlTestScalarStyle.Folded,
        _ => YamlTestScalarStyle.Plain
    };
}
