using System.Globalization;

namespace OfficeIMO.Yaml.Tests.Conformance;

/// <summary>
/// OfficeIMO.Yaml over raw bytes, so its own encoding detection handles byte order marks, UTF-16 and UTF-32: events come
/// from <see cref="YamlReader"/>, loading and round-tripping from <see cref="YamlDocument"/>.
/// </summary>
public sealed class OfficeIMOYamlConformanceEngine : IYamlConformanceEngine {
    /// <inheritdoc />
    public string Name => "OfficeIMO.Yaml";

    /// <inheritdoc />
    public YamlEngineFeatures Features => YamlEngineFeatures.Events | YamlEngineFeatures.Load | YamlEngineFeatures.RoundTrip | YamlEngineFeatures.Lossless;

    /// <inheritdoc />
    public IReadOnlyList<YamlTestEvent> Parse(byte[] yaml) {
        var events = new List<YamlTestEvent>();
        using var reader = new YamlReader(new MemoryStream(yaml, writable: false));
        while (reader.Read()) {
            YamlTestEvent? converted = Convert(reader);
            if (converted != null) events.Add(converted);
        }

        return events;
    }

    /// <inheritdoc />
    public IReadOnlyList<YamlConformanceNode> Load(byte[] yaml) {
        YamlDocument document = YamlDocument.Load(new MemoryStream(yaml, writable: false));
        var documents = new List<YamlConformanceNode>(document.Documents.Count);
        foreach (YamlDocumentPart part in document.Documents) documents.Add(ToConformance(part.Root));
        return documents;
    }

    /// <summary>
    /// Loads and saves the stream without edits. Duplicate keys are admitted (last wins): the suite's valid cases are
    /// valid syntax, and a lossless editor must open files that repeat a key.
    /// </summary>
    public byte[] RoundTrip(byte[] yaml) {
        YamlDocument document = YamlDocument.Load(new MemoryStream(yaml, writable: false), LenientKeys);
        var output = new MemoryStream();
        document.Save(output);
        return output.ToArray();
    }

    private static readonly YamlLoadOptions LenientKeys = new() { DuplicateKeys = YamlDuplicateKeyPolicy.LastWins };

    /// <summary>Converts a node through the model's public typed view, resolving aliases.</summary>
    public static YamlConformanceNode ToConformance(YamlNode node) {
        switch (node) {
            case YamlAlias alias:
                return ToConformance(alias.Target);
            case YamlSequence sequence: {
                var items = new List<YamlConformanceNode>(sequence.Count);
                foreach (YamlNode item in sequence) items.Add(ToConformance(item));
                return YamlConformanceNode.Sequence(items);
            }
            case YamlMapping mapping: {
                var entries = new List<KeyValuePair<YamlConformanceNode, YamlConformanceNode>>(mapping.Count);
                foreach (YamlMappingEntry entry in mapping) entries.Add(new(ToConformance(entry.Key), ToConformance(entry.Value)));
                return YamlConformanceNode.Mapping(entries);
            }
            default:
                if (node.IsNull) return YamlConformanceNode.Null;
                if (node.TryGetValue(out bool flag)) return YamlConformanceNode.Boolean(flag);
                if (node.TryGetValue(out long integer)) return YamlConformanceNode.Integer(integer.ToString(CultureInfo.InvariantCulture));
                if (node.TryGetValue(out double number)) return YamlConformanceNode.Float(FormatFloat(number));
                return YamlConformanceNode.String(node.AsString());
        }
    }

    private static string FormatFloat(double value) =>
        double.IsPositiveInfinity(value) ? "Infinity" : double.IsNegativeInfinity(value) ? "-Infinity" : double.IsNaN(value) ? "NaN" : value.ToString("R", CultureInfo.InvariantCulture);

    /// <summary>Converts the reader's current event; comments have no suite equivalent and return null.</summary>
    public static YamlTestEvent? Convert(YamlReader reader) => reader.TokenType switch {
        YamlTokenType.StreamStart => YamlTestEvent.StreamStart(),
        YamlTokenType.StreamEnd => YamlTestEvent.StreamEnd(),
        YamlTokenType.DocumentStart => YamlTestEvent.DocumentStart(!reader.IsImplicit),
        YamlTokenType.DocumentEnd => YamlTestEvent.DocumentEnd(!reader.IsImplicit),
        YamlTokenType.MappingStart => YamlTestEvent.MappingStart(reader.Style == YamlStyle.Flow, reader.Anchor, reader.Tag),
        YamlTokenType.MappingEnd => YamlTestEvent.MappingEnd(),
        YamlTokenType.SequenceStart => YamlTestEvent.SequenceStart(reader.Style == YamlStyle.Flow, reader.Anchor, reader.Tag),
        YamlTokenType.SequenceEnd => YamlTestEvent.SequenceEnd(),
        YamlTokenType.Scalar => YamlTestEvent.Scalar(reader.GetString(), Style(reader.Style), reader.Anchor, reader.Tag),
        YamlTokenType.Alias => YamlTestEvent.Alias(reader.GetString()),
        _ => null
    };

    private static YamlTestScalarStyle Style(YamlStyle style) => style switch {
        YamlStyle.SingleQuoted => YamlTestScalarStyle.SingleQuoted,
        YamlStyle.DoubleQuoted => YamlTestScalarStyle.DoubleQuoted,
        YamlStyle.Literal => YamlTestScalarStyle.Literal,
        YamlStyle.Folded => YamlTestScalarStyle.Folded,
        _ => YamlTestScalarStyle.Plain
    };
}
