namespace OfficeIMO.Yaml.Tests.Conformance;

/// <summary>
/// OfficeIMO.Yaml through <see cref="YamlReader"/> over a byte stream, so the reader's own encoding detection handles
/// byte order marks, UTF-16 and UTF-32. Loading and round-tripping arrive with the document model.
/// </summary>
public sealed class OfficeIMOYamlConformanceEngine : IYamlConformanceEngine {
    /// <inheritdoc />
    public string Name => "OfficeIMO.Yaml";

    /// <inheritdoc />
    public YamlEngineFeatures Features => YamlEngineFeatures.Events;

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
    public IReadOnlyList<YamlConformanceNode> Load(byte[] yaml) => throw new NotSupportedException("OfficeIMO.Yaml has no document model yet.");

    /// <inheritdoc />
    public byte[] RoundTrip(byte[] yaml) => throw new NotSupportedException("OfficeIMO.Yaml has no document model yet.");

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
