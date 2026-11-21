namespace OfficeIMO.Yaml;

/// <summary>
/// Identifies the event a <see cref="YamlReader"/> is positioned on. Events follow the YAML 1.2 serialization
/// tree: a stream holds documents, a document holds one root node, and collections hold nodes.
/// </summary>
public enum YamlTokenType : byte {
    /// <summary>The reader has not been advanced yet, or it has passed the end of the stream.</summary>
    None = 0,
    /// <summary>The start of the stream. It is always the first event.</summary>
    StreamStart,
    /// <summary>The end of the stream. It is always the last event.</summary>
    StreamEnd,
    /// <summary>The start of a document. <see cref="YamlReader.IsImplicit"/> tells whether a <c>---</c> marker is present.</summary>
    DocumentStart,
    /// <summary>The end of a document. <see cref="YamlReader.IsImplicit"/> tells whether a <c>...</c> marker is present.</summary>
    DocumentEnd,
    /// <summary>The start of a sequence. <see cref="YamlReader.Style"/> is <see cref="YamlStyle.Block"/> or <see cref="YamlStyle.Flow"/>.</summary>
    SequenceStart,
    /// <summary>The end of the innermost open sequence.</summary>
    SequenceEnd,
    /// <summary>The start of a mapping. Its entries follow as alternating key and value nodes.</summary>
    MappingStart,
    /// <summary>The end of the innermost open mapping.</summary>
    MappingEnd,
    /// <summary>A scalar node, including the empty node an omitted value stands for.</summary>
    Scalar,
    /// <summary>An alias node (<c>*name</c>). <see cref="YamlReader.Anchor"/> holds the referenced anchor name.</summary>
    Alias,
    /// <summary>A comment. Comment events are produced only when <see cref="YamlReaderOptions.IncludeComments"/> is set.</summary>
    Comment
}
