namespace OfficeIMO.Yaml;

/// <summary>Scanner token kinds, following the token alphabet of the YAML 1.2 reference scanners.</summary>
internal enum YamlTokenKind : byte {
    None = 0,
    StreamStart,
    StreamEnd,
    VersionDirective,
    TagDirective,
    ReservedDirective,
    DocumentStart,
    DocumentEnd,
    BlockSequenceStart,
    BlockMappingStart,
    BlockEnd,
    FlowSequenceStart,
    FlowSequenceEnd,
    FlowMappingStart,
    FlowMappingEnd,
    BlockEntry,
    FlowEntry,
    Key,
    Value,
    Alias,
    Anchor,
    Tag,
    Scalar
}

/// <summary>Facts the scanner records so consumers can skip decoding work in the common case.</summary>
[Flags]
internal enum YamlTokenFlags : byte {
    None = 0,
    /// <summary>The scalar spans more than one line, so its value is folded.</summary>
    MultiLine = 1,
    /// <summary>The scalar contains escapes (<c>\</c> in double quotes, <c>''</c> in single quotes).</summary>
    Escapes = 2,
    /// <summary>The tag is verbatim (<c>!&lt;...&gt;</c>).</summary>
    Verbatim = 4,
    /// <summary>The tag suffix contains percent escapes.</summary>
    PercentEscapes = 8
}

/// <summary>Block scalar chomping indicator.</summary>
internal enum YamlChomping : byte {
    Clip = 0,
    Strip = 1,
    Keep = 2
}

/// <summary>
/// One scanner token. Offsets are absolute UTF-16 offsets into the decoded text; lines and columns are zero-based.
/// Zero-width tokens (implicit keys, block collection starts and ends) have <see cref="Start"/> equal to <see cref="End"/>.
/// Everything between the spans of consecutive tokens is trivia: whitespace, line breaks and comments.
/// </summary>
internal struct YamlToken {
    public YamlTokenKind Kind;
    public YamlStyle Style;
    public YamlTokenFlags Flags;
    public YamlChomping Chomping;
    public int Start;
    public int End;
    public int StartLine;
    public int StartColumn;
    public int EndLine;
    public int EndColumn;
    /// <summary>Content span: the scalar between quotes or after the block header, an anchor or alias name, a tag suffix, a directive argument.</summary>
    public int ValueStart;
    public int ValueEnd;
    /// <summary>Kind-specific data: block scalar content indentation, tag handle end offset, %YAML major*1000+minor, %TAG prefix start.</summary>
    public int Extra;
    /// <summary>Second kind-specific value: %TAG handle end.</summary>
    public int Extra2;
}

/// <summary>A position that may become an implicit key once a ':' follows on the same line.</summary>
internal struct YamlSimpleKey {
    public bool Possible;
    public bool Required;
    public bool TabBefore;
    public int TokenNumber;
    public int Offset;
    public int Line;
    public int Column;
}

/// <summary>A comment recorded when comment events are requested. Offsets are absolute.</summary>
internal struct YamlComment {
    public int Start;
    public int End;
    public int Line;
    public int Column;
}
