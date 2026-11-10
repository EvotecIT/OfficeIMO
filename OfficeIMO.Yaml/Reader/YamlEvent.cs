namespace OfficeIMO.Yaml;

/// <summary>
/// One parser event. Offsets are absolute; lines and columns are zero-based. Node events start at their first
/// property (anchor or tag) when present, so the span from <see cref="Start"/> to <see cref="End"/> covers the node's
/// full source text for scalars, aliases and flow collections, and the properties for block collection starts.
/// </summary>
internal struct YamlEvent {
    public YamlTokenType Type;
    public YamlStyle Style;
    public YamlTokenFlags Flags;
    public YamlChomping Chomping;
    public bool Implicit;
    public bool IsKey;
    public int Start;
    public int StartLine;
    public int StartColumn;
    public int End;
    public int EndLine;
    public int EndColumn;
    public int ValueStart;
    public int ValueEnd;
    public int Indent;
    public int AnchorStart;
    public int AnchorEnd;
    public bool HasTag;
    public YamlTokenFlags TagFlags;
    public string? TagPrefix;
    public int TagSuffixStart;
    public int TagSuffixEnd;
}

/// <summary>Parser states, following the productions of the YAML 1.2.2 serialization grammar.</summary>
internal enum YamlParserState : byte {
    StreamStart,
    ImplicitDocumentStart,
    DocumentStart,
    DocumentContent,
    DocumentEnd,
    BlockNode,
    BlockSequenceEntry,
    IndentlessSequenceEntry,
    BlockMappingKey,
    BlockMappingValue,
    FlowSequenceFirstEntry,
    FlowSequenceEntry,
    FlowSequenceEntryMappingKey,
    FlowSequenceEntryMappingEmptyKey,
    FlowSequenceEntryMappingValue,
    FlowSequenceEntryMappingEnd,
    FlowMappingFirstKey,
    FlowMappingKey,
    FlowMappingValue,
    End
}
