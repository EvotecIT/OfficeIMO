namespace OfficeIMO.Yaml.Tests.Conformance;

/// <summary>Event kinds of the yaml-test-suite <c>test.event</c> format.</summary>
public enum YamlTestEventKind {
    /// <summary><c>+STR</c></summary>
    StreamStart,
    /// <summary><c>-STR</c></summary>
    StreamEnd,
    /// <summary><c>+DOC</c></summary>
    DocumentStart,
    /// <summary><c>-DOC</c></summary>
    DocumentEnd,
    /// <summary><c>+MAP</c></summary>
    MappingStart,
    /// <summary><c>-MAP</c></summary>
    MappingEnd,
    /// <summary><c>+SEQ</c></summary>
    SequenceStart,
    /// <summary><c>-SEQ</c></summary>
    SequenceEnd,
    /// <summary><c>=VAL</c></summary>
    Scalar,
    /// <summary><c>=ALI</c></summary>
    Alias
}

/// <summary>Presentation style of a scalar.</summary>
public enum YamlTestScalarStyle {
    /// <summary>Plain (<c>:</c> in event text).</summary>
    Plain,
    /// <summary>Single-quoted (<c>'</c>).</summary>
    SingleQuoted,
    /// <summary>Double-quoted (<c>"</c>).</summary>
    DoubleQuoted,
    /// <summary>Literal block (<c>|</c>).</summary>
    Literal,
    /// <summary>Folded block (<c>&gt;</c>).</summary>
    Folded
}

/// <summary>One parser event in the shape the yaml-test-suite describes.</summary>
public sealed class YamlTestEvent {
    private YamlTestEvent(YamlTestEventKind kind, bool isExplicit, bool isFlow, string? anchor, string? tag, YamlTestScalarStyle style, string? value) {
        Kind = kind;
        IsExplicit = isExplicit;
        IsFlow = isFlow;
        Anchor = anchor;
        Tag = tag;
        Style = style;
        Value = value;
    }

    /// <summary>Event kind.</summary>
    public YamlTestEventKind Kind { get; }

    /// <summary>For documents: the start had a <c>---</c> marker or the end had a <c>...</c> marker.</summary>
    public bool IsExplicit { get; }

    /// <summary>For mapping and sequence starts: the collection uses flow style.</summary>
    public bool IsFlow { get; }

    /// <summary>Anchor name without <c>&amp;</c>, or null.</summary>
    public string? Anchor { get; }

    /// <summary>
    /// Resolved tag, for example <c>tag:yaml.org,2002:str</c>, a local tag such as <c>!foo</c>,
    /// <c>!</c> for the non-specific tag, or null when the node has no tag.
    /// </summary>
    public string? Tag { get; }

    /// <summary>Scalar style; only meaningful for scalars.</summary>
    public YamlTestScalarStyle Style { get; }

    /// <summary>Scalar content, or the alias name without <c>*</c>.</summary>
    public string? Value { get; }

    /// <summary>Creates <c>+STR</c>.</summary>
    public static YamlTestEvent StreamStart() => new(YamlTestEventKind.StreamStart, false, false, null, null, default, null);

    /// <summary>Creates <c>-STR</c>.</summary>
    public static YamlTestEvent StreamEnd() => new(YamlTestEventKind.StreamEnd, false, false, null, null, default, null);

    /// <summary>Creates <c>+DOC</c>.</summary>
    public static YamlTestEvent DocumentStart(bool isExplicit) => new(YamlTestEventKind.DocumentStart, isExplicit, false, null, null, default, null);

    /// <summary>Creates <c>-DOC</c>.</summary>
    public static YamlTestEvent DocumentEnd(bool isExplicit) => new(YamlTestEventKind.DocumentEnd, isExplicit, false, null, null, default, null);

    /// <summary>Creates <c>+MAP</c>.</summary>
    public static YamlTestEvent MappingStart(bool isFlow, string? anchor = null, string? tag = null) =>
        new(YamlTestEventKind.MappingStart, false, isFlow, Empty(anchor), Empty(tag), default, null);

    /// <summary>Creates <c>-MAP</c>.</summary>
    public static YamlTestEvent MappingEnd() => new(YamlTestEventKind.MappingEnd, false, false, null, null, default, null);

    /// <summary>Creates <c>+SEQ</c>.</summary>
    public static YamlTestEvent SequenceStart(bool isFlow, string? anchor = null, string? tag = null) =>
        new(YamlTestEventKind.SequenceStart, false, isFlow, Empty(anchor), Empty(tag), default, null);

    /// <summary>Creates <c>-SEQ</c>.</summary>
    public static YamlTestEvent SequenceEnd() => new(YamlTestEventKind.SequenceEnd, false, false, null, null, default, null);

    /// <summary>Creates <c>=VAL</c>.</summary>
    public static YamlTestEvent Scalar(string value, YamlTestScalarStyle style, string? anchor = null, string? tag = null) =>
        new(YamlTestEventKind.Scalar, false, false, Empty(anchor), Empty(tag), style, value ?? string.Empty);

    /// <summary>Creates <c>=ALI</c>.</summary>
    public static YamlTestEvent Alias(string name) => new(YamlTestEventKind.Alias, false, false, null, null, default, name);

    /// <summary>Returns the suite's text form of this event.</summary>
    public override string ToString() => YamlTestEventText.Format(this);

    private static string? Empty(string? value) => string.IsNullOrEmpty(value) ? null : value;
}
