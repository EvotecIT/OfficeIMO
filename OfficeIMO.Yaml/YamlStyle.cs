namespace OfficeIMO.Yaml;

/// <summary>
/// The presentation style of the node the reader is positioned on. Styles describe source syntax only; they never
/// change a node's value under the active schema, except that only <see cref="Plain"/> scalars resolve implicitly.
/// </summary>
public enum YamlStyle : byte {
    /// <summary>The current event has no style (stream, document, end and comment events).</summary>
    None = 0,
    /// <summary>An unquoted scalar, including an empty node. Only plain scalars resolve to null, boolean or number without a tag.</summary>
    Plain,
    /// <summary>A scalar in single quotes; <c>''</c> stands for one quote.</summary>
    SingleQuoted,
    /// <summary>A scalar in double quotes, which supports backslash escapes.</summary>
    DoubleQuoted,
    /// <summary>A literal block scalar (<c>|</c>), which keeps line breaks.</summary>
    Literal,
    /// <summary>A folded block scalar (<c>&gt;</c>), which folds line breaks into spaces.</summary>
    Folded,
    /// <summary>A collection laid out by indentation.</summary>
    Block,
    /// <summary>A collection in brackets (<c>[ ]</c>) or braces (<c>{ }</c>).</summary>
    Flow
}
