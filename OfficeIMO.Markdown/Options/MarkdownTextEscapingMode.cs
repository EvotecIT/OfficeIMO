namespace OfficeIMO.Markdown;

/// <summary>Escaping policy for literal inline text and unstructured pipe-table cells.</summary>
public enum MarkdownTextEscapingMode {
    /// <summary>Preserves the existing writer behavior, including Markdown-capable table strings.</summary>
    Default,

    /// <summary>
    /// Writes literal text with Markdown backslash escapes, including HTML-like text and entity starts.
    /// Parentheses and ordinary ampersands remain literal. Unstructured table strings are treated as text;
    /// parsed and structured table content retains its explicit Markdown formatting.
    /// </summary>
    LiteralBackslash
}
