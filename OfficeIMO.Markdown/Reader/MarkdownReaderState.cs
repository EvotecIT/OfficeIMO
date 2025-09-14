using System.Collections.Generic;

namespace OfficeIMO.Markdown;

/// <summary>
/// Mutable per-parse state shared across block and inline parsers.
/// </summary>
public sealed class MarkdownReaderState {
    public Dictionary<string, (string Url, string? Title)> LinkRefs { get; } = new Dictionary<string, (string, string?)>(System.StringComparer.OrdinalIgnoreCase);
}
