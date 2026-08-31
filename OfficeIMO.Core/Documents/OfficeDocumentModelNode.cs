using System;
using System.Collections.Generic;

namespace OfficeIMO;

/// <summary>
/// Source-neutral recursive document structure used by hierarchical text formats.
/// Format-specific syntax and lossless extension data remain owned by the native format package.
/// </summary>
public sealed class OfficeDocumentModelNode {
    /// <summary>Stable node identifier when one is available.</summary>
    public string Id { get; set; } = string.Empty;

    /// <summary>Producer-normalized node kind, such as outline, section, paragraph, list, or table.</summary>
    public string Kind { get; set; } = string.Empty;

    /// <summary>Primary plain-text value.</summary>
    public string Text { get; set; } = string.Empty;

    /// <summary>Optional nesting or heading level.</summary>
    public int? Level { get; set; }

    /// <summary>Portable semantic attributes, using expanded names when a source attribute is namespaced.</summary>
    public IReadOnlyDictionary<string, string> Attributes { get; set; } =
        new Dictionary<string, string>(StringComparer.Ordinal);

    /// <summary>Child nodes in source order.</summary>
    public IReadOnlyList<OfficeDocumentModelNode> Children { get; set; } = Array.Empty<OfficeDocumentModelNode>();

    /// <summary>Source location.</summary>
    public OfficeDocumentModelLocation Location { get; set; } = new OfficeDocumentModelLocation();
}
