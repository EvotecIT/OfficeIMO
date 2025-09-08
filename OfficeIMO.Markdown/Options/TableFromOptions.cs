using System.Collections.Generic;

namespace OfficeIMO.Markdown;

/// <summary>
/// Options for projecting objects/sequences into tables.
/// </summary>
public sealed class TableFromOptions {
    /// <summary>Optional set of property names to include (if specified, only these are used).</summary>
    public HashSet<string> Include { get; } = new HashSet<string>(System.StringComparer.OrdinalIgnoreCase);
    /// <summary>Optional set of property names to exclude.</summary>
    public HashSet<string> Exclude { get; } = new HashSet<string>(System.StringComparer.OrdinalIgnoreCase);
    /// <summary>Optional order of property names (others appear later in natural order).</summary>
    public List<string> Order { get; } = new List<string>();
    /// <summary>Optional mapping from property name to header text.</summary>
    public Dictionary<string, string> HeaderRenames { get; } = new Dictionary<string, string>(System.StringComparer.OrdinalIgnoreCase);
    /// <summary>Optional per-column alignment, applied by header name order when rendering.</summary>
    public List<ColumnAlignment> Alignments { get; } = new List<ColumnAlignment>();
}
