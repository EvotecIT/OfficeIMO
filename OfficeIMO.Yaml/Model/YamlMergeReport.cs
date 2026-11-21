namespace OfficeIMO.Yaml;

/// <summary>
/// What a merge changed in the base document, as paths in the grammar <see cref="YamlNode.Select"/> reads, plus the
/// findings it recorded: kind conflicts, edits of nodes that aliases share, renamed anchors, expanded aliases and keys a
/// null could not remove.
/// </summary>
public sealed class YamlMergeReport {
    internal YamlMergeReport(List<string> added, List<string> replaced, List<string> removed, IReadOnlyList<OfficeDiagnostic> diagnostics) {
        Added = added;
        Replaced = replaced;
        Removed = removed;
        Diagnostics = diagnostics;
    }

    /// <summary>Gets the paths of entries, items and documents the overlay added, in the order they were added.</summary>
    public IReadOnlyList<string> Added { get; }

    /// <summary>
    /// Gets the paths whose value the overlay changed: scalars with a different value, values of another kind and
    /// sequences replaced with different items. Values the overlay restates unchanged are not listed.
    /// </summary>
    public IReadOnlyList<string> Replaced { get; }

    /// <summary>Gets the paths of keys an overlay null removed under <see cref="YamlMergeOptions.NullRemovesKey"/>.</summary>
    public IReadOnlyList<string> Removed { get; }

    /// <summary>
    /// Gets the findings the merge recorded; they are also added to the base document's
    /// <see cref="YamlDocument.Diagnostics"/>.
    /// </summary>
    public IReadOnlyList<OfficeDiagnostic> Diagnostics { get; }
}
