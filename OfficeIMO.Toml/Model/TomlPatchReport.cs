namespace OfficeIMO.Toml;

/// <summary>
/// What a JSON Patch or JSON Merge Patch changed in a document, as paths in the grammar <see cref="TomlNode.Select"/>
/// reads (<c>$</c> for the root table), plus the findings of the edits.
/// </summary>
public sealed class TomlPatchReport {
    internal TomlPatchReport(List<string> added, List<string> replaced, List<string> removed, List<OfficeDiagnostic> diagnostics) {
        Added = added;
        Replaced = replaced;
        Removed = removed;
        Diagnostics = diagnostics;
    }

    /// <summary>
    /// Gets the paths of keys, array items and tables the patch added, in the order they were added: <c>add</c>,
    /// <c>copy</c> and the target of <c>move</c> where nothing was, and new keys of a merge patch.
    /// </summary>
    public IReadOnlyList<string> Added { get; }

    /// <summary>
    /// Gets the paths whose value the patch changed: <c>replace</c>, and <c>add</c>, <c>copy</c> or <c>move</c> onto an
    /// existing key. A merge patch lists only values that differ from the patch.
    /// </summary>
    public IReadOnlyList<string> Replaced { get; }

    /// <summary>Gets the paths the patch removed: <c>remove</c>, the source of <c>move</c>, and merge patch nulls.</summary>
    public IReadOnlyList<string> Removed { get; }

    /// <summary>Gets the findings of the patch. None are recorded today; the list keeps the shape of <see cref="TomlMergeReport"/>.</summary>
    public IReadOnlyList<OfficeDiagnostic> Diagnostics { get; }
}
