namespace OfficeIMO.Yaml;

/// <summary>
/// What a JSON Patch or JSON Merge Patch changed in a document, as paths in the grammar <see cref="YamlNode.Select"/>
/// reads, plus the findings the edits recorded, such as edits of nodes that aliases share.
/// </summary>
public sealed class YamlPatchReport {
    internal YamlPatchReport(List<string> added, List<string> replaced, List<string> removed, IReadOnlyList<OfficeDiagnostic> diagnostics) {
        Added = added;
        Replaced = replaced;
        Removed = removed;
        Diagnostics = diagnostics;
    }

    /// <summary>
    /// Gets the paths of entries, items and documents the patch added, in the order they were added: <c>add</c>,
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

    /// <summary>
    /// Gets the findings the patch recorded; they are also added to the document's <see cref="YamlDocument.Diagnostics"/>.
    /// </summary>
    public IReadOnlyList<OfficeDiagnostic> Diagnostics { get; }
}
