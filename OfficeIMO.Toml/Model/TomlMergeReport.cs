namespace OfficeIMO.Toml;

/// <summary>
/// What a merge changed in the base document, as paths in the grammar <see cref="TomlNode.Select"/> reads, plus the
/// kind conflicts it resolved.
/// </summary>
public sealed class TomlMergeReport {
    internal TomlMergeReport(List<string> added, List<string> replaced, List<OfficeDiagnostic> diagnostics) {
        Added = added;
        Replaced = replaced;
        Diagnostics = diagnostics;
    }

    /// <summary>Gets the paths of keys and array items the overlay added, in the order they were added.</summary>
    public IReadOnlyList<string> Added { get; }

    /// <summary>
    /// Gets the paths whose value the overlay changed: values with a different value, values of another kind and arrays
    /// replaced with different items. Values the overlay restates unchanged are not listed.
    /// </summary>
    public IReadOnlyList<string> Replaced { get; }

    /// <summary>
    /// Gets the findings of the merge: a <c>toml-merge-kind-conflict</c> warning for each base value the overlay replaced
    /// with one of another kind, such as a table over a string.
    /// </summary>
    public IReadOnlyList<OfficeDiagnostic> Diagnostics { get; }
}
