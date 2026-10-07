namespace OfficeIMO.Epub;

/// <summary>A publisher-labelled link from an index entry to existing publication content.</summary>
public sealed class EpubIndexLocator {
    /// <summary>Manifest identifier of an XHTML spine document.</summary>
    public string ManifestId { get; set; } = string.Empty;
    /// <summary>Optional existing body-content identifier. Null links to the document.</summary>
    public string? FragmentId { get; set; }
    /// <summary>Visible locator, such as a source page label or section title. No pagination is inferred.</summary>
    public string Label { get; set; } = string.Empty;
}
