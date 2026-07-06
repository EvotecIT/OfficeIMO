namespace OfficeIMO.Pdf;

/// <summary>
/// Configures the generic preservation checks used after PDF rewrite and manipulation operations.
/// </summary>
public sealed class PdfRewritePreservationOptions {
    private readonly List<string> _requiredTextMarkers = new List<string>();
    private readonly HashSet<string> _allowedMetadataChanges = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

    /// <summary>True when the rewritten PDF must keep the original page count.</summary>
    public bool PreservePageCount { get; set; } = true;

    /// <summary>True when each unchanged page must keep width, height, and rotation.</summary>
    public bool PreservePageGeometry { get; set; } = true;

    /// <summary>True when Info dictionary metadata fields must be preserved except for explicitly allowed fields.</summary>
    public bool PreserveMetadata { get; set; } = true;

    /// <summary>True when outline/bookmark count must not be lost.</summary>
    public bool PreserveOutlines { get; set; } = true;

    /// <summary>True when named destination count must not be lost.</summary>
    public bool PreserveNamedDestinations { get; set; } = true;

    /// <summary>True when page-label count and label ranges must not be lost.</summary>
    public bool PreservePageLabels { get; set; } = true;

    /// <summary>True when simple link annotation count must not be lost.</summary>
    public bool PreserveLinkAnnotations { get; set; } = true;

    /// <summary>True when generic annotation count must not be lost.</summary>
    public bool PreserveAnnotations { get; set; } = true;

    /// <summary>True when simple AcroForm field count and form markers must not be lost.</summary>
    public bool PreserveForms { get; set; } = true;

    /// <summary>True when embedded file/associated file count must not be lost.</summary>
    public bool PreserveEmbeddedFiles { get; set; } = true;

    /// <summary>True when XMP metadata markers must not be lost.</summary>
    public bool PreserveXmpMetadata { get; set; } = true;

    /// <summary>True when output intent count must not be lost.</summary>
    public bool PreserveOutputIntents { get; set; } = true;

    /// <summary>True when optional content/layer markers must not be lost.</summary>
    public bool PreserveOptionalContent { get; set; } = true;

    /// <summary>True when tagged PDF structure markers and readable structure metadata must not be lost.</summary>
    public bool PreserveTaggedContent { get; set; } = true;

    /// <summary>True when catalog page mode, layout, language, and viewer preferences must not be lost.</summary>
    public bool PreserveCatalogViewSettings { get; set; } = true;

    /// <summary>True when document open-action destination metadata must not be changed.</summary>
    public bool PreserveOpenAction { get; set; } = true;

    /// <summary>True when catalog-level active action metadata must not be changed.</summary>
    public bool PreserveCatalogActions { get; set; } = true;

    /// <summary>True when page-level additional action metadata must not be changed.</summary>
    public bool PreservePageActions { get; set; } = true;

    /// <summary>True when viewer preference values must not be changed.</summary>
    public bool PreserveViewerPreferences { get; set; } = true;

    /// <summary>True when header, catalog, and effective PDF version markers must not be changed.</summary>
    public bool PreserveDocumentVersionState { get; set; } = true;

    /// <summary>True when xref-stream, object-stream, previous-revision, and incremental-update markers must not be lost.</summary>
    public bool PreserveRevisionStructure { get; set; } = true;

    /// <summary>True when security, signature, DSS/VRI, DocMDP, and usage-rights markers must not be lost or changed.</summary>
    public bool PreserveSecurityState { get; set; } = true;

    /// <summary>Text markers that must remain extractable from the rewritten PDF.</summary>
    public IList<string> RequiredTextMarkers => _requiredTextMarkers;

    /// <summary>Metadata fields that the rewrite is allowed to change. Supported names: Title, Author, Subject, Keywords.</summary>
    public ISet<string> AllowedMetadataChanges => _allowedMetadataChanges;

    /// <summary>Adds required text markers and returns this options object for fluent setup.</summary>
    public PdfRewritePreservationOptions RequireTextMarkers(params string[] markers) {
        Guard.NotNull(markers, nameof(markers));
        for (int i = 0; i < markers.Length; i++) {
            if (!string.IsNullOrEmpty(markers[i])) {
                _requiredTextMarkers.Add(markers[i]);
            }
        }

        return this;
    }

    /// <summary>Allows specific metadata fields to change and returns this options object for fluent setup.</summary>
    public PdfRewritePreservationOptions AllowMetadataChanges(params string[] fieldNames) {
        Guard.NotNull(fieldNames, nameof(fieldNames));
        for (int i = 0; i < fieldNames.Length; i++) {
            if (!string.IsNullOrWhiteSpace(fieldNames[i])) {
                _allowedMetadataChanges.Add(fieldNames[i]);
            }
        }

        return this;
    }
}
