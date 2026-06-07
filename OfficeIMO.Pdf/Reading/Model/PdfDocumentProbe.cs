namespace OfficeIMO.Pdf;

/// <summary>
/// Lightweight PDF feature markers that can be read before full parsing.
/// </summary>
public sealed class PdfDocumentProbe {
    internal PdfDocumentProbe(string? headerVersion, bool hasEncryption, bool hasSignatures, bool hasForms, bool hasAnnotations, bool hasOutlines, bool hasCatalogViewSettings, bool hasPageLabels, bool hasCatalogNameTrees, bool hasNamedDestinations, bool hasOpenActions, bool hasViewerPreferences, bool hasTaggedContent, bool hasXmpMetadata, bool hasCatalogUri, bool hasOutputIntents, bool hasEmbeddedFiles, bool hasOptionalContent, bool hasActiveContent, PdfDocumentSecurityInfo security) {
        HeaderVersion = headerVersion;
        HasEncryption = hasEncryption;
        HasSignatures = hasSignatures;
        HasForms = hasForms;
        HasAnnotations = hasAnnotations;
        HasOutlines = hasOutlines;
        HasCatalogViewSettings = hasCatalogViewSettings;
        HasPageLabels = hasPageLabels;
        HasCatalogNameTrees = hasCatalogNameTrees;
        HasNamedDestinations = hasNamedDestinations;
        HasOpenActions = hasOpenActions;
        HasViewerPreferences = hasViewerPreferences;
        HasTaggedContent = hasTaggedContent;
        HasXmpMetadata = hasXmpMetadata;
        HasCatalogUri = hasCatalogUri;
        HasOutputIntents = hasOutputIntents;
        HasEmbeddedFiles = hasEmbeddedFiles;
        HasOptionalContent = hasOptionalContent;
        HasActiveContent = hasActiveContent;
        Security = security;
    }

    /// <summary>PDF header version, for example 1.4, when a header is present.</summary>
    public string? HeaderVersion { get; }

    /// <summary>True when the document contains encryption markers.</summary>
    public bool HasEncryption { get; }

    /// <summary>True when the document contains digital signature markers.</summary>
    public bool HasSignatures { get; }

    /// <summary>True when the document contains AcroForm or form-field markers.</summary>
    public bool HasForms { get; }

    /// <summary>True when the document contains annotation markers.</summary>
    public bool HasAnnotations { get; }

    /// <summary>True when the document contains outline/bookmark markers.</summary>
    public bool HasOutlines { get; }

    /// <summary>True when the document contains catalog page mode or layout markers.</summary>
    public bool HasCatalogViewSettings { get; }

    /// <summary>True when the document contains page label markers.</summary>
    public bool HasPageLabels { get; }

    /// <summary>True when the document contains catalog name-tree markers.</summary>
    public bool HasCatalogNameTrees { get; }

    /// <summary>True when the document contains named destination markers.</summary>
    public bool HasNamedDestinations { get; }

    /// <summary>True when the document contains document open action markers.</summary>
    public bool HasOpenActions { get; }

    /// <summary>True when the document contains viewer preference markers.</summary>
    public bool HasViewerPreferences { get; }

    /// <summary>True when the document contains tagged PDF structure markers.</summary>
    public bool HasTaggedContent { get; }

    /// <summary>True when the document contains XMP metadata stream markers.</summary>
    public bool HasXmpMetadata { get; }

    /// <summary>True when the document catalog contains a URI dictionary.</summary>
    public bool HasCatalogUri { get; }

    /// <summary>True when the document contains output intent markers.</summary>
    public bool HasOutputIntents { get; }

    /// <summary>True when the document contains embedded file markers.</summary>
    public bool HasEmbeddedFiles { get; }

    /// <summary>True when the document contains optional content/layer markers.</summary>
    public bool HasOptionalContent { get; }

    /// <summary>True when the document contains active content markers such as JavaScript actions.</summary>
    public bool HasActiveContent { get; }

    /// <summary>Lightweight security, signature, and revision markers read from the PDF bytes.</summary>
    public PdfDocumentSecurityInfo Security { get; }
}
