namespace OfficeIMO.Pdf;

/// <summary>
/// Controls whether a source-to-PDF conversion may read resources from the host environment.
/// Explicit in-memory font and image data remains available in every policy.
/// </summary>
public sealed class PdfResourcePolicy {
    /// <summary>
    /// Creates the balanced document-conversion default. Installed fonts and bounded in-source resources are allowed
    /// for Unicode fidelity, while arbitrary local-file and remote-resource reads remain disabled.
    /// </summary>
    public static PdfResourcePolicy CreateDefault() => new PdfResourcePolicy {
        AllowSystemFontEmbedding = true,
        AllowDataUris = true,
        AllowEmbeddedPackageResources = true
    };

    /// <summary>
    /// Creates the portable deterministic policy. Host fonts, local files, and remote resources are disabled;
    /// bounded in-document data resources remain available.
    /// </summary>
    public static PdfResourcePolicy CreatePortableDeterministic() => new PdfResourcePolicy {
        AllowDataUris = true,
        AllowEmbeddedPackageResources = true
    };

    /// <summary>
    /// Creates an explicit trusted-host policy. Resolvers and format-specific options are still required
    /// before local or remote content is loaded.
    /// </summary>
    public static PdfResourcePolicy CreateTrustedHost() => new PdfResourcePolicy {
        AllowSystemFontEmbedding = true,
        AllowDocumentFontEmbedding = true,
        AllowLocalFileAccess = true,
        AllowRemoteResourceResolution = true,
        AllowDataUris = true,
        AllowEmbeddedPackageResources = true
    };

    /// <summary>When true, converters may load and embed installed system fonts. Enabled by the balanced default.</summary>
    public bool AllowSystemFontEmbedding { get; set; }

    /// <summary>
    /// When true, Office converters may use font family names from the source document to locate and embed installed
    /// system fonts. Disabled by default because document-controlled names must not trigger host font disclosure.
    /// </summary>
    public bool AllowDocumentFontEmbedding { get; set; }

    /// <summary>When true, format-specific local-resource options may read files from explicitly allowed locations.</summary>
    public bool AllowLocalFileAccess { get; set; }

    /// <summary>When true, application-supplied resolvers may obtain remote resources.</summary>
    public bool AllowRemoteResourceResolution { get; set; }

    /// <summary>When true, bounded data-URI resources embedded in the source document may be decoded.</summary>
    public bool AllowDataUris { get; set; } = true;

    /// <summary>When true, bounded resources already contained in a source package, such as MHTML <c>cid:</c> parts, may be resolved.</summary>
    public bool AllowEmbeddedPackageResources { get; set; } = true;

    /// <summary>Creates an independent policy snapshot.</summary>
    public PdfResourcePolicy Clone() => new PdfResourcePolicy {
        AllowSystemFontEmbedding = AllowSystemFontEmbedding,
        AllowDocumentFontEmbedding = AllowDocumentFontEmbedding,
        AllowLocalFileAccess = AllowLocalFileAccess,
        AllowRemoteResourceResolution = AllowRemoteResourceResolution,
        AllowDataUris = AllowDataUris,
        AllowEmbeddedPackageResources = AllowEmbeddedPackageResources
    };
}
