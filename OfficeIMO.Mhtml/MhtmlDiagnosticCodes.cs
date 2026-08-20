namespace OfficeIMO.Mhtml;

/// <summary>Stable diagnostics emitted by the MHTML archive boundary.</summary>
public static class MhtmlDiagnosticCodes {
    /// <summary>More than one related resource declared the same Content-ID; archive order wins.</summary>
    public const string DuplicateContentId = "MHTML_RESOURCE_CONTENT_ID_DUPLICATE";
    /// <summary>More than one related resource resolved to the same Content-Location; archive order wins.</summary>
    public const string DuplicateContentLocation = "MHTML_RESOURCE_CONTENT_LOCATION_DUPLICATE";
    /// <summary>A related resource declared a Content-Location that could not be resolved against the archive base URI.</summary>
    public const string InvalidContentLocation = "MHTML_RESOURCE_CONTENT_LOCATION_INVALID";
}
