namespace OfficeIMO.Html;

/// <summary>Describes the observable renderer behavior for a declared capability.</summary>
public enum HtmlRenderSupportLevel {
    /// <summary>The declared feature subset is represented without a known renderer fallback.</summary>
    Full,
    /// <summary>A useful documented subset is represented and unsupported cases are diagnosed.</summary>
    Partial,
    /// <summary>The feature uses a deterministic documented fallback when encountered.</summary>
    Fallback,
    /// <summary>The feature is intentionally ignored and diagnosed.</summary>
    Ignored,
    /// <summary>The feature or resource is rejected at a safety boundary and diagnosed.</summary>
    Rejected
}
