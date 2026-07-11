namespace OfficeIMO.Html;

/// <summary>
/// Classifies the conversion impact represented by a diagnostic independently from its severity.
/// </summary>
public enum HtmlConversionLossKind {
    /// <summary>No content fidelity loss is represented.</summary>
    None,

    /// <summary>Content was retained through an approximate or fallback representation.</summary>
    Approximation,

    /// <summary>Some source content could not be represented and was omitted.</summary>
    Omission,

    /// <summary>The requested conversion could not produce a usable semantic result.</summary>
    Failure
}
