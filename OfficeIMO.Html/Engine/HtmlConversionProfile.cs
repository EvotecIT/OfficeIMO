namespace OfficeIMO.Html;

/// <summary>
/// Shared OfficeIMO HTML conversion profile understood by gallery contracts, diagnostics, scoring, and adapters.
/// </summary>
public enum HtmlConversionProfile {
    /// <summary>
    /// Prioritizes clean semantic document output over browser-perfect visual reproduction.
    /// </summary>
    Semantic,

    /// <summary>
    /// Balances semantic output, document styling, tables, forms, images, and diagnostics for common office documents.
    /// </summary>
    Document,

    /// <summary>
    /// Captures the high-fidelity print/PDF ambition where layout preservation is more important than editability.
    /// </summary>
    HighFidelityPrint
}
