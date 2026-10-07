namespace OfficeIMO.Html;

/// <summary>Editable source after a bounded CSS-hidden-content filter for the selected media.</summary>
public sealed class HtmlVisibleContentResult : HtmlConversionResult<HtmlConversionDocument> {
    internal HtmlVisibleContentResult(HtmlConversionDocument document,
        IReadOnlyList<HtmlDiagnostic> diagnostics, int omittedElementCount, int appliedStylesheetCount)
        : base(document) {
        AddDiagnostics(diagnostics);
        OmittedElementCount = omittedElementCount;
        AppliedStylesheetCount = appliedStylesheetCount;
    }

    /// <summary>Elements omitted because their computed display or opacity makes them invisible.</summary>
    public int OmittedElementCount { get; }

    /// <summary>Linked stylesheets applied to decide visibility.</summary>
    public int AppliedStylesheetCount { get; }
}
