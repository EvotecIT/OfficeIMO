namespace OfficeIMO.Pdf;

internal sealed class PdfGeneratedFormAccessibilityEvidence {
    public PdfGeneratedFormAccessibilityEvidence(string fieldName, int widgetCount, bool hasAccessibleName) {
        if (widgetCount <= 0) {
            throw new ArgumentOutOfRangeException(nameof(widgetCount), widgetCount, "PDF form evidence widget count must be positive.");
        }

        FieldName = fieldName ?? string.Empty;
        WidgetCount = widgetCount;
        HasAccessibleName = hasAccessibleName;
    }

    public string FieldName { get; }

    public int WidgetCount { get; }

    public bool HasAccessibleName { get; }
}
