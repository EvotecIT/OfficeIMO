namespace OfficeIMO.Pdf;

/// <summary>
/// Preserves the original source-compatible form-operation calls without making
/// an untyped <see langword="null"/> ambiguous with <see cref="PdfLoadOptions"/>.
/// </summary>
public static class PdfDocumentFormsCompatibilityExtensions {
    /// <summary>Attempts to fill string form values with explicit form options.</summary>
    public static PdfOperationResult<PdfDocument> TryFill(
        this PdfDocumentForms forms,
        IReadOnlyDictionary<string, string> fieldValues,
        PdfFormFillerOptions formOptions) {
        Guard.NotNull(forms, nameof(forms));
        return forms.TryFill(fieldValues, formOptions, readOptions: null);
    }

    /// <summary>Attempts to fill typed form values with explicit form options.</summary>
    public static PdfOperationResult<PdfDocument> TryFill(
        this PdfDocumentForms forms,
        IReadOnlyDictionary<string, PdfFormFieldValue> fieldValues,
        PdfFormFillerOptions formOptions) {
        Guard.NotNull(forms, nameof(forms));
        return forms.TryFill(fieldValues, formOptions, readOptions: null);
    }

    /// <summary>Attempts to flatten form fields with explicit form options.</summary>
    public static PdfOperationResult<PdfDocument> TryFlatten(
        this PdfDocumentForms forms,
        PdfFormFillerOptions formOptions) {
        Guard.NotNull(forms, nameof(forms));
        return forms.TryFlatten(formOptions, readOptions: null);
    }

    /// <summary>Attempts to fill and flatten string form values with explicit form options.</summary>
    public static PdfOperationResult<PdfDocument> TryFillAndFlatten(
        this PdfDocumentForms forms,
        IReadOnlyDictionary<string, string> fieldValues,
        PdfFormFillerOptions formOptions) {
        Guard.NotNull(forms, nameof(forms));
        return forms.TryFillAndFlatten(fieldValues, formOptions, readOptions: null);
    }

    /// <summary>Attempts to fill and flatten typed form values with explicit form options.</summary>
    public static PdfOperationResult<PdfDocument> TryFillAndFlatten(
        this PdfDocumentForms forms,
        IReadOnlyDictionary<string, PdfFormFieldValue> fieldValues,
        PdfFormFillerOptions formOptions) {
        Guard.NotNull(forms, nameof(forms));
        return forms.TryFillAndFlatten(fieldValues, formOptions, readOptions: null);
    }
}
