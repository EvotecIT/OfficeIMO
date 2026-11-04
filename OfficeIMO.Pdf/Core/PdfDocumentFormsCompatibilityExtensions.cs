namespace OfficeIMO.Pdf;

/// <summary>
/// Preserves the original source-compatible form-operation calls without making
/// an untyped <see langword="null"/> ambiguous with <see cref="PdfLoadOptions"/>.
/// </summary>
internal static class PdfDocumentFormsCompatibilityExtensions {
    /// <summary>Attempts to fill string form values with explicit form options.</summary>
    internal static PdfOperationResult<PdfDocument> FillResult(
        this PdfDocumentForms forms,
        IReadOnlyDictionary<string, string> fieldValues,
        PdfFormFillerOptions formOptions) {
        Guard.NotNull(forms, nameof(forms));
        return forms.FillResult(fieldValues, formOptions, readOptions: null);
    }

    /// <summary>Attempts to fill typed form values with explicit form options.</summary>
    internal static PdfOperationResult<PdfDocument> FillResult(
        this PdfDocumentForms forms,
        IReadOnlyDictionary<string, PdfFormFieldValue> fieldValues,
        PdfFormFillerOptions formOptions) {
        Guard.NotNull(forms, nameof(forms));
        return forms.FillResult(fieldValues, formOptions, readOptions: null);
    }

    /// <summary>Attempts to flatten form fields with explicit form options.</summary>
    internal static PdfOperationResult<PdfDocument> FlattenResult(
        this PdfDocumentForms forms,
        PdfFormFillerOptions formOptions) {
        Guard.NotNull(forms, nameof(forms));
        return forms.FlattenResult(formOptions, readOptions: null);
    }

    /// <summary>Attempts to fill and flatten string form values with explicit form options.</summary>
    internal static PdfOperationResult<PdfDocument> FillAndFlattenResult(
        this PdfDocumentForms forms,
        IReadOnlyDictionary<string, string> fieldValues,
        PdfFormFillerOptions formOptions) {
        Guard.NotNull(forms, nameof(forms));
        return forms.FillAndFlattenResult(fieldValues, formOptions, readOptions: null);
    }

    /// <summary>Attempts to fill and flatten typed form values with explicit form options.</summary>
    internal static PdfOperationResult<PdfDocument> FillAndFlattenResult(
        this PdfDocumentForms forms,
        IReadOnlyDictionary<string, PdfFormFieldValue> fieldValues,
        PdfFormFillerOptions formOptions) {
        Guard.NotNull(forms, nameof(forms));
        return forms.FillAndFlattenResult(fieldValues, formOptions, readOptions: null);
    }
}
