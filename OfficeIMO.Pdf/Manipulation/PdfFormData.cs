namespace OfficeIMO.Pdf;

/// <summary>Exports and imports AcroForm values through the shared reader and filler engines.</summary>
internal static class PdfFormData {
    /// <summary>Exports readable named fields, including multi-value choice fields.</summary>
    public static PdfFormDataSet Export(byte[] pdf, PdfLoadOptions? options = null) {
        Guard.NotNull(pdf, nameof(pdf)); PdfReadDocument document = PdfReadDocument.Open(pdf, options); var fields = new List<PdfFormDataField>();
        foreach (PdfFormField field in document.FormFields) {
            if (field.IsNoExport || string.IsNullOrEmpty(field.Name)) continue;
            IReadOnlyList<string> values = ResolveExportValues(field);
            fields.Add(new PdfFormDataField(field.Name!, values));
        }
        return new PdfFormDataSet(fields);
    }
    private static IReadOnlyList<string> ResolveExportValues(PdfFormField field) {
        if (field.IsButtonField
            && !string.Equals(field.Value, "Off", StringComparison.Ordinal)
            && field.Options.Count == field.Widgets.Count) {
            for (int index = 0; index < field.Widgets.Count; index++) {
                if (string.Equals(field.Widgets[index].AppearanceState, field.Value, StringComparison.Ordinal)) {
                    return new[] { field.Options[index].ExportValue };
                }
            }
        }
        return field.Values.Count > 0 ? field.Values : new[] { field.Value ?? string.Empty };
    }
    /// <summary>Exports readable fields as XFDF.</summary>
    public static string ExportXfdf(byte[] pdf, PdfLoadOptions? options = null) => Export(pdf, options).ToXfdf();
    /// <summary>Imports typed form data through the validated full-rewrite filler.</summary>
    public static byte[] Import(byte[] pdf, PdfFormDataSet data, PdfFormFillerOptions? options = null) => Import(pdf, data, options, readOptions: null);
    internal static byte[] Import(byte[] pdf, PdfFormDataSet data, PdfFormFillerOptions? options, PdfLoadOptions? readOptions) { Guard.NotNull(data, nameof(data)); return PdfFormFiller.FillFields(pdf, data.ToFieldValues(), options, readOptions); }
    /// <summary>Imports XFDF through the validated full-rewrite filler.</summary>
    public static byte[] ImportXfdf(byte[] pdf, string xfdf, PdfFormFillerOptions? options = null) => ImportXfdf(pdf, xfdf, options, readOptions: null);
    internal static byte[] ImportXfdf(byte[] pdf, string xfdf, PdfFormFillerOptions? options, PdfLoadOptions? readOptions) {
        PdfFormFillerOptions effective = options ?? new PdfFormFillerOptions();
        PdfFormDataSet data = PdfFormDataSet.ParseXfdf(
            xfdf,
            effective.MaxXfdfFields,
            effective.MaxXfdfValueCharacters,
            effective.MaxXfdfDocumentCharacters);
        return Import(pdf, data, effective, readOptions);
    }
}
