namespace OfficeIMO.Pdf;

/// <summary>Exports and imports AcroForm values through the shared reader and filler engines.</summary>
internal static class PdfFormData {
    /// <summary>Exports readable named fields, including multi-value choice fields.</summary>
    public static PdfFormDataSet Export(byte[] pdf, PdfReadOptions? options = null) {
        Guard.NotNull(pdf, nameof(pdf)); PdfReadDocument document = PdfReadDocument.Open(pdf, options); var fields = new List<PdfFormDataField>();
        foreach (PdfFormField field in document.FormFields) {
            if (string.IsNullOrEmpty(field.Name)) continue;
            IReadOnlyList<string> values = field.Values.Count > 0 ? field.Values : new[] { field.Value ?? string.Empty };
            fields.Add(new PdfFormDataField(field.Name!, values));
        }
        return new PdfFormDataSet(fields);
    }
    /// <summary>Exports readable fields as XFDF.</summary>
    public static string ExportXfdf(byte[] pdf, PdfReadOptions? options = null) => Export(pdf, options).ToXfdf();
    /// <summary>Imports typed form data through the validated full-rewrite filler.</summary>
    public static byte[] Import(byte[] pdf, PdfFormDataSet data, PdfFormFillerOptions? options = null) { Guard.NotNull(data, nameof(data)); return PdfFormFiller.FillFields(pdf, data.ToFieldValues(), options); }
    /// <summary>Imports XFDF through the validated full-rewrite filler.</summary>
    public static byte[] ImportXfdf(byte[] pdf, string xfdf, PdfFormFillerOptions? options = null) => Import(pdf, PdfFormDataSet.ParseXfdf(xfdf), options);
}
