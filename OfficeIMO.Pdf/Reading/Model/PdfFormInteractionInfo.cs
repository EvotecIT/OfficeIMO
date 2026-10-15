namespace OfficeIMO.Pdf;

/// <summary>Authenticated form metadata for a viewer's field controls, without extracted page content.</summary>
public sealed class PdfFormInteractionInfo {
    internal PdfFormInteractionInfo(IReadOnlyList<PdfFormField> fields, bool hasXfa) {
        Fields = fields;
        HasXfa = hasXfa;
    }

    /// <summary>Fields, values, options, and widget geometry in the PDF's page coordinate space.</summary>
    public IReadOnlyList<PdfFormField> Fields { get; }

    /// <summary>Whether the AcroForm also contains XFA data.</summary>
    public bool HasXfa { get; }
}
