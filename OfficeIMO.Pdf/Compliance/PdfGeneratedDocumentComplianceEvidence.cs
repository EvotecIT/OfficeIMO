namespace OfficeIMO.Pdf;

internal sealed class PdfGeneratedDocumentComplianceEvidence {
    public PdfGeneratedDocumentComplianceEvidence(
        System.Collections.Generic.IReadOnlyList<PdfStandardFont> standardFonts,
        System.Collections.Generic.IReadOnlyList<PdfGeneratedImageAccessibilityEvidence> images,
        System.Collections.Generic.IReadOnlyList<PdfGeneratedDrawingAccessibilityEvidence> drawings,
        System.Collections.Generic.IReadOnlyList<PdfGeneratedFormAccessibilityEvidence> forms) {
        StandardFonts = standardFonts;
        Images = images;
        Drawings = drawings;
        Forms = forms;
    }

    public System.Collections.Generic.IReadOnlyList<PdfStandardFont> StandardFonts { get; }

    public System.Collections.Generic.IReadOnlyList<PdfGeneratedImageAccessibilityEvidence> Images { get; }

    public System.Collections.Generic.IReadOnlyList<PdfGeneratedDrawingAccessibilityEvidence> Drawings { get; }

    public System.Collections.Generic.IReadOnlyList<PdfGeneratedFormAccessibilityEvidence> Forms { get; }
}
