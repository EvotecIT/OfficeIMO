
namespace OfficeIMO.Pdf;

public sealed partial class PdfOptions {
    /// <summary>
    /// Attaches a snapshot of an invoice document using the existing Factur-X PDF/A-3 groundwork.
    /// This does not render invoice content or certify agreement between visible content and XML.
    /// The profile is inferred from the XML unless explicitly supplied; conflicting declarations are rejected.
    /// The caller must still validate the resulting PDF and invoice together.
    /// </summary>
    public PdfOptions UseFacturXDocument(
        PdfCiiInvoiceDocument invoice,
        string? conformanceLevel = null,
        string version = "1.0",
        PdfAssociatedFileRelationship relationship = PdfAssociatedFileRelationship.Data,
        string? description = "Factur-X/ZUGFeRD invoice XML",
        PdfTextFallbackFeatures textFallbacks = PdfTextFallbackFeatures.DocumentFont) {
        Guard.NotNull(invoice, nameof(invoice));
        return UseFacturX(invoice.ToBytes(), conformanceLevel, version, relationship, description, textFallbacks);
    }
}
