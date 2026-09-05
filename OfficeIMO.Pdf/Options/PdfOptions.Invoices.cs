using OfficeIMO.Invoices;

namespace OfficeIMO.Pdf;

public sealed partial class PdfOptions {
    /// <summary>
    /// Attaches a snapshot of an invoice document using the existing Factur-X PDF/A-3 groundwork.
    /// This does not render invoice content or certify agreement between visible content and XML.
    /// The caller selects the declared profile and must validate the resulting PDF and invoice together.
    /// </summary>
    public PdfOptions UseFacturXDocument(
        CiiInvoiceDocument invoice,
        string conformanceLevel = "EN 16931",
        string version = "1.0",
        PdfAssociatedFileRelationship relationship = PdfAssociatedFileRelationship.Data,
        string? description = "Factur-X/ZUGFeRD invoice XML",
        PdfTextFallbackFeatures textFallbacks = PdfTextFallbackFeatures.DocumentFont) {
        Guard.NotNull(invoice, nameof(invoice));
        return UseFacturX(invoice.ToBytes(), conformanceLevel, version, relationship, description, textFallbacks);
    }
}
