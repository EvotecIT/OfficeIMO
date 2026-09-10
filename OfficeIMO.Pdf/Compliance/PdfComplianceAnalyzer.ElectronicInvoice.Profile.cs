namespace OfficeIMO.Pdf;

internal static partial class PdfComplianceAnalyzer {
    private static PdfComplianceRequirement BuildElectronicInvoiceProfileConsistencyRequirement(PdfOptions options) {
        string? diagnostic = PdfElectronicInvoiceProfile.GetDiagnostic(options.ElectronicInvoiceMetadata, options.EmbeddedFiles);
        return new PdfComplianceRequirement(
            "einvoice-profile-consistency",
            "Invoice XML and XMP profile agreement",
            diagnostic == null ? PdfComplianceRequirementStatus.Satisfied : PdfComplianceRequirementStatus.Missing,
            diagnostic ?? "The unique invoice XML guideline and canonical XMP conformance level agree.");
    }

    private static PdfComplianceRequirement BuildReadbackInvoiceProfileConsistencyRequirement(
        PdfXmpMetadataInfo? xmp,
        IReadOnlyList<PdfExtractedAttachment>? attachments,
        System.Threading.CancellationToken cancellationToken) {
        const string id = "readback-einvoice-profile-consistency";
        const string name = "Readback invoice XML and XMP profile agreement";
        if (attachments == null)
            return new PdfComplianceRequirement(id, name, PdfComplianceRequirementStatus.Missing,
                "Analyze exact PDF bytes to compare the embedded invoice profile with its XMP declaration.");
        var files = new List<PdfEmbeddedFile>();
        var diagnostics = new List<string>();
        foreach (PdfExtractedAttachment attachment in attachments) {
            cancellationToken.ThrowIfCancellationRequested();
            if (TryCreateReadbackEmbeddedFile(attachment, diagnostics, cancellationToken, out PdfEmbeddedFile? file)) files.Add(file!);
        }
        PdfElectronicInvoiceMetadata? metadata = null;
        if (!string.IsNullOrWhiteSpace(xmp?.ElectronicInvoiceDocumentType) &&
            !string.IsNullOrWhiteSpace(xmp?.ElectronicInvoiceDocumentFileName) &&
            !string.IsNullOrWhiteSpace(xmp?.ElectronicInvoiceVersion) &&
            !string.IsNullOrWhiteSpace(xmp?.ElectronicInvoiceConformanceLevel)) {
            try {
                metadata = new PdfElectronicInvoiceMetadata(xmp!.ElectronicInvoiceDocumentType!,
                    xmp.ElectronicInvoiceDocumentFileName!, xmp.ElectronicInvoiceVersion!, xmp.ElectronicInvoiceConformanceLevel!);
            } catch (ArgumentException exception) {
                diagnostics.Add(exception.Message);
            }
        }
        string? diagnostic = PdfElectronicInvoiceProfile.GetDiagnostic(metadata, files);
        if (diagnostics.Count != 0) diagnostic = string.Join(" ", diagnostics) + " " + diagnostic;
        return new PdfComplianceRequirement(id, name,
            diagnostic == null ? PdfComplianceRequirementStatus.Satisfied : PdfComplianceRequirementStatus.Missing,
            diagnostic ?? "The saved PDF's unique invoice XML guideline and canonical XMP conformance level agree.");
    }
}
