namespace OfficeIMO.Pdf;

internal static partial class PdfComplianceAnalyzer {
    private static PdfComplianceRequirement BuildElectronicInvoiceXmlTaxPartyIdentifierRequirement(PdfOptions options) {
        IReadOnlyList<PdfEmbeddedFile> embeddedFiles = options.EmbeddedFiles;
        var diagnostics = new List<string>();
        for (int i = 0; i < embeddedFiles.Count; i++) {
            PdfEmbeddedFile file = embeddedFiles[i];
            if (!IsFacturXCiiAttachment(file, diagnostics)) {
                continue;
            }

            if (!PdfCiiDocumentHeaderInspector.TryReadTaxPartyIdentifiers(file, out PdfCiiTaxPartyIdentifierEvidence? evidence, out string? taxDiagnostic)) {
                return new PdfComplianceRequirement(
                    "einvoice-xml-tax-party-identifiers",
                    "EN 16931 XML tax party identifiers",
                    PdfComplianceRequirementStatus.Missing,
                    taxDiagnostic!);
            }

            if (!evidence!.HasApplicableTradeTax) {
                return new PdfComplianceRequirement(
                    "einvoice-xml-tax-party-identifiers",
                    "EN 16931 XML tax party identifiers",
                    PdfComplianceRequirementStatus.Missing,
                    "Set factur-x.xml ApplicableTradeTax before checking category-specific party tax identifiers.");
            }

            var missingDiagnostics = new List<string>();
            if (evidence.MissingSellerIdentifierCategories.Count > 0) {
                missingDiagnostics.Add("seller VAT/tax identifier for categories " + string.Join(", ", evidence.MissingSellerIdentifierCategories.ToArray()));
            }

            if (evidence.MissingBuyerIdentifierCategories.Count > 0) {
                missingDiagnostics.Add("buyer VAT/tax identifier for categories " + string.Join(", ", evidence.MissingBuyerIdentifierCategories.ToArray()));
            }

            if (missingDiagnostics.Count > 0) {
                return new PdfComplianceRequirement(
                    "einvoice-xml-tax-party-identifiers",
                    "EN 16931 XML tax party identifiers",
                    PdfComplianceRequirementStatus.Missing,
                    "Set factur-x.xml party tax identifiers required by Peppol category rules before Mustang validation: " + string.Join("; ", missingDiagnostics.ToArray()) + ".");
            }

            var forbiddenDiagnostics = new List<string>();
            if (evidence.ForbiddenSellerVatIdentifierCategories.Count > 0) {
                forbiddenDiagnostics.Add("seller VAT identifier for categories " + string.Join(", ", evidence.ForbiddenSellerVatIdentifierCategories.ToArray()));
            }

            if (evidence.ForbiddenBuyerVatIdentifierCategories.Count > 0) {
                forbiddenDiagnostics.Add("buyer VAT identifier for categories " + string.Join(", ", evidence.ForbiddenBuyerVatIdentifierCategories.ToArray()));
            }

            if (forbiddenDiagnostics.Count > 0) {
                return new PdfComplianceRequirement(
                    "einvoice-xml-tax-party-identifiers",
                    "EN 16931 XML tax party identifiers",
                    PdfComplianceRequirementStatus.Missing,
                    "Remove factur-x.xml party VAT identifiers forbidden by Peppol category O before Mustang validation: " + string.Join("; ", forbiddenDiagnostics.ToArray()) + ".");
            }

            return new PdfComplianceRequirement(
                "einvoice-xml-tax-party-identifiers",
                "EN 16931 XML tax party identifiers",
                PdfComplianceRequirementStatus.Satisfied,
                "The factur-x.xml CrossIndustryInvoice seller and buyer VAT/tax registration markers satisfy category-specific Peppol readiness for AE, E, G, K, and O e-invoice categories.");
        }

        string diagnostic = diagnostics.Count == 0
            ? "Attach a canonical factur-x.xml CrossIndustryInvoice file before checking EN 16931 XML tax party identifiers."
            : string.Join(" ", diagnostics.Distinct(StringComparer.Ordinal).ToArray());
        return new PdfComplianceRequirement(
            "einvoice-xml-tax-party-identifiers",
            "EN 16931 XML tax party identifiers",
            PdfComplianceRequirementStatus.Missing,
            diagnostic);
    }
}
