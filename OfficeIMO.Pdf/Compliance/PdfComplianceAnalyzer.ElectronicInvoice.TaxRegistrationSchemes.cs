namespace OfficeIMO.Pdf;

internal static partial class PdfComplianceAnalyzer {
    private static PdfComplianceRequirement BuildElectronicInvoiceXmlPartyTaxRegistrationSchemeRequirement(PdfOptions options) {
        IReadOnlyList<PdfEmbeddedFile> embeddedFiles = options.EmbeddedFiles;
        var diagnostics = new List<string>();
        for (int i = 0; i < embeddedFiles.Count; i++) {
            PdfEmbeddedFile file = embeddedFiles[i];
            if (!IsFacturXCiiAttachment(file, diagnostics)) {
                continue;
            }

            if (!PdfCiiDocumentHeaderInspector.TryReadPartyTaxRegistrationSchemes(file, out PdfCiiPartyTaxRegistrationSchemeEvidence? evidence, out string? schemeDiagnostic)) {
                return new PdfComplianceRequirement(
                    "einvoice-xml-party-tax-registration-scheme",
                    "EN 16931 XML party tax registration scheme",
                    PdfComplianceRequirementStatus.Missing,
                    schemeDiagnostic!);
            }

            var missingFields = new List<string>();
            if (!evidence!.HasSellerTaxRegistrationId) {
                missingFields.Add("SellerTradeParty SpecifiedTaxRegistration ID");
            } else if (!evidence.HasSellerTaxRegistrationSchemeId) {
                missingFields.Add("SellerTradeParty SpecifiedTaxRegistration ID schemeID");
            }

            if (evidence.HasBuyerTaxRegistrationId && !evidence.HasBuyerTaxRegistrationSchemeId) {
                missingFields.Add("BuyerTradeParty SpecifiedTaxRegistration ID schemeID");
            }

            if (missingFields.Count > 0) {
                return new PdfComplianceRequirement(
                    "einvoice-xml-party-tax-registration-scheme",
                    "EN 16931 XML party tax registration scheme",
                    PdfComplianceRequirementStatus.Missing,
                    "Set factur-x.xml tax registration scheme metadata before Mustang validation: " + string.Join(", ", missingFields.ToArray()) + ".");
            }

            return new PdfComplianceRequirement(
                "einvoice-xml-party-tax-registration-scheme",
                "EN 16931 XML party tax registration scheme",
                PdfComplianceRequirementStatus.Satisfied,
                "The factur-x.xml CrossIndustryInvoice tax registration identifiers include schemeID metadata where identifiers are present for e-invoice readiness.");
        }

        string diagnostic = diagnostics.Count == 0
            ? "Attach a canonical factur-x.xml CrossIndustryInvoice file before checking EN 16931 XML party tax registration schemes."
            : string.Join(" ", diagnostics.Distinct(StringComparer.Ordinal).ToArray());
        return new PdfComplianceRequirement(
            "einvoice-xml-party-tax-registration-scheme",
            "EN 16931 XML party tax registration scheme",
            PdfComplianceRequirementStatus.Missing,
            diagnostic);
    }
}
