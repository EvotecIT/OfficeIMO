namespace OfficeIMO.Pdf;

internal static partial class PdfComplianceAnalyzer {
    private static PdfComplianceRequirement BuildElectronicInvoiceXmlTaxExemptionReasonRequirement(PdfOptions options) {
        IReadOnlyList<PdfEmbeddedFile> embeddedFiles = options.EmbeddedFiles;
        var diagnostics = new List<string>();
        for (int i = 0; i < embeddedFiles.Count; i++) {
            PdfEmbeddedFile file = embeddedFiles[i];
            if (!IsFacturXCiiAttachment(file, diagnostics)) {
                continue;
            }

            if (!PdfCiiDocumentHeaderInspector.TryReadTaxExemptionReasons(file, out PdfCiiTaxExemptionReasonEvidence? evidence, out string? taxDiagnostic)) {
                return new PdfComplianceRequirement(
                    "einvoice-xml-tax-exemption-reason",
                    "EN 16931 XML tax exemption reason",
                    PdfComplianceRequirementStatus.Missing,
                    taxDiagnostic!);
            }

            if (!evidence!.HasApplicableTradeTax) {
                return new PdfComplianceRequirement(
                    "einvoice-xml-tax-exemption-reason",
                    "EN 16931 XML tax exemption reason",
                    PdfComplianceRequirementStatus.Missing,
                    "Set factur-x.xml ApplicableTradeTax before checking tax exemption reason markers.");
            }

            if (!evidence.AllRequiredCategoriesHaveReason) {
                return new PdfComplianceRequirement(
                    "einvoice-xml-tax-exemption-reason",
                    "EN 16931 XML tax exemption reason",
                    PdfComplianceRequirementStatus.Missing,
                    "Set factur-x.xml ExemptionReason or ExemptionReasonCode for VAT categories that Peppol requires to explain before Mustang validation. Missing categories: " + string.Join(", ", evidence.MissingReasonCategories.ToArray()) + ".");
            }

            if (!evidence.AllForbiddenCategoriesOmitReason) {
                return new PdfComplianceRequirement(
                    "einvoice-xml-tax-exemption-reason",
                    "EN 16931 XML tax exemption reason",
                    PdfComplianceRequirementStatus.Missing,
                    "Remove factur-x.xml ExemptionReason and ExemptionReasonCode from VAT categories where Peppol forbids exemption reasons before Mustang validation. Categories with reason markers: " + string.Join(", ", evidence.ForbiddenReasonCategories.ToArray()) + ".");
            }

            if (!evidence.AllNotSubjectReasonCodesAreCanonical) {
                return new PdfComplianceRequirement(
                    "einvoice-xml-tax-exemption-reason",
                    "EN 16931 XML tax exemption reason",
                    PdfComplianceRequirementStatus.Missing,
                    "Set factur-x.xml category O ExemptionReasonCode to VATEX-EU-O when a reason code is used before Mustang validation. Found: " + string.Join(", ", evidence.InvalidNotSubjectReasonCodes.ToArray()) + ".");
            }

            return new PdfComplianceRequirement(
                "einvoice-xml-tax-exemption-reason",
                "EN 16931 XML tax exemption reason",
                PdfComplianceRequirementStatus.Satisfied,
                "The factur-x.xml CrossIndustryInvoice VAT exemption reason markers satisfy Peppol reason presence rules for AE, E, G, K, O, S, Z, L, and M e-invoice categories, including canonical VATEX-EU-O reason-code semantics for category O.");
        }

        string diagnostic = diagnostics.Count == 0
            ? "Attach a canonical factur-x.xml CrossIndustryInvoice file before checking EN 16931 XML tax exemption reasons."
            : string.Join(" ", diagnostics.Distinct(StringComparer.Ordinal).ToArray());
        return new PdfComplianceRequirement(
            "einvoice-xml-tax-exemption-reason",
            "EN 16931 XML tax exemption reason",
            PdfComplianceRequirementStatus.Missing,
            diagnostic);
    }
}
