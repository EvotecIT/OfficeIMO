namespace OfficeIMO.Pdf;

internal static partial class PdfComplianceAnalyzer {
    private static readonly string[] ElectronicInvoiceTaxCategoryCodes = {
        "AE",
        "E",
        "S",
        "Z",
        "G",
        "O",
        "K",
        "L",
        "M",
        "B"
    };

    private static PdfComplianceRequirement BuildElectronicInvoiceXmlTaxCategoryCodeRequirement(PdfOptions options) {
        IReadOnlyList<PdfEmbeddedFile> embeddedFiles = options.EmbeddedFiles;
        var diagnostics = new List<string>();
        for (int i = 0; i < embeddedFiles.Count; i++) {
            PdfEmbeddedFile file = embeddedFiles[i];
            if (!IsFacturXCiiAttachment(file, diagnostics)) {
                continue;
            }

            if (!PdfCiiDocumentHeaderInspector.TryReadTaxCategoryCodes(file, out PdfCiiTaxCategoryCodeEvidence? evidence, out string? taxDiagnostic)) {
                return new PdfComplianceRequirement(
                    "einvoice-xml-tax-category-code",
                    "EN 16931 XML tax category code",
                    PdfComplianceRequirementStatus.Missing,
                    taxDiagnostic!);
            }

            var missingFields = new List<string>();
            if (!evidence!.HasApplicableTradeTax) {
                missingFields.Add("ApplicableTradeTax");
            }

            if (!evidence.HasCategoryCode) {
                missingFields.Add("ApplicableTradeTax CategoryCode");
            }

            if (missingFields.Count > 0) {
                return new PdfComplianceRequirement(
                    "einvoice-xml-tax-category-code",
                    "EN 16931 XML tax category code",
                    PdfComplianceRequirementStatus.Missing,
                    "Set factur-x.xml tax category code essentials before Mustang validation: " + string.Join(", ", missingFields.ToArray()) + ".");
            }

            var invalidCategoryCodes = new List<string>();
            for (int j = 0; j < evidence.CategoryCodes.Count; j++) {
                if (!IsKnownElectronicInvoiceTaxCategoryCode(evidence.CategoryCodes[j])) {
                    invalidCategoryCodes.Add(evidence.CategoryCodes[j]);
                }
            }

            if (invalidCategoryCodes.Count > 0) {
                return new PdfComplianceRequirement(
                    "einvoice-xml-tax-category-code",
                    "EN 16931 XML tax category code",
                    PdfComplianceRequirementStatus.Missing,
                    "Set factur-x.xml ApplicableTradeTax CategoryCode to a Peppol UNCL5305 duty or tax category code before Mustang validation. Found: " + string.Join(", ", invalidCategoryCodes.Distinct(StringComparer.Ordinal).ToArray()) + ".");
            }

            if ((evidence.HasLineNotSubjectTaxCategory || evidence.HasAllowanceChargeNotSubjectTaxCategory) &&
                evidence.HeaderNotSubjectTaxBreakdownCount != 1) {
                return new PdfComplianceRequirement(
                    "einvoice-xml-tax-category-code",
                    "EN 16931 XML tax category code",
                    PdfComplianceRequirementStatus.Missing,
                    "Set exactly one factur-x.xml header ApplicableTradeTax CategoryCode O breakdown when a line, document-level allowance, or document-level charge uses Peppol category O before Mustang validation. Header category O breakdown count: " + evidence.HeaderNotSubjectTaxBreakdownCount.ToString(System.Globalization.CultureInfo.InvariantCulture) + ".");
            }

            if (evidence.HeaderNotSubjectTaxBreakdownCount > 1) {
                return new PdfComplianceRequirement(
                    "einvoice-xml-tax-category-code",
                    "EN 16931 XML tax category code",
                    PdfComplianceRequirementStatus.Missing,
                    "Keep exactly one factur-x.xml header ApplicableTradeTax CategoryCode O breakdown before Mustang validation. Header category O breakdown count: " + evidence.HeaderNotSubjectTaxBreakdownCount.ToString(System.Globalization.CultureInfo.InvariantCulture) + ".");
            }

            if (evidence.HeaderNotSubjectTaxBreakdownCount > 0 && evidence.NonNotSubjectHeaderTaxBreakdownCategories.Count > 0) {
                return new PdfComplianceRequirement(
                    "einvoice-xml-tax-category-code",
                    "EN 16931 XML tax category code",
                    PdfComplianceRequirementStatus.Missing,
                    "Do not combine a factur-x.xml header ApplicableTradeTax CategoryCode O breakdown with other VAT breakdown categories before Mustang validation. Other header tax categories: " + string.Join(", ", evidence.NonNotSubjectHeaderTaxBreakdownCategories.ToArray()) + ".");
            }

            return new PdfComplianceRequirement(
                "einvoice-xml-tax-category-code",
                "EN 16931 XML tax category code",
                PdfComplianceRequirementStatus.Satisfied,
                "The factur-x.xml CrossIndustryInvoice tax category codes are from the Peppol UNCL5305 duty or tax category code list, including category-O line, allowance, and charge breakdown exclusivity, for e-invoice readiness.");
        }

        string diagnostic = diagnostics.Count == 0
            ? "Attach a canonical factur-x.xml CrossIndustryInvoice file before checking EN 16931 XML tax category codes."
            : string.Join(" ", diagnostics.Distinct(StringComparer.Ordinal).ToArray());
        return new PdfComplianceRequirement(
            "einvoice-xml-tax-category-code",
            "EN 16931 XML tax category code",
            PdfComplianceRequirementStatus.Missing,
            diagnostic);
    }

    private static bool IsKnownElectronicInvoiceTaxCategoryCode(string categoryCode) {
        string normalized = categoryCode.Trim().ToUpperInvariant();
        for (int i = 0; i < ElectronicInvoiceTaxCategoryCodes.Length; i++) {
            if (string.Equals(normalized, ElectronicInvoiceTaxCategoryCodes[i], StringComparison.Ordinal)) {
                return true;
            }
        }

        return false;
    }
}
