namespace OfficeIMO.Pdf;

internal static partial class PdfComplianceAnalyzer {
    private static PdfComplianceRequirement BuildElectronicInvoiceXmlTaxCategoryRateRequirement(PdfOptions options) {
        IReadOnlyList<PdfEmbeddedFile> embeddedFiles = options.EmbeddedFiles;
        var diagnostics = new List<string>();
        for (int i = 0; i < embeddedFiles.Count; i++) {
            PdfEmbeddedFile file = embeddedFiles[i];
            if (!IsFacturXCiiAttachment(file, diagnostics)) {
                continue;
            }

            if (!PdfCiiDocumentHeaderInspector.TryReadTaxCategoryRates(file, out PdfCiiTaxCategoryRateEvidence? evidence, out string? taxDiagnostic)) {
                return new PdfComplianceRequirement(
                    "einvoice-xml-tax-category-rate",
                    "EN 16931 XML tax category rate",
                    PdfComplianceRequirementStatus.Missing,
                    taxDiagnostic!);
            }

            var missingFields = new List<string>();
            if (!evidence!.HasApplicableTradeTax) {
                missingFields.Add("ApplicableTradeTax");
            }

            if (!evidence.HasTaxCategoryRate && !evidence.HasRateRequirementCoverage && evidence.ForbiddenRateCategoryCodes.Count == 0) {
                missingFields.Add("ApplicableTradeTax CategoryCode and RateApplicablePercent");
            }

            if (missingFields.Count > 0) {
                return new PdfComplianceRequirement(
                    "einvoice-xml-tax-category-rate",
                    "EN 16931 XML tax category rate",
                    PdfComplianceRequirementStatus.Missing,
                    "Set factur-x.xml tax category/rate markers before Mustang validation: " + string.Join(", ", missingFields.ToArray()) + ".");
            }

            if (!string.IsNullOrWhiteSpace(evidence.ParseDiagnostic)) {
                return new PdfComplianceRequirement(
                    "einvoice-xml-tax-category-rate",
                    "EN 16931 XML tax category rate",
                    PdfComplianceRequirementStatus.Missing,
                    evidence.ParseDiagnostic!);
            }

            if (evidence.MissingRateCategoryCodes.Count > 0) {
                return new PdfComplianceRequirement(
                    "einvoice-xml-tax-category-rate",
                    "EN 16931 XML tax category rate",
                    PdfComplianceRequirementStatus.Missing,
                    "Set factur-x.xml RateApplicablePercent for every non-O ApplicableTradeTax before Mustang validation. Missing tax category rate categories: " + string.Join(", ", evidence.MissingRateCategoryCodes.ToArray()) + ".");
            }

            if (evidence.ForbiddenRateCategoryCodes.Count > 0) {
                return new PdfComplianceRequirement(
                    "einvoice-xml-tax-category-rate",
                    "EN 16931 XML tax category rate",
                    PdfComplianceRequirementStatus.Missing,
                    "Remove factur-x.xml ApplicableTradeTax RateApplicablePercent for Peppol category O before Mustang validation. Forbidden tax category rate categories: " + string.Join(", ", evidence.ForbiddenRateCategoryCodes.ToArray()) + ".");
            }

            if (!evidence.AllZeroRatedCategoriesUseZeroRate) {
                return new PdfComplianceRequirement(
                    "einvoice-xml-tax-category-rate",
                    "EN 16931 XML tax category rate",
                    PdfComplianceRequirementStatus.Missing,
                    "Set VAT category rates that Peppol requires to be zero to 0 before Mustang validation. Non-zero category/rate values: " + string.Join(", ", evidence.NonZeroRatedCategoryRates.ToArray()) + ".");
            }

            return new PdfComplianceRequirement(
                "einvoice-xml-tax-category-rate",
                "EN 16931 XML tax category rate",
                PdfComplianceRequirementStatus.Satisfied,
                "The factur-x.xml CrossIndustryInvoice VAT category rates satisfy zero-rate semantics for AE, E, G, K, and Z categories and rate absence for category O header, line, allowance, and charge tax markers.");
        }

        string diagnostic = diagnostics.Count == 0
            ? "Attach a canonical factur-x.xml CrossIndustryInvoice file before checking EN 16931 XML tax category rates."
            : string.Join(" ", diagnostics.Distinct(StringComparer.Ordinal).ToArray());
        return new PdfComplianceRequirement(
            "einvoice-xml-tax-category-rate",
            "EN 16931 XML tax category rate",
            PdfComplianceRequirementStatus.Missing,
            diagnostic);
    }
}
