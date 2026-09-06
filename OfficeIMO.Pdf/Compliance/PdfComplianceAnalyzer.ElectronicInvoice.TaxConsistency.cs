namespace OfficeIMO.Pdf;

internal static partial class PdfComplianceAnalyzer {
    private static PdfComplianceRequirement BuildElectronicInvoiceXmlTaxCategoryConsistencyRequirement(PdfOptions options) {
        IReadOnlyList<PdfEmbeddedFile> embeddedFiles = options.EmbeddedFiles;
        var diagnostics = new List<string>();
        for (int i = 0; i < embeddedFiles.Count; i++) {
            PdfEmbeddedFile file = embeddedFiles[i];
            if (!IsFacturXCiiAttachment(file, diagnostics)) {
                continue;
            }

            if (!PdfCiiDocumentHeaderInspector.TryReadTaxCategoryConsistency(file, out PdfCiiTaxCategoryConsistencyEvidence? evidence, out string? consistencyDiagnostic)) {
                return new PdfComplianceRequirement(
                    "einvoice-xml-tax-category-consistency",
                    "EN 16931 XML tax category consistency",
                    PdfComplianceRequirementStatus.Missing,
                    consistencyDiagnostic!);
            }

            var missingFields = new List<string>();
            if (!evidence!.HasHeaderApplicableTradeTax) {
                missingFields.Add("header ApplicableTradeTax");
            }

            if (!evidence.HasLineApplicableTradeTax) {
                missingFields.Add("line ApplicableTradeTax");
            }

            if (!evidence.HasHeaderTaxCategoryRate) {
                missingFields.Add("header ApplicableTradeTax CategoryCode and required RateApplicablePercent");
            }

            if (!evidence.HasLineTaxCategoryRate) {
                missingFields.Add("line ApplicableTradeTax CategoryCode and required RateApplicablePercent");
            }

            if (missingFields.Count > 0) {
                return new PdfComplianceRequirement(
                    "einvoice-xml-tax-category-consistency",
                    "EN 16931 XML tax category consistency",
                    PdfComplianceRequirementStatus.Missing,
                    "Set factur-x.xml tax category/rate markers before Mustang validation: " + string.Join(", ", missingFields.ToArray()) + ".");
            }

            if (!evidence.AllLineTaxCategoryRatesMatchHeaderBreakdown) {
                return new PdfComplianceRequirement(
                    "einvoice-xml-tax-category-consistency",
                    "EN 16931 XML tax category consistency",
                    PdfComplianceRequirementStatus.Missing,
                    "Match every line ApplicableTradeTax CategoryCode/RateApplicablePercent, or category O without a rate, to a header tax breakdown before Mustang validation. Unmatched line tax category/rate: " + string.Join(", ", evidence.UnmatchedLineTaxCategoryRates.ToArray()) + ".");
            }

            if (!evidence.AllAllowanceChargeTaxCategoryRatesMatchHeaderBreakdown) {
                return new PdfComplianceRequirement(
                    "einvoice-xml-tax-category-consistency",
                    "EN 16931 XML tax category consistency",
                    PdfComplianceRequirementStatus.Missing,
                    "Match every document-level allowance/charge CategoryTradeTax CategoryCode/RateApplicablePercent, or category O without a rate, to a header tax breakdown before Mustang validation. Unmatched allowance/charge tax category/rate: " + string.Join(", ", evidence.UnmatchedAllowanceChargeTaxCategoryRates.ToArray()) + ".");
            }

            return new PdfComplianceRequirement(
                "einvoice-xml-tax-category-consistency",
                "EN 16931 XML tax category consistency",
                PdfComplianceRequirementStatus.Satisfied,
                "The factur-x.xml CrossIndustryInvoice line and allowance/charge tax category/rate markers and category-O rate absence match the header tax breakdown markers for e-invoice readiness.");
        }

        string diagnostic = diagnostics.Count == 0
            ? "Attach a canonical factur-x.xml CrossIndustryInvoice file before checking EN 16931 XML tax category consistency."
            : string.Join(" ", diagnostics.Distinct(StringComparer.Ordinal).ToArray());
        return new PdfComplianceRequirement(
            "einvoice-xml-tax-category-consistency",
            "EN 16931 XML tax category consistency",
            PdfComplianceRequirementStatus.Missing,
            diagnostic);
    }
}
