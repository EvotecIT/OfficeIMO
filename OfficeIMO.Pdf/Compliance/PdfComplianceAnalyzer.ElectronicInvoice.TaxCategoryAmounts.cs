namespace OfficeIMO.Pdf;

internal static partial class PdfComplianceAnalyzer {
    private static PdfComplianceRequirement BuildElectronicInvoiceXmlTaxCategoryAmountRequirement(PdfOptions options) {
        IReadOnlyList<PdfEmbeddedFile> embeddedFiles = options.EmbeddedFiles;
        var diagnostics = new List<string>();
        for (int i = 0; i < embeddedFiles.Count; i++) {
            PdfEmbeddedFile file = embeddedFiles[i];
            if (!IsFacturXCiiAttachment(file, diagnostics)) {
                continue;
            }

            if (!PdfCiiDocumentHeaderInspector.TryReadTaxCategoryAmounts(file, out PdfCiiTaxCategoryAmountEvidence? evidence, out string? taxDiagnostic)) {
                return new PdfComplianceRequirement(
                    "einvoice-xml-tax-category-amount",
                    "EN 16931 XML tax category amount",
                    PdfComplianceRequirementStatus.Missing,
                    taxDiagnostic!);
            }

            var missingFields = new List<string>();
            if (!evidence!.HasApplicableTradeTax) {
                missingFields.Add("ApplicableTradeTax");
            }

            if (!evidence.HasTaxCategoryAmount) {
                missingFields.Add("ApplicableTradeTax CategoryCode and CalculatedAmount");
            }

            if (missingFields.Count > 0) {
                return new PdfComplianceRequirement(
                    "einvoice-xml-tax-category-amount",
                    "EN 16931 XML tax category amount",
                    PdfComplianceRequirementStatus.Missing,
                    "Set factur-x.xml tax category/amount markers before Mustang validation: " + string.Join(", ", missingFields.ToArray()) + ".");
            }

            if (!string.IsNullOrWhiteSpace(evidence.ParseDiagnostic)) {
                return new PdfComplianceRequirement(
                    "einvoice-xml-tax-category-amount",
                    "EN 16931 XML tax category amount",
                    PdfComplianceRequirementStatus.Missing,
                    evidence.ParseDiagnostic!);
            }

            if (!evidence.AllZeroRatedCategoriesUseZeroAmount) {
                return new PdfComplianceRequirement(
                    "einvoice-xml-tax-category-amount",
                    "EN 16931 XML tax category amount",
                    PdfComplianceRequirementStatus.Missing,
                    "Set VAT category tax amounts that Peppol requires to be zero to 0 before Mustang validation. Non-zero category/amount values: " + string.Join(", ", evidence.NonZeroRatedCategoryAmounts.ToArray()) + ".");
            }

            if (!evidence.AllStandardRatedCategoryAmountsMatchRate) {
                return new PdfComplianceRequirement(
                    "einvoice-xml-tax-category-amount",
                    "EN 16931 XML tax category amount",
                    PdfComplianceRequirementStatus.Missing,
                    "Set standard-rated VAT category tax amounts to taxable basis multiplied by VAT rate before Mustang validation. Mismatched category/rate tax amounts: " + string.Join("; ", evidence.MismatchedStandardRatedCategoryAmounts.ToArray()) + ".");
            }

            return new PdfComplianceRequirement(
                "einvoice-xml-tax-category-amount",
                "EN 16931 XML tax category amount",
                PdfComplianceRequirementStatus.Satisfied,
                "The factur-x.xml CrossIndustryInvoice VAT category tax amounts satisfy zero-amount semantics for AE, E, G, K, O, and Z categories and standard-rated taxable-basis times rate semantics.");
        }

        string diagnostic = diagnostics.Count == 0
            ? "Attach a canonical factur-x.xml CrossIndustryInvoice file before checking EN 16931 XML tax category amounts."
            : string.Join(" ", diagnostics.Distinct(StringComparer.Ordinal).ToArray());
        return new PdfComplianceRequirement(
            "einvoice-xml-tax-category-amount",
            "EN 16931 XML tax category amount",
            PdfComplianceRequirementStatus.Missing,
            diagnostic);
    }
}
