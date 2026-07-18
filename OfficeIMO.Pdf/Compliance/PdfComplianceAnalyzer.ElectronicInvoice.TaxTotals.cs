namespace OfficeIMO.Pdf;

internal static partial class PdfComplianceAnalyzer {
    private static PdfComplianceRequirement BuildElectronicInvoiceXmlTaxTotalConsistencyRequirement(PdfOptions options) {
        IReadOnlyList<PdfEmbeddedFile> embeddedFiles = options.EmbeddedFiles;
        var diagnostics = new List<string>();
        for (int i = 0; i < embeddedFiles.Count; i++) {
            PdfEmbeddedFile file = embeddedFiles[i];
            if (!IsFacturXCiiAttachment(file, diagnostics)) {
                continue;
            }

            if (!PdfCiiDocumentHeaderInspector.TryReadTaxTotalConsistency(file, out PdfCiiTaxTotalConsistencyEvidence? evidence, out string? consistencyDiagnostic)) {
                return new PdfComplianceRequirement(
                    "einvoice-xml-tax-total-consistency",
                    "EN 16931 XML tax total consistency",
                    PdfComplianceRequirementStatus.Missing,
                    consistencyDiagnostic!);
            }

            if (!string.IsNullOrWhiteSpace(evidence!.ParseDiagnostic)) {
                return new PdfComplianceRequirement(
                    "einvoice-xml-tax-total-consistency",
                    "EN 16931 XML tax total consistency",
                    PdfComplianceRequirementStatus.Missing,
                    evidence.ParseDiagnostic!);
            }

            var missingFields = new List<string>();
            if (!evidence.TaxBasisBreakdownSum.HasValue) {
                missingFields.Add("ApplicableTradeTax BasisAmount");
            }

            if (!evidence.TaxCalculatedBreakdownSum.HasValue) {
                missingFields.Add("ApplicableTradeTax CalculatedAmount");
            }

            if (!evidence.TaxBasisTotalAmount.HasValue) {
                missingFields.Add("TaxBasisTotalAmount");
            }

            if (!evidence.TaxTotalAmount.HasValue) {
                missingFields.Add("TaxTotalAmount");
            }

            if (missingFields.Count > 0) {
                return new PdfComplianceRequirement(
                    "einvoice-xml-tax-total-consistency",
                    "EN 16931 XML tax total consistency",
                    PdfComplianceRequirementStatus.Missing,
                    "Set factur-x.xml tax breakdown and tax total amounts before Mustang validation: " + string.Join(", ", missingFields.ToArray()) + ".");
            }

            var mismatchDiagnostics = new List<string>();
            if (!evidence.TaxBasisBreakdownMatchesTotal) {
                mismatchDiagnostics.Add("ApplicableTradeTax BasisAmount sum must match TaxBasisTotalAmount.");
            }

            if (!evidence.TaxCalculatedBreakdownMatchesTotal) {
                mismatchDiagnostics.Add("ApplicableTradeTax CalculatedAmount sum must match TaxTotalAmount.");
            }

            if (evidence.NotSubjectHeaderBasisAmount.HasValue &&
                evidence.NotSubjectLineNetAmountSum.HasValue &&
                !evidence.NotSubjectHeaderBasisMatchesLineNetAmount) {
                mismatchDiagnostics.Add("Category O ApplicableTradeTax BasisAmount must match category O line net amounts minus document-level category O allowances plus category O charges.");
            }

            if (!evidence.AdjustedBasisAmountsMatch) {
                mismatchDiagnostics.Add("ApplicableTradeTax BasisAmount must match same category/rate line net amounts minus document-level allowances plus charges: " + string.Join("; ", evidence.AdjustedBasisMismatches.ToArray()) + ".");
            }

            if (mismatchDiagnostics.Count > 0) {
                return new PdfComplianceRequirement(
                    "einvoice-xml-tax-total-consistency",
                    "EN 16931 XML tax total consistency",
                    PdfComplianceRequirementStatus.Missing,
                    "Fix factur-x.xml tax totals before Mustang validation: " + string.Join(" ", mismatchDiagnostics.ToArray()));
            }

            return new PdfComplianceRequirement(
                "einvoice-xml-tax-total-consistency",
                "EN 16931 XML tax total consistency",
                PdfComplianceRequirementStatus.Satisfied,
                "The factur-x.xml CrossIndustryInvoice tax breakdown basis/calculated amounts match header tax total amounts and category/rate adjusted taxable basis, including category-O taxable basis, for e-invoice readiness.");
        }

        string diagnostic = diagnostics.Count == 0
            ? "Attach a canonical factur-x.xml CrossIndustryInvoice file before checking EN 16931 XML tax total consistency."
            : string.Join(" ", diagnostics.Distinct(StringComparer.Ordinal).ToArray());
        return new PdfComplianceRequirement(
            "einvoice-xml-tax-total-consistency",
            "EN 16931 XML tax total consistency",
            PdfComplianceRequirementStatus.Missing,
            diagnostic);
    }
}
