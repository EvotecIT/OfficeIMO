namespace OfficeIMO.Pdf;

internal static partial class PdfComplianceAnalyzer {
    private static PdfComplianceRequirement BuildElectronicInvoiceXmlLineAmountConsistencyRequirement(PdfOptions options) {
        IReadOnlyList<PdfEmbeddedFile> embeddedFiles = options.EmbeddedFiles;
        var diagnostics = new List<string>();
        for (int i = 0; i < embeddedFiles.Count; i++) {
            PdfEmbeddedFile file = embeddedFiles[i];
            if (!IsFacturXCiiAttachment(file, diagnostics)) {
                continue;
            }

            if (!PdfCiiDocumentHeaderInspector.TryReadLineAmountConsistency(file, out PdfCiiLineAmountConsistencyEvidence? evidence, out string? consistencyDiagnostic)) {
                return new PdfComplianceRequirement(
                    "einvoice-xml-line-amount-consistency",
                    "EN 16931 XML line amount consistency",
                    PdfComplianceRequirementStatus.Missing,
                    consistencyDiagnostic!);
            }

            if (!string.IsNullOrWhiteSpace(evidence!.ParseDiagnostic)) {
                return new PdfComplianceRequirement(
                    "einvoice-xml-line-amount-consistency",
                    "EN 16931 XML line amount consistency",
                    PdfComplianceRequirementStatus.Missing,
                    evidence.ParseDiagnostic!);
            }

            var missingFields = new List<string>();
            if (!evidence.HasIncludedSupplyChainTradeLineItem) {
                missingFields.Add("IncludedSupplyChainTradeLineItem");
            }

            if (!evidence.HasBilledQuantity) {
                missingFields.Add("BilledQuantity");
            }

            if (!evidence.HasPriceChargeAmount) {
                missingFields.Add("ProductTradePrice ChargeAmount");
            }

            if (!evidence.HasLineTotalAmount) {
                missingFields.Add("LineTotalAmount");
            }

            if (missingFields.Count > 0) {
                return new PdfComplianceRequirement(
                    "einvoice-xml-line-amount-consistency",
                    "EN 16931 XML line amount consistency",
                    PdfComplianceRequirementStatus.Missing,
                    "Set factur-x.xml line quantity, price, and total amounts before Mustang validation: " + string.Join(", ", missingFields.ToArray()) + ".");
            }

            if (!evidence.AllLineAmountsMatch) {
                return new PdfComplianceRequirement(
                    "einvoice-xml-line-amount-consistency",
                    "EN 16931 XML line amount consistency",
                    PdfComplianceRequirementStatus.Missing,
                    "Fix factur-x.xml line totals before Mustang validation: BilledQuantity times ProductTradePrice ChargeAmount divided by BasisQuantity, when present, must match LineTotalAmount for line(s): " + string.Join(", ", evidence.MismatchedLineIds.ToArray()) + ".");
            }

            return new PdfComplianceRequirement(
                "einvoice-xml-line-amount-consistency",
                "EN 16931 XML line amount consistency",
                PdfComplianceRequirementStatus.Satisfied,
                "The factur-x.xml CrossIndustryInvoice line quantity, price, and line total amount markers are internally consistent for e-invoice readiness.");
        }

        string diagnostic = diagnostics.Count == 0
            ? "Attach a canonical factur-x.xml CrossIndustryInvoice file before checking EN 16931 XML line amount consistency."
            : string.Join(" ", diagnostics.Distinct(StringComparer.Ordinal).ToArray());
        return new PdfComplianceRequirement(
            "einvoice-xml-line-amount-consistency",
            "EN 16931 XML line amount consistency",
            PdfComplianceRequirementStatus.Missing,
            diagnostic);
    }
}
