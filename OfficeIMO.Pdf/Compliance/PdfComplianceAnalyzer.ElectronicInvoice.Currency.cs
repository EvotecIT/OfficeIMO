namespace OfficeIMO.Pdf;

internal static partial class PdfComplianceAnalyzer {
    private static PdfComplianceRequirement BuildElectronicInvoiceXmlCurrencyConsistencyRequirement(PdfOptions options) {
        IReadOnlyList<PdfEmbeddedFile> embeddedFiles = options.EmbeddedFiles;
        var diagnostics = new List<string>();
        for (int i = 0; i < embeddedFiles.Count; i++) {
            PdfEmbeddedFile file = embeddedFiles[i];
            if (!IsFacturXCiiAttachment(file, diagnostics)) {
                continue;
            }

            if (!PdfCiiDocumentHeaderInspector.TryReadCurrencyConsistency(file, out PdfCiiCurrencyConsistencyEvidence? evidence, out string? consistencyDiagnostic)) {
                return new PdfComplianceRequirement(
                    "einvoice-xml-currency-consistency",
                    "EN 16931 XML currency consistency",
                    PdfComplianceRequirementStatus.Missing,
                    consistencyDiagnostic!);
            }

            var missingFields = new List<string>();
            if (!evidence!.HasInvoiceCurrencyCode) {
                missingFields.Add("InvoiceCurrencyCode");
            }

            if (!evidence.HasCurrencyAmount) {
                missingFields.Add("currency-bearing monetary amounts");
            }

            if (missingFields.Count > 0) {
                return new PdfComplianceRequirement(
                    "einvoice-xml-currency-consistency",
                    "EN 16931 XML currency consistency",
                    PdfComplianceRequirementStatus.Missing,
                    "Set factur-x.xml invoice currency markers before Mustang validation: " + string.Join(", ", missingFields.ToArray()) + ".");
            }

            if (evidence.MismatchedAmountCurrencyFields.Count > 0) {
                return new PdfComplianceRequirement(
                    "einvoice-xml-currency-consistency",
                    "EN 16931 XML currency consistency",
                    PdfComplianceRequirementStatus.Missing,
                    "Set factur-x.xml monetary amount currencyID values to match InvoiceCurrencyCode " + evidence.InvoiceCurrencyCode + ": " + string.Join(", ", evidence.MismatchedAmountCurrencyFields.ToArray()) + ".");
            }

            return new PdfComplianceRequirement(
                "einvoice-xml-currency-consistency",
                "EN 16931 XML currency consistency",
                PdfComplianceRequirementStatus.Satisfied,
                "The factur-x.xml CrossIndustryInvoice monetary amount currencyID values, when present, match the invoice currency code for e-invoice readiness.");
        }

        string diagnostic = diagnostics.Count == 0
            ? "Attach a canonical factur-x.xml CrossIndustryInvoice file before checking EN 16931 XML currency consistency."
            : string.Join(" ", diagnostics.Distinct(StringComparer.Ordinal).ToArray());
        return new PdfComplianceRequirement(
            "einvoice-xml-currency-consistency",
            "EN 16931 XML currency consistency",
            PdfComplianceRequirementStatus.Missing,
            diagnostic);
    }
}
