namespace OfficeIMO.Pdf;

internal static partial class PdfComplianceAnalyzer {
    private static PdfComplianceRequirement BuildElectronicInvoiceXmlAllowanceChargeReasonRequirement(PdfOptions options) {
        IReadOnlyList<PdfEmbeddedFile> embeddedFiles = options.EmbeddedFiles;
        var diagnostics = new List<string>();
        for (int i = 0; i < embeddedFiles.Count; i++) {
            PdfEmbeddedFile file = embeddedFiles[i];
            if (!IsFacturXCiiAttachment(file, diagnostics)) {
                continue;
            }

            if (!PdfCiiDocumentHeaderInspector.TryReadAllowanceChargeReasons(file, out PdfCiiAllowanceChargeReasonEvidence? evidence, out string? reasonDiagnostic)) {
                return new PdfComplianceRequirement(
                    "einvoice-xml-allowance-charge-reason",
                    "EN 16931 XML allowance/charge reasons",
                    PdfComplianceRequirementStatus.Missing,
                    reasonDiagnostic!);
            }

            if (!evidence!.AllAllowanceChargesHaveReason) {
                var missingFields = new List<string>();
                if (evidence.MissingAllowanceReasons.Count > 0) {
                    missingFields.Add("document-level allowance Reason or ReasonCode: " + string.Join(", ", evidence.MissingAllowanceReasons.ToArray()));
                }

                if (evidence.MissingChargeReasons.Count > 0) {
                    missingFields.Add("document-level charge Reason or ReasonCode: " + string.Join(", ", evidence.MissingChargeReasons.ToArray()));
                }

                return new PdfComplianceRequirement(
                    "einvoice-xml-allowance-charge-reason",
                    "EN 16931 XML allowance/charge reasons",
                    PdfComplianceRequirementStatus.Missing,
                    "Set factur-x.xml document-level allowance/charge reason markers before Mustang validation: " + string.Join("; ", missingFields.ToArray()) + ".");
            }

            return new PdfComplianceRequirement(
                "einvoice-xml-allowance-charge-reason",
                "EN 16931 XML allowance/charge reasons",
                PdfComplianceRequirementStatus.Satisfied,
                "The factur-x.xml CrossIndustryInvoice document-level allowances and charges include Reason or ReasonCode markers for e-invoice readiness.");
        }

        string diagnostic = diagnostics.Count == 0
            ? "Attach a canonical factur-x.xml CrossIndustryInvoice file before checking EN 16931 XML allowance/charge reason essentials."
            : string.Join(" ", diagnostics.Distinct(StringComparer.Ordinal).ToArray());
        return new PdfComplianceRequirement(
            "einvoice-xml-allowance-charge-reason",
            "EN 16931 XML allowance/charge reasons",
            PdfComplianceRequirementStatus.Missing,
            diagnostic);
    }
}
