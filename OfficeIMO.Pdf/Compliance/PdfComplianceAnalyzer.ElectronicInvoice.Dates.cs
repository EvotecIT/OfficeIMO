namespace OfficeIMO.Pdf;

internal static partial class PdfComplianceAnalyzer {
    private static PdfComplianceRequirement BuildElectronicInvoiceXmlDateFormatRequirement(PdfOptions options) {
        IReadOnlyList<PdfEmbeddedFile> embeddedFiles = options.EmbeddedFiles;
        var diagnostics = new List<string>();
        for (int i = 0; i < embeddedFiles.Count; i++) {
            PdfEmbeddedFile file = embeddedFiles[i];
            if (!IsFacturXCiiAttachment(file, diagnostics)) {
                continue;
            }

            if (!PdfCiiDocumentHeaderInspector.TryReadDateFormats(file, out PdfCiiDateFormatEvidence? evidence, out string? dateDiagnostic)) {
                return new PdfComplianceRequirement(
                    "einvoice-xml-date-format",
                    "EN 16931 XML date format",
                    PdfComplianceRequirementStatus.Missing,
                    dateDiagnostic!);
            }

            var missingFields = new List<string>();
            if (!evidence!.HasIssueDateTime) {
                missingFields.Add("ExchangedDocument IssueDateTime");
            }

            if (missingFields.Count > 0) {
                return new PdfComplianceRequirement(
                    "einvoice-xml-date-format",
                    "EN 16931 XML date format",
                    PdfComplianceRequirementStatus.Missing,
                    "Set factur-x.xml parseable CII date markers before Mustang validation: " + string.Join(", ", missingFields.ToArray()) + ".");
            }

            if (!evidence.AllKnownDatesAreParseable) {
                return new PdfComplianceRequirement(
                    "einvoice-xml-date-format",
                    "EN 16931 XML date format",
                    PdfComplianceRequirementStatus.Missing,
                    "Set factur-x.xml CII dates to parseable DateTimeString values matching their format attribute: " + string.Join(", ", evidence.InvalidDateFields.ToArray()) + ".");
            }

            return new PdfComplianceRequirement(
                "einvoice-xml-date-format",
                "EN 16931 XML date format",
                PdfComplianceRequirementStatus.Satisfied,
                "The factur-x.xml CrossIndustryInvoice issue and payment due date markers use parseable CII DateTimeString values for e-invoice readiness.");
        }

        string diagnostic = diagnostics.Count == 0
            ? "Attach a canonical factur-x.xml CrossIndustryInvoice file before checking EN 16931 XML date formats."
            : string.Join(" ", diagnostics.Distinct(StringComparer.Ordinal).ToArray());
        return new PdfComplianceRequirement(
            "einvoice-xml-date-format",
            "EN 16931 XML date format",
            PdfComplianceRequirementStatus.Missing,
            diagnostic);
    }
}
