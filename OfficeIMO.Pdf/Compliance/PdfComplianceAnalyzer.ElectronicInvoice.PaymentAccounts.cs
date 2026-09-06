namespace OfficeIMO.Pdf;

internal static partial class PdfComplianceAnalyzer {
    private static PdfComplianceRequirement BuildElectronicInvoiceXmlPaymentAccountFormatRequirement(PdfOptions options) {
        IReadOnlyList<PdfEmbeddedFile> embeddedFiles = options.EmbeddedFiles;
        var diagnostics = new List<string>();
        for (int i = 0; i < embeddedFiles.Count; i++) {
            PdfEmbeddedFile file = embeddedFiles[i];
            if (!IsFacturXCiiAttachment(file, diagnostics)) {
                continue;
            }

            if (!PdfCiiDocumentHeaderInspector.TryReadPaymentAccounts(file, out PdfCiiPaymentAccountEvidence? evidence, out string? paymentDiagnostic)) {
                return new PdfComplianceRequirement(
                    "einvoice-xml-payment-account-format",
                    "EN 16931 XML payment account format",
                    PdfComplianceRequirementStatus.Missing,
                    paymentDiagnostic!);
            }

            var missingFields = new List<string>();
            if (!evidence!.HasSpecifiedTradeSettlementPaymentMeans) {
                missingFields.Add("SpecifiedTradeSettlementPaymentMeans");
            }

            bool requiresCreditorAccount = RequiresElectronicInvoiceCreditorAccount(evidence.TypeCodes);
            bool hasCreditorAccountData = evidence.HasCreditorFinancialAccount || evidence.HasAccountId;
            if ((requiresCreditorAccount || hasCreditorAccountData) && !evidence.HasCreditorFinancialAccount) {
                missingFields.Add("PayeePartyCreditorFinancialAccount");
            }

            if ((requiresCreditorAccount || hasCreditorAccountData) && !evidence.HasAccountId) {
                missingFields.Add("PayeePartyCreditorFinancialAccount IBANID or ProprietaryID");
            }

            if (missingFields.Count > 0) {
                return new PdfComplianceRequirement(
                    "einvoice-xml-payment-account-format",
                    "EN 16931 XML payment account format",
                    PdfComplianceRequirementStatus.Missing,
                    "Set factur-x.xml creditor account identifiers before Mustang validation: " + string.Join(", ", missingFields.ToArray()) + ".");
            }

            if (!evidence.AllIbanIdsAreValid) {
                return new PdfComplianceRequirement(
                    "einvoice-xml-payment-account-format",
                    "EN 16931 XML payment account format",
                    PdfComplianceRequirementStatus.Missing,
                    "Set factur-x.xml PayeePartyCreditorFinancialAccount IBANID to a valid IBAN checksum value before Mustang validation: " + string.Join(", ", evidence.InvalidIbanIds.ToArray()) + ".");
            }

            return new PdfComplianceRequirement(
                "einvoice-xml-payment-account-format",
                "EN 16931 XML payment account format",
                PdfComplianceRequirementStatus.Satisfied,
                requiresCreditorAccount
                    ? "The factur-x.xml CrossIndustryInvoice creditor account identifiers are present, and supplied IBAN values pass checksum validation for e-invoice readiness."
                    : "The factur-x.xml CrossIndustryInvoice payment means type code does not require creditor account identifiers, and supplied IBAN values pass checksum validation for e-invoice readiness.");
        }

        string diagnostic = diagnostics.Count == 0
            ? "Attach a canonical factur-x.xml CrossIndustryInvoice file before checking EN 16931 XML payment account formats."
            : string.Join(" ", diagnostics.Distinct(StringComparer.Ordinal).ToArray());
        return new PdfComplianceRequirement(
            "einvoice-xml-payment-account-format",
            "EN 16931 XML payment account format",
            PdfComplianceRequirementStatus.Missing,
            diagnostic);
    }
}
