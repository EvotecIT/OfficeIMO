namespace OfficeIMO.Pdf;

internal static partial class PdfComplianceAnalyzer {
    private static readonly string[] ElectronicInvoiceDocumentTypeCodes = {
        "71",
        "80",
        "81",
        "82",
        "83",
        "84",
        "102",
        "130",
        "202",
        "203",
        "204",
        "211",
        "218",
        "219",
        "261",
        "262",
        "295",
        "296",
        "308",
        "325",
        "326",
        "331",
        "380",
        "381",
        "382",
        "383",
        "384",
        "385",
        "386",
        "387",
        "388",
        "389",
        "390",
        "393",
        "394",
        "395",
        "396",
        "420",
        "456",
        "457",
        "458",
        "527",
        "532",
        "553",
        "575",
        "623",
        "633",
        "751",
        "780",
        "817",
        "870",
        "875",
        "876",
        "877",
        "935"
    };

    private static PdfComplianceRequirement BuildElectronicInvoiceXmlDocumentTypeCodeRequirement(PdfOptions options) {
        IReadOnlyList<PdfEmbeddedFile> embeddedFiles = options.EmbeddedFiles;
        var diagnostics = new List<string>();
        for (int i = 0; i < embeddedFiles.Count; i++) {
            PdfEmbeddedFile file = embeddedFiles[i];
            if (!IsFacturXCiiAttachment(file, diagnostics)) {
                continue;
            }

            if (!PdfCiiDocumentHeaderInspector.TryRead(file, out PdfCiiDocumentHeaderEvidence? evidence, out string? documentDiagnostic)) {
                return new PdfComplianceRequirement(
                    "einvoice-xml-document-type-code",
                    "EN 16931 XML document type code",
                    PdfComplianceRequirementStatus.Missing,
                    documentDiagnostic!);
            }

            string typeCode = evidence!.TypeCode ?? string.Empty;
            if (string.IsNullOrWhiteSpace(typeCode)) {
                return new PdfComplianceRequirement(
                    "einvoice-xml-document-type-code",
                    "EN 16931 XML document type code",
                    PdfComplianceRequirementStatus.Missing,
                    "Set factur-x.xml ExchangedDocument TypeCode before Mustang validation.");
            }

            if (!IsKnownElectronicInvoiceDocumentTypeCode(typeCode)) {
                return new PdfComplianceRequirement(
                    "einvoice-xml-document-type-code",
                    "EN 16931 XML document type code",
                    PdfComplianceRequirementStatus.Missing,
                    "Set factur-x.xml ExchangedDocument TypeCode to an invoice or credit-note related UNTDID 1001 code before Mustang validation. Found: " + typeCode + ".");
            }

            return new PdfComplianceRequirement(
                "einvoice-xml-document-type-code",
                "EN 16931 XML document type code",
                PdfComplianceRequirementStatus.Satisfied,
                "The factur-x.xml CrossIndustryInvoice document type code is from the invoice and credit-note related UNTDID 1001 code lists for e-invoice readiness.");
        }

        string diagnostic = diagnostics.Count == 0
            ? "Attach a canonical factur-x.xml CrossIndustryInvoice file before checking EN 16931 XML document type code."
            : string.Join(" ", diagnostics.Distinct(StringComparer.Ordinal).ToArray());
        return new PdfComplianceRequirement(
            "einvoice-xml-document-type-code",
            "EN 16931 XML document type code",
            PdfComplianceRequirementStatus.Missing,
            diagnostic);
    }

    private static bool IsKnownElectronicInvoiceDocumentTypeCode(string typeCode) {
        string normalized = typeCode.Trim();
        for (int i = 0; i < ElectronicInvoiceDocumentTypeCodes.Length; i++) {
            if (string.Equals(normalized, ElectronicInvoiceDocumentTypeCodes[i], StringComparison.Ordinal)) {
                return true;
            }
        }

        return false;
    }
}
