namespace OfficeIMO.Pdf;

internal static partial class PdfComplianceAnalyzer {
    private static readonly string[] ElectronicInvoicePaymentMeansCodes = {
        "1",
        "2",
        "3",
        "4",
        "5",
        "6",
        "7",
        "8",
        "9",
        "10",
        "11",
        "12",
        "13",
        "14",
        "15",
        "16",
        "17",
        "18",
        "19",
        "20",
        "21",
        "22",
        "23",
        "24",
        "25",
        "26",
        "27",
        "28",
        "29",
        "30",
        "31",
        "32",
        "33",
        "34",
        "35",
        "36",
        "37",
        "38",
        "39",
        "40",
        "41",
        "42",
        "43",
        "44",
        "45",
        "46",
        "47",
        "48",
        "49",
        "50",
        "51",
        "52",
        "53",
        "54",
        "55",
        "56",
        "57",
        "58",
        "59",
        "60",
        "61",
        "62",
        "63",
        "64",
        "65",
        "66",
        "67",
        "68",
        "69",
        "70",
        "74",
        "75",
        "76",
        "77",
        "78",
        "91",
        "92",
        "93",
        "94",
        "95",
        "96",
        "97",
        "98",
        "ZZZ"
    };

    private static PdfComplianceRequirement BuildElectronicInvoiceXmlPaymentMeansCodeRequirement(PdfOptions options) {
        IReadOnlyList<PdfEmbeddedFile> embeddedFiles = options.EmbeddedFiles;
        var diagnostics = new List<string>();
        for (int i = 0; i < embeddedFiles.Count; i++) {
            PdfEmbeddedFile file = embeddedFiles[i];
            if (!IsFacturXCiiAttachment(file, diagnostics)) {
                continue;
            }

            if (!PdfCiiDocumentHeaderInspector.TryReadPaymentMeansCodes(file, out PdfCiiPaymentMeansCodeEvidence? evidence, out string? paymentDiagnostic)) {
                return new PdfComplianceRequirement(
                    "einvoice-xml-payment-means-code",
                    "EN 16931 XML payment means code",
                    PdfComplianceRequirementStatus.Missing,
                    paymentDiagnostic!);
            }

            var missingFields = new List<string>();
            if (!evidence!.HasSpecifiedTradeSettlementPaymentMeans) {
                missingFields.Add("SpecifiedTradeSettlementPaymentMeans");
            }

            if (!evidence.HasTypeCode) {
                missingFields.Add("SpecifiedTradeSettlementPaymentMeans TypeCode");
            }

            if (evidence.MissingTypeCodePaymentMeans.Count > 0) {
                missingFields.Add("SpecifiedTradeSettlementPaymentMeans TypeCode on " + string.Join(", ", evidence.MissingTypeCodePaymentMeans.ToArray()));
            }

            if (missingFields.Count > 0) {
                return new PdfComplianceRequirement(
                    "einvoice-xml-payment-means-code",
                    "EN 16931 XML payment means code",
                    PdfComplianceRequirementStatus.Missing,
                    "Set factur-x.xml payment means type code before Mustang validation: " + string.Join(", ", missingFields.ToArray()) + ".");
            }

            var invalidTypeCodes = new List<string>();
            for (int j = 0; j < evidence.TypeCodes.Count; j++) {
                if (!IsKnownElectronicInvoicePaymentMeansCode(evidence.TypeCodes[j])) {
                    invalidTypeCodes.Add(evidence.TypeCodes[j]);
                }
            }

            if (invalidTypeCodes.Count > 0) {
                return new PdfComplianceRequirement(
                    "einvoice-xml-payment-means-code",
                    "EN 16931 XML payment means code",
                    PdfComplianceRequirementStatus.Missing,
                    "Set factur-x.xml SpecifiedTradeSettlementPaymentMeans TypeCode to a UNCL4461 payment means code before Mustang validation. Found: " + string.Join(", ", invalidTypeCodes.Distinct(StringComparer.Ordinal).ToArray()) + ".");
            }

            return new PdfComplianceRequirement(
                "einvoice-xml-payment-means-code",
                "EN 16931 XML payment means code",
                PdfComplianceRequirementStatus.Satisfied,
                "The factur-x.xml CrossIndustryInvoice payment means type code is from the UNCL4461 payment means code list for e-invoice readiness.");
        }

        string diagnostic = diagnostics.Count == 0
            ? "Attach a canonical factur-x.xml CrossIndustryInvoice file before checking EN 16931 XML payment means code."
            : string.Join(" ", diagnostics.Distinct(StringComparer.Ordinal).ToArray());
        return new PdfComplianceRequirement(
            "einvoice-xml-payment-means-code",
            "EN 16931 XML payment means code",
            PdfComplianceRequirementStatus.Missing,
            diagnostic);
    }

    private static bool IsKnownElectronicInvoicePaymentMeansCode(string typeCode) {
        string normalized = typeCode.Trim();
        for (int i = 0; i < ElectronicInvoicePaymentMeansCodes.Length; i++) {
            if (string.Equals(normalized, ElectronicInvoicePaymentMeansCodes[i], StringComparison.Ordinal)) {
                return true;
            }
        }

        return false;
    }

    private static bool RequiresElectronicInvoiceCreditorAccount(IReadOnlyList<string> typeCodes) {
        if (typeCodes.Count == 0) {
            return true;
        }

        for (int i = 0; i < typeCodes.Count; i++) {
            string normalized = typeCodes[i].Trim();
            if (string.Equals(normalized, "30", StringComparison.Ordinal) ||
                string.Equals(normalized, "31", StringComparison.Ordinal) ||
                string.Equals(normalized, "42", StringComparison.Ordinal) ||
                string.Equals(normalized, "45", StringComparison.Ordinal) ||
                string.Equals(normalized, "58", StringComparison.Ordinal) ||
                string.Equals(normalized, "59", StringComparison.Ordinal)) {
                return true;
            }
        }

        return false;
    }
}
