namespace OfficeIMO.Pdf;

internal static partial class PdfComplianceAnalyzer {
    private static readonly string[] ElectronicInvoiceElectronicAddressSchemeIds = {
        "0002",
        "0007",
        "0009",
        "0037",
        "0060",
        "0088",
        "0096",
        "0097",
        "0106",
        "0130",
        "0135",
        "0142",
        "0147",
        "0151",
        "0154",
        "0158",
        "0170",
        "0177",
        "0183",
        "0184",
        "0188",
        "0190",
        "0191",
        "0192",
        "0193",
        "0194",
        "0195",
        "0196",
        "0198",
        "0199",
        "0200",
        "0201",
        "0202",
        "0203",
        "0204",
        "0205",
        "0208",
        "0209",
        "0210",
        "0211",
        "0212",
        "0213",
        "0215",
        "0216",
        "0217",
        "0218",
        "0221",
        "0225",
        "0230",
        "0235",
        "0240",
        "0244",
        "0245",
        "9910",
        "9913",
        "9914",
        "9915",
        "9918",
        "9919",
        "9920",
        "9922",
        "9923",
        "9924",
        "9925",
        "9926",
        "9927",
        "9928",
        "9929",
        "9930",
        "9931",
        "9932",
        "9933",
        "9934",
        "9935",
        "9936",
        "9937",
        "9938",
        "9939",
        "9940",
        "9941",
        "9942",
        "9943",
        "9944",
        "9945",
        "9946",
        "9947",
        "9948",
        "9949",
        "9950",
        "9951",
        "9952",
        "9953",
        "9957",
        "9959"
    };

    private static PdfComplianceRequirement BuildElectronicInvoiceXmlElectronicAddressRequirement(PdfOptions options) {
        IReadOnlyList<PdfEmbeddedFile> embeddedFiles = options.EmbeddedFiles;
        var diagnostics = new List<string>();
        for (int i = 0; i < embeddedFiles.Count; i++) {
            PdfEmbeddedFile file = embeddedFiles[i];
            if (!IsFacturXCiiAttachment(file, diagnostics)) {
                continue;
            }

            if (!PdfCiiDocumentHeaderInspector.TryReadElectronicAddresses(file, out PdfCiiElectronicAddressEvidence? evidence, out string? addressDiagnostic)) {
                return new PdfComplianceRequirement(
                    "einvoice-xml-electronic-address",
                    "EN 16931 XML electronic address",
                    PdfComplianceRequirementStatus.Missing,
                    addressDiagnostic!);
            }

            var missingFields = new List<string>();
            if (!evidence!.HasSellerUriUniversalCommunication) {
                missingFields.Add("SellerTradeParty URIUniversalCommunication");
            }

            if (!evidence.HasSellerUriId) {
                missingFields.Add("SellerTradeParty URIUniversalCommunication URIID");
            }

            if (!evidence.HasSellerSchemeId) {
                missingFields.Add("SellerTradeParty URIUniversalCommunication URIID schemeID");
            }

            if (!evidence.HasBuyerUriUniversalCommunication) {
                missingFields.Add("BuyerTradeParty URIUniversalCommunication");
            }

            if (!evidence.HasBuyerUriId) {
                missingFields.Add("BuyerTradeParty URIUniversalCommunication URIID");
            }

            if (!evidence.HasBuyerSchemeId) {
                missingFields.Add("BuyerTradeParty URIUniversalCommunication URIID schemeID");
            }

            if (missingFields.Count > 0) {
                return new PdfComplianceRequirement(
                    "einvoice-xml-electronic-address",
                    "EN 16931 XML electronic address",
                    PdfComplianceRequirementStatus.Missing,
                    "Set factur-x.xml seller and buyer electronic address essentials before Mustang validation: " + string.Join(", ", missingFields.ToArray()) + ".");
            }

            var invalidSchemeIds = new List<string>();
            AddInvalidElectronicAddressSchemeIds(evidence.SellerSchemeIds, invalidSchemeIds);
            AddInvalidElectronicAddressSchemeIds(evidence.BuyerSchemeIds, invalidSchemeIds);

            if (invalidSchemeIds.Count > 0) {
                return new PdfComplianceRequirement(
                    "einvoice-xml-electronic-address",
                    "EN 16931 XML electronic address",
                    PdfComplianceRequirementStatus.Missing,
                    "Set factur-x.xml SellerTradeParty and BuyerTradeParty URIUniversalCommunication URIID schemeID to a CEF Electronic Address Scheme (EAS) code before Mustang validation. Found: " + string.Join(", ", invalidSchemeIds.Distinct(StringComparer.Ordinal).ToArray()) + ".");
            }

            return new PdfComplianceRequirement(
                "einvoice-xml-electronic-address",
                "EN 16931 XML electronic address",
                PdfComplianceRequirementStatus.Satisfied,
                "The factur-x.xml CrossIndustryInvoice seller and buyer electronic address scheme identifiers are from the CEF Electronic Address Scheme (EAS) code list for e-invoice routing readiness.");
        }

        string diagnostic = diagnostics.Count == 0
            ? "Attach a canonical factur-x.xml CrossIndustryInvoice file before checking EN 16931 XML electronic addresses."
            : string.Join(" ", diagnostics.Distinct(StringComparer.Ordinal).ToArray());
        return new PdfComplianceRequirement(
            "einvoice-xml-electronic-address",
            "EN 16931 XML electronic address",
            PdfComplianceRequirementStatus.Missing,
            diagnostic);
    }

    private static void AddInvalidElectronicAddressSchemeIds(IReadOnlyList<string> schemeIds, List<string> invalidSchemeIds) {
        for (int i = 0; i < schemeIds.Count; i++) {
            if (!IsKnownElectronicInvoiceElectronicAddressSchemeId(schemeIds[i])) {
                invalidSchemeIds.Add(schemeIds[i]);
            }
        }
    }

    private static bool IsKnownElectronicInvoiceElectronicAddressSchemeId(string schemeId) {
        string normalized = schemeId.Trim();
        for (int i = 0; i < ElectronicInvoiceElectronicAddressSchemeIds.Length; i++) {
            if (string.Equals(normalized, ElectronicInvoiceElectronicAddressSchemeIds[i], StringComparison.Ordinal)) {
                return true;
            }
        }

        return false;
    }
}
