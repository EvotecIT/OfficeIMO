namespace OfficeIMO.Pdf;

internal static partial class PdfComplianceAnalyzer {
    private static PdfComplianceRequirement BuildElectronicInvoiceXmlUnitCodeRequirement(PdfOptions options) {
        IReadOnlyList<PdfEmbeddedFile> embeddedFiles = options.EmbeddedFiles;
        var diagnostics = new List<string>();
        for (int i = 0; i < embeddedFiles.Count; i++) {
            PdfEmbeddedFile file = embeddedFiles[i];
            if (!IsFacturXCiiAttachment(file, diagnostics)) {
                continue;
            }

            if (!PdfCiiDocumentHeaderInspector.TryReadUnitCodes(file, out PdfCiiUnitCodeEvidence? evidence, out string? unitDiagnostic)) {
                return new PdfComplianceRequirement(
                    "einvoice-xml-unit-code",
                    "EN 16931 XML unit code",
                    PdfComplianceRequirementStatus.Missing,
                    unitDiagnostic!);
            }

            var missingFields = new List<string>();
            if (!evidence!.HasIncludedSupplyChainTradeLineItem) {
                missingFields.Add("IncludedSupplyChainTradeLineItem");
            }

            if (!evidence.HasBilledQuantity) {
                missingFields.Add("SpecifiedLineTradeDelivery BilledQuantity");
            }

            if (!evidence.HasBilledQuantityUnitCode) {
                missingFields.Add("SpecifiedLineTradeDelivery BilledQuantity unitCode");
            }

            if (missingFields.Count > 0) {
                return new PdfComplianceRequirement(
                    "einvoice-xml-unit-code",
                    "EN 16931 XML unit code",
                    PdfComplianceRequirementStatus.Missing,
                    "Set factur-x.xml line quantity unit-code essentials before Mustang validation: " + string.Join(", ", missingFields.ToArray()) + ".");
            }

            var invalidUnitCodes = new List<string>();
            for (int j = 0; j < evidence.UnitCodes.Count; j++) {
                if (!IsKnownElectronicInvoiceUnitCode(evidence.UnitCodes[j])) {
                    invalidUnitCodes.Add(evidence.UnitCodes[j].Trim().ToUpperInvariant());
                }
            }

            if (invalidUnitCodes.Count > 0) {
                return new PdfComplianceRequirement(
                    "einvoice-xml-unit-code",
                    "EN 16931 XML unit code",
                    PdfComplianceRequirementStatus.Missing,
                    "Set factur-x.xml SpecifiedLineTradeDelivery BilledQuantity unitCode values to UN/ECE Recommendation 20 codes with Rec 21 X-prefixed packaging codes before Mustang validation. Found: " + string.Join(", ", invalidUnitCodes.Distinct(StringComparer.Ordinal).ToArray()) + ".");
            }

            return new PdfComplianceRequirement(
                "einvoice-xml-unit-code",
                "EN 16931 XML unit code",
                PdfComplianceRequirementStatus.Satisfied,
                "The factur-x.xml CrossIndustryInvoice billed quantity unit codes are from the UN/ECE Recommendation 20 code list with Rec 21 X-prefixed packaging codes for e-invoice readiness.");
        }

        string diagnostic = diagnostics.Count == 0
            ? "Attach a canonical factur-x.xml CrossIndustryInvoice file before checking EN 16931 XML unit codes."
            : string.Join(" ", diagnostics.Distinct(StringComparer.Ordinal).ToArray());
        return new PdfComplianceRequirement(
            "einvoice-xml-unit-code",
            "EN 16931 XML unit code",
            PdfComplianceRequirementStatus.Missing,
            diagnostic);
    }

    private static bool IsKnownElectronicInvoiceUnitCode(string unitCode) {
        string normalized = " " + unitCode.Trim().ToUpperInvariant() + " ";
        for (int i = 0; i < ElectronicInvoiceUnitCodeChunks.Length; i++) {
            if (ElectronicInvoiceUnitCodeChunks[i].Contains(normalized)) {
                return true;
            }
        }

        return false;
    }
}
