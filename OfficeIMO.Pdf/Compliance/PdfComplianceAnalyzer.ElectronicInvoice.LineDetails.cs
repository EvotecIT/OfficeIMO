namespace OfficeIMO.Pdf;

internal static partial class PdfComplianceAnalyzer {
    private static PdfComplianceRequirement BuildElectronicInvoiceXmlLineTaxRequirement(PdfOptions options) {
        IReadOnlyList<PdfEmbeddedFile> embeddedFiles = options.EmbeddedFiles;
        var diagnostics = new List<string>();
        for (int i = 0; i < embeddedFiles.Count; i++) {
            PdfEmbeddedFile file = embeddedFiles[i];
            if (!IsFacturXCiiAttachment(file, diagnostics)) {
                continue;
            }

            if (!PdfCiiDocumentHeaderInspector.TryReadLineTax(file, out PdfCiiLineTaxEvidence? evidence, out string? taxDiagnostic)) {
                return new PdfComplianceRequirement(
                    "einvoice-xml-line-tax",
                    "EN 16931 XML line tax",
                    PdfComplianceRequirementStatus.Missing,
                    taxDiagnostic!);
            }

            var missingFields = new List<string>();
            if (!evidence!.HasIncludedSupplyChainTradeLineItem) {
                missingFields.Add("IncludedSupplyChainTradeLineItem");
            }

            if (!evidence.HasSpecifiedLineTradeSettlement) {
                missingFields.Add("SpecifiedLineTradeSettlement");
            }

            if (!evidence.HasApplicableTradeTax) {
                missingFields.Add("ApplicableTradeTax");
            }

            if (!evidence.HasTypeCode) {
                missingFields.Add("ApplicableTradeTax TypeCode");
            }

            if (!evidence.HasCategoryCode) {
                missingFields.Add("ApplicableTradeTax CategoryCode");
            }

            if (!evidence.HasRateRequirementCoverage && evidence.MissingRateCategoryCodes.Count == 0 && evidence.ForbiddenRateCategoryCodes.Count == 0) {
                missingFields.Add("ApplicableTradeTax RateApplicablePercent");
            }

            if (evidence.MissingLineTaxFields.Count > 0) {
                missingFields.AddRange(evidence.MissingLineTaxFields);
            }

            if (missingFields.Count > 0) {
                return new PdfComplianceRequirement(
                    "einvoice-xml-line-tax",
                    "EN 16931 XML line tax",
                    PdfComplianceRequirementStatus.Missing,
                    "Set factur-x.xml line trade tax essentials before Mustang validation: " + string.Join(", ", missingFields.ToArray()) + ".");
            }

            if (evidence.MissingRateCategoryCodes.Count > 0) {
                return new PdfComplianceRequirement(
                    "einvoice-xml-line-tax",
                    "EN 16931 XML line tax",
                    PdfComplianceRequirementStatus.Missing,
                    "Set factur-x.xml line ApplicableTradeTax RateApplicablePercent before Mustang validation for categories that require an invoiced item VAT rate. Missing line tax rate categories: " + string.Join(", ", evidence.MissingRateCategoryCodes.ToArray()) + ".");
            }

            if (evidence.ForbiddenRateCategoryCodes.Count > 0) {
                return new PdfComplianceRequirement(
                    "einvoice-xml-line-tax",
                    "EN 16931 XML line tax",
                    PdfComplianceRequirementStatus.Missing,
                    "Remove factur-x.xml line ApplicableTradeTax RateApplicablePercent for Peppol category O before Mustang validation. Forbidden line tax rate categories: " + string.Join(", ", evidence.ForbiddenRateCategoryCodes.ToArray()) + ".");
            }

            var invalidTypeCodes = new List<string>();
            for (int j = 0; j < evidence.TypeCodes.Count; j++) {
                string typeCode = evidence.TypeCodes[j];
                if (!string.Equals(typeCode.Trim(), "VAT", StringComparison.Ordinal)) {
                    invalidTypeCodes.Add(typeCode);
                }
            }

            if (invalidTypeCodes.Count > 0) {
                return new PdfComplianceRequirement(
                    "einvoice-xml-line-tax",
                    "EN 16931 XML line tax",
                    PdfComplianceRequirementStatus.Missing,
                    "Set factur-x.xml line ApplicableTradeTax TypeCode to VAT before Mustang validation. Found: " + string.Join(", ", invalidTypeCodes.Distinct(StringComparer.Ordinal).ToArray()) + ".");
            }

            return new PdfComplianceRequirement(
                "einvoice-xml-line-tax",
                "EN 16931 XML line tax",
                PdfComplianceRequirementStatus.Satisfied,
                "The factur-x.xml CrossIndustryInvoice includes VAT line trade settlement tax type/category markers and category-aware rate presence for e-invoice readiness.");
        }

        string diagnostic = diagnostics.Count == 0
            ? "Attach a canonical factur-x.xml CrossIndustryInvoice file before checking EN 16931 XML line tax essentials."
            : string.Join(" ", diagnostics.Distinct(StringComparer.Ordinal).ToArray());
        return new PdfComplianceRequirement(
            "einvoice-xml-line-tax",
            "EN 16931 XML line tax",
            PdfComplianceRequirementStatus.Missing,
            diagnostic);
    }
}
