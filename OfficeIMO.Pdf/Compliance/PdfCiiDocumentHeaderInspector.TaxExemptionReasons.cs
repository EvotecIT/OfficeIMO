namespace OfficeIMO.Pdf;

internal static partial class PdfCiiDocumentHeaderInspector {
    internal static bool TryReadTaxExemptionReasons(PdfEmbeddedFile file, out PdfCiiTaxExemptionReasonEvidence? evidence, out string? diagnostic) {
        Guard.NotNull(file, nameof(file));
        evidence = null;

        try {
            using (var stream = new MemoryStream(file.DataSnapshot))
            using (var reader = System.Xml.XmlReader.Create(stream, new System.Xml.XmlReaderSettings {
                DtdProcessing = System.Xml.DtdProcessing.Prohibit,
                XmlResolver = null
            })) {
                bool sawRoot = false;
                bool hasApplicableTradeTax = false;
                var missingReasonCategories = new List<string>();
                var forbiddenReasonCategories = new List<string>();
                var invalidNotSubjectReasonCodes = new List<string>();

                while (reader.Read()) {
                    if (reader.NodeType != System.Xml.XmlNodeType.Element) {
                        continue;
                    }

                    if (!sawRoot) {
                        sawRoot = true;
                        if (!IsCiiRoot(reader)) {
                            diagnostic = "Attach UN/CEFACT CrossIndustryInvoice XML in factur-x.xml.";
                            return false;
                        }
                    }

                    if (string.Equals(reader.LocalName, "ApplicableHeaderTradeSettlement", StringComparison.Ordinal)) {
                        ReadHeaderTaxExemptionReasons(reader, ref hasApplicableTradeTax, missingReasonCategories, forbiddenReasonCategories, invalidNotSubjectReasonCodes);
                    }
                }

                if (!sawRoot) {
                    diagnostic = "Attach non-empty UN/CEFACT CrossIndustryInvoice XML in factur-x.xml.";
                    return false;
                }

                evidence = new PdfCiiTaxExemptionReasonEvidence(
                    hasApplicableTradeTax,
                    missingReasonCategories.Count == 0,
                    missingReasonCategories.Distinct(StringComparer.Ordinal).ToArray(),
                    forbiddenReasonCategories.Count == 0,
                    forbiddenReasonCategories.Distinct(StringComparer.Ordinal).ToArray(),
                    invalidNotSubjectReasonCodes.Count == 0,
                    invalidNotSubjectReasonCodes.Distinct(StringComparer.Ordinal).ToArray());
                diagnostic = null;
                return true;
            }
        } catch (System.Xml.XmlException ex) {
            diagnostic = "Attach parseable XML in factur-x.xml: " + ex.Message;
            return false;
        }
    }

    private static void ReadHeaderTaxExemptionReasons(
        System.Xml.XmlReader reader,
        ref bool hasApplicableTradeTax,
        List<string> missingReasonCategories,
        List<string> forbiddenReasonCategories,
        List<string> invalidNotSubjectReasonCodes) {
        if (reader.IsEmptyElement) {
            return;
        }

        int depth = reader.Depth;
        while (reader.Read()) {
            if (reader.NodeType == System.Xml.XmlNodeType.Element &&
                string.Equals(reader.LocalName, "ApplicableTradeTax", StringComparison.Ordinal)) {
                hasApplicableTradeTax = true;
                ReadTaxExemptionReasonSemantics(reader, missingReasonCategories, forbiddenReasonCategories, invalidNotSubjectReasonCodes);
                continue;
            }

            if (reader.NodeType == System.Xml.XmlNodeType.EndElement &&
                reader.Depth == depth &&
                string.Equals(reader.LocalName, "ApplicableHeaderTradeSettlement", StringComparison.Ordinal)) {
                break;
            }
        }
    }

    private static void ReadTaxExemptionReasonSemantics(
        System.Xml.XmlReader reader,
        List<string> missingReasonCategories,
        List<string> forbiddenReasonCategories,
        List<string> invalidNotSubjectReasonCodes) {
        if (reader.IsEmptyElement) {
            return;
        }

        string? categoryCode = null;
        bool hasReason = false;
        string? exemptionReasonCode = null;
        int depth = reader.Depth;
        while (reader.Read()) {
            if (reader.NodeType == System.Xml.XmlNodeType.Element && reader.Depth == depth + 1) {
                if (string.Equals(reader.LocalName, "CategoryCode", StringComparison.Ordinal)) {
                    categoryCode = NormalizeTaxCategoryCode(ReadElementText(reader));
                    continue;
                }

                if (string.Equals(reader.LocalName, "ExemptionReason", StringComparison.Ordinal)) {
                    hasReason = hasReason || !string.IsNullOrWhiteSpace(ReadElementText(reader));
                    continue;
                }

                if (string.Equals(reader.LocalName, "ExemptionReasonCode", StringComparison.Ordinal)) {
                    exemptionReasonCode = ReadElementText(reader).Trim();
                    hasReason = hasReason || !string.IsNullOrWhiteSpace(exemptionReasonCode);
                    continue;
                }
            }

            if (reader.NodeType == System.Xml.XmlNodeType.EndElement &&
                reader.Depth == depth &&
                string.Equals(reader.LocalName, "ApplicableTradeTax", StringComparison.Ordinal)) {
                break;
            }
        }

        if (string.IsNullOrWhiteSpace(categoryCode)) {
            return;
        }

        if (RequiresTaxExemptionReason(categoryCode!) && !hasReason) {
            missingReasonCategories.Add(categoryCode!);
        }

        if (ForbidsTaxExemptionReason(categoryCode!) && hasReason) {
            forbiddenReasonCategories.Add(categoryCode!);
        }

        if (string.Equals(categoryCode, "O", StringComparison.Ordinal) &&
            !string.IsNullOrWhiteSpace(exemptionReasonCode) &&
            !IsNotSubjectToVatExemptionReasonCode(exemptionReasonCode!)) {
            invalidNotSubjectReasonCodes.Add(exemptionReasonCode!);
        }
    }

    private static bool IsNotSubjectToVatExemptionReasonCode(string exemptionReasonCode) =>
        string.Equals(exemptionReasonCode.Trim(), "VATEX-EU-O", StringComparison.OrdinalIgnoreCase);

    private static bool RequiresTaxExemptionReason(string categoryCode) =>
        string.Equals(categoryCode, "AE", StringComparison.Ordinal) ||
        string.Equals(categoryCode, "E", StringComparison.Ordinal) ||
        string.Equals(categoryCode, "G", StringComparison.Ordinal) ||
        string.Equals(categoryCode, "K", StringComparison.Ordinal) ||
        string.Equals(categoryCode, "O", StringComparison.Ordinal);

    private static bool ForbidsTaxExemptionReason(string categoryCode) =>
        string.Equals(categoryCode, "S", StringComparison.Ordinal) ||
        string.Equals(categoryCode, "Z", StringComparison.Ordinal) ||
        string.Equals(categoryCode, "L", StringComparison.Ordinal) ||
        string.Equals(categoryCode, "M", StringComparison.Ordinal);
}
