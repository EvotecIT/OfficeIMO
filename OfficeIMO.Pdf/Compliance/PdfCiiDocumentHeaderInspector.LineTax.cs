namespace OfficeIMO.Pdf;

internal static partial class PdfCiiDocumentHeaderInspector {
    internal static bool TryReadLineTax(PdfEmbeddedFile file, out PdfCiiLineTaxEvidence? evidence, out string? diagnostic) {
        Guard.NotNull(file, nameof(file));
        evidence = null;

        try {
            using (var stream = new MemoryStream(file.DataSnapshot))
            using (var reader = System.Xml.XmlReader.Create(stream, new System.Xml.XmlReaderSettings {
                DtdProcessing = System.Xml.DtdProcessing.Prohibit,
                XmlResolver = null
            })) {
                bool sawRoot = false;
                bool hasLineItem = false;
                bool hasSettlement = false;
                bool hasTradeTax = false;
                bool hasTypeCode = false;
                bool hasCategoryCode = false;
                bool hasRateApplicablePercent = false;
                bool hasRateRequirementCoverage = false;
                var typeCodes = new List<string>();
                var missingLineTaxFields = new List<string>();
                var missingRateCategoryCodes = new List<string>();
                var forbiddenRateCategoryCodes = new List<string>();
                int lineItemNumber = 0;

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

                    if (string.Equals(reader.LocalName, "IncludedSupplyChainTradeLineItem", StringComparison.Ordinal)) {
                        hasLineItem = true;
                        lineItemNumber++;
                        bool lineHasSettlement = false;
                        bool lineHasTradeTax = false;
                        bool lineHasTypeCode = false;
                        bool lineHasCategoryCode = false;
                        bool lineHasRateApplicablePercent = false;
                        bool lineHasRateRequirementCoverage = false;
                        ReadLineTax(reader, ref lineHasSettlement, ref lineHasTradeTax, ref lineHasTypeCode, ref lineHasCategoryCode, ref lineHasRateApplicablePercent, ref lineHasRateRequirementCoverage, typeCodes, missingRateCategoryCodes, forbiddenRateCategoryCodes);
                        hasSettlement = hasSettlement || lineHasSettlement;
                        hasTradeTax = hasTradeTax || lineHasTradeTax;
                        hasTypeCode = hasTypeCode || lineHasTypeCode;
                        hasCategoryCode = hasCategoryCode || lineHasCategoryCode;
                        hasRateApplicablePercent = hasRateApplicablePercent || lineHasRateApplicablePercent;
                        hasRateRequirementCoverage = hasRateRequirementCoverage || lineHasRateRequirementCoverage;
                        AddMissingLineTaxFields(lineItemNumber, lineHasSettlement, lineHasTradeTax, lineHasTypeCode, lineHasCategoryCode, missingLineTaxFields);
                    }
                }

                if (!sawRoot) {
                    diagnostic = "Attach non-empty UN/CEFACT CrossIndustryInvoice XML in factur-x.xml.";
                    return false;
                }

                evidence = new PdfCiiLineTaxEvidence(
                    hasLineItem,
                    hasSettlement,
                    hasTradeTax,
                    hasTypeCode,
                    hasCategoryCode,
                    hasRateApplicablePercent,
                    hasRateRequirementCoverage,
                    typeCodes,
                    missingLineTaxFields,
                    missingRateCategoryCodes.Distinct(StringComparer.Ordinal).ToArray(),
                    forbiddenRateCategoryCodes.Distinct(StringComparer.Ordinal).ToArray());
                diagnostic = null;
                return true;
            }
        } catch (System.Xml.XmlException ex) {
            diagnostic = "Attach parseable XML in factur-x.xml: " + ex.Message;
            return false;
        }
    }

    private static void ReadLineTax(
        System.Xml.XmlReader reader,
        ref bool hasSettlement,
        ref bool hasTradeTax,
        ref bool hasTypeCode,
        ref bool hasCategoryCode,
        ref bool hasRateApplicablePercent,
        ref bool hasRateRequirementCoverage,
        List<string> typeCodes,
        List<string> missingRateCategoryCodes,
        List<string> forbiddenRateCategoryCodes) {
        if (reader.IsEmptyElement) {
            return;
        }

        int depth = reader.Depth;
        while (reader.Read()) {
            if (reader.NodeType == System.Xml.XmlNodeType.Element) {
                if (string.Equals(reader.LocalName, "SpecifiedLineTradeSettlement", StringComparison.Ordinal)) {
                    hasSettlement = true;
                    ReadLineTradeSettlementTax(reader, ref hasTradeTax, ref hasTypeCode, ref hasCategoryCode, ref hasRateApplicablePercent, ref hasRateRequirementCoverage, typeCodes, missingRateCategoryCodes, forbiddenRateCategoryCodes);
                    continue;
                }
            }

            if (reader.NodeType == System.Xml.XmlNodeType.EndElement &&
                reader.Depth == depth &&
                string.Equals(reader.LocalName, "IncludedSupplyChainTradeLineItem", StringComparison.Ordinal)) {
                break;
            }
        }
    }

    private static void ReadLineTradeSettlementTax(
        System.Xml.XmlReader reader,
        ref bool hasTradeTax,
        ref bool hasTypeCode,
        ref bool hasCategoryCode,
        ref bool hasRateApplicablePercent,
        ref bool hasRateRequirementCoverage,
        List<string> typeCodes,
        List<string> missingRateCategoryCodes,
        List<string> forbiddenRateCategoryCodes) {
        if (reader.IsEmptyElement) {
            return;
        }

        int depth = reader.Depth;
        while (reader.Read()) {
            if (reader.NodeType == System.Xml.XmlNodeType.Element) {
                if (string.Equals(reader.LocalName, "ApplicableTradeTax", StringComparison.Ordinal)) {
                    hasTradeTax = true;
                    ReadLineApplicableTradeTax(reader, ref hasTypeCode, ref hasCategoryCode, ref hasRateApplicablePercent, ref hasRateRequirementCoverage, typeCodes, missingRateCategoryCodes, forbiddenRateCategoryCodes);
                    continue;
                }
            }

            if (reader.NodeType == System.Xml.XmlNodeType.EndElement &&
                reader.Depth == depth &&
                string.Equals(reader.LocalName, "SpecifiedLineTradeSettlement", StringComparison.Ordinal)) {
                break;
            }
        }
    }

    private static void ReadLineApplicableTradeTax(
        System.Xml.XmlReader reader,
        ref bool hasTypeCode,
        ref bool hasCategoryCode,
        ref bool hasRateApplicablePercent,
        ref bool hasRateRequirementCoverage,
        List<string> typeCodes,
        List<string> missingRateCategoryCodes,
        List<string> forbiddenRateCategoryCodes) {
        if (reader.IsEmptyElement) {
            return;
        }

        string? categoryCode = null;
        string? rawRate = null;
        int depth = reader.Depth;
        while (reader.Read()) {
            if (reader.NodeType == System.Xml.XmlNodeType.Element && reader.Depth == depth + 1) {
                if (string.Equals(reader.LocalName, "TypeCode", StringComparison.Ordinal)) {
                    string typeCode = ReadElementText(reader);
                    if (!string.IsNullOrWhiteSpace(typeCode)) {
                        hasTypeCode = true;
                        typeCodes.Add(typeCode.Trim());
                    }

                    continue;
                }

                if (string.Equals(reader.LocalName, "CategoryCode", StringComparison.Ordinal)) {
                    categoryCode = NormalizeTaxCategoryCode(ReadElementText(reader));
                    hasCategoryCode = hasCategoryCode || !string.IsNullOrWhiteSpace(categoryCode);
                    continue;
                }

                if (string.Equals(reader.LocalName, "RateApplicablePercent", StringComparison.Ordinal)) {
                    rawRate = ReadElementText(reader).Trim();
                    hasRateApplicablePercent = hasRateApplicablePercent || !string.IsNullOrWhiteSpace(rawRate);
                    continue;
                }
            }

            if (reader.NodeType == System.Xml.XmlNodeType.EndElement &&
                reader.Depth == depth &&
                string.Equals(reader.LocalName, "ApplicableTradeTax", StringComparison.Ordinal)) {
                break;
            }
        }

        string normalizedCategoryCode = categoryCode ?? string.Empty;
        if (string.IsNullOrWhiteSpace(normalizedCategoryCode)) {
            return;
        }

        if (string.Equals(normalizedCategoryCode, "O", StringComparison.Ordinal)) {
            hasRateRequirementCoverage = true;
            if (!string.IsNullOrWhiteSpace(rawRate)) {
                forbiddenRateCategoryCodes.Add(normalizedCategoryCode);
            }

            return;
        }

        if (string.IsNullOrWhiteSpace(rawRate)) {
            missingRateCategoryCodes.Add(normalizedCategoryCode);
        } else {
            hasRateRequirementCoverage = true;
        }
    }

    private static void AddMissingLineTaxFields(
        int lineItemNumber,
        bool hasSettlement,
        bool hasTradeTax,
        bool hasTypeCode,
        bool hasCategoryCode,
        List<string> missingFields) {
        string prefix = "line " + lineItemNumber.ToString(System.Globalization.CultureInfo.InvariantCulture) + " ";
        if (!hasSettlement) {
            missingFields.Add(prefix + "SpecifiedLineTradeSettlement");
        }

        if (!hasTradeTax) {
            missingFields.Add(prefix + "ApplicableTradeTax");
        }

        if (!hasTypeCode) {
            missingFields.Add(prefix + "ApplicableTradeTax TypeCode");
        }

        if (!hasCategoryCode) {
            missingFields.Add(prefix + "ApplicableTradeTax CategoryCode");
        }
    }
}
