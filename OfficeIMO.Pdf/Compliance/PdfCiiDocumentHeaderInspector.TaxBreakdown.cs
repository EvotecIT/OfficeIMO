namespace OfficeIMO.Pdf;

internal static partial class PdfCiiDocumentHeaderInspector {
    internal static bool TryReadTaxBreakdown(PdfEmbeddedFile file, out PdfCiiTaxBreakdownEvidence? evidence, out string? diagnostic) {
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
                bool hasTypeCode = false;
                bool hasCategoryCode = false;
                bool hasRateApplicablePercent = false;
                bool hasBasisAmount = false;
                bool hasCalculatedAmount = false;
                var typeCodes = new List<string>();
                var missingTypeCodeBreakdowns = new List<string>();

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
                        ReadHeaderTradeSettlementTaxBreakdown(reader, ref hasApplicableTradeTax, ref hasTypeCode, ref hasCategoryCode, ref hasRateApplicablePercent, ref hasBasisAmount, ref hasCalculatedAmount, typeCodes, missingTypeCodeBreakdowns);
                    }
                }

                if (!sawRoot) {
                    diagnostic = "Attach non-empty UN/CEFACT CrossIndustryInvoice XML in factur-x.xml.";
                    return false;
                }

                evidence = new PdfCiiTaxBreakdownEvidence(hasApplicableTradeTax, hasTypeCode && missingTypeCodeBreakdowns.Count == 0, hasCategoryCode, hasRateApplicablePercent, hasBasisAmount, hasCalculatedAmount, typeCodes, missingTypeCodeBreakdowns);
                diagnostic = null;
                return true;
            }
        } catch (System.Xml.XmlException ex) {
            diagnostic = "Attach parseable XML in factur-x.xml: " + ex.Message;
            return false;
        }
    }

    private static void ReadHeaderTradeSettlementTaxBreakdown(
        System.Xml.XmlReader reader,
        ref bool hasApplicableTradeTax,
        ref bool hasTypeCode,
        ref bool hasCategoryCode,
        ref bool hasRateApplicablePercent,
        ref bool hasBasisAmount,
        ref bool hasCalculatedAmount,
        List<string> typeCodes,
        List<string> missingTypeCodeBreakdowns) {
        if (reader.IsEmptyElement) {
            return;
        }

        int depth = reader.Depth;
        int breakdownIndex = 0;
        while (reader.Read()) {
            if (reader.NodeType == System.Xml.XmlNodeType.Element) {
                if (string.Equals(reader.LocalName, "ApplicableTradeTax", StringComparison.Ordinal)) {
                    hasApplicableTradeTax = true;
                    breakdownIndex++;
                    ReadTradeTax(reader, breakdownIndex, ref hasTypeCode, ref hasCategoryCode, ref hasRateApplicablePercent, ref hasBasisAmount, ref hasCalculatedAmount, typeCodes, missingTypeCodeBreakdowns);
                    continue;
                }
            }

            if (reader.NodeType == System.Xml.XmlNodeType.EndElement &&
                reader.Depth == depth &&
                string.Equals(reader.LocalName, "ApplicableHeaderTradeSettlement", StringComparison.Ordinal)) {
                break;
            }
        }
    }

    private static void ReadTradeTax(
        System.Xml.XmlReader reader,
        int breakdownIndex,
        ref bool hasTypeCode,
        ref bool hasCategoryCode,
        ref bool hasRateApplicablePercent,
        ref bool hasBasisAmount,
        ref bool hasCalculatedAmount,
        List<string> typeCodes,
        List<string> missingTypeCodeBreakdowns) {
        if (reader.IsEmptyElement) {
            missingTypeCodeBreakdowns.Add("ApplicableTradeTax #" + breakdownIndex.ToString(System.Globalization.CultureInfo.InvariantCulture));
            return;
        }

        int depth = reader.Depth;
        bool tradeTaxHasTypeCode = false;
        while (reader.Read()) {
            if (reader.NodeType == System.Xml.XmlNodeType.Element) {
                if (reader.Depth == depth + 1 && string.Equals(reader.LocalName, "TypeCode", StringComparison.Ordinal)) {
                    string typeCode = ReadElementText(reader);
                    if (!string.IsNullOrWhiteSpace(typeCode)) {
                        hasTypeCode = true;
                        tradeTaxHasTypeCode = true;
                        typeCodes.Add(typeCode.Trim());
                    }

                    continue;
                }

                if (reader.Depth == depth + 1 && string.Equals(reader.LocalName, "CategoryCode", StringComparison.Ordinal)) {
                    hasCategoryCode = hasCategoryCode || !string.IsNullOrWhiteSpace(ReadElementText(reader));
                    continue;
                }

                if (reader.Depth == depth + 1 && string.Equals(reader.LocalName, "RateApplicablePercent", StringComparison.Ordinal)) {
                    hasRateApplicablePercent = hasRateApplicablePercent || !string.IsNullOrWhiteSpace(ReadElementText(reader));
                    continue;
                }

                if (reader.Depth == depth + 1 && string.Equals(reader.LocalName, "BasisAmount", StringComparison.Ordinal)) {
                    hasBasisAmount = hasBasisAmount || !string.IsNullOrWhiteSpace(ReadElementText(reader));
                    continue;
                }

                if (reader.Depth == depth + 1 && string.Equals(reader.LocalName, "CalculatedAmount", StringComparison.Ordinal)) {
                    hasCalculatedAmount = hasCalculatedAmount || !string.IsNullOrWhiteSpace(ReadElementText(reader));
                    continue;
                }
            }

            if (reader.NodeType == System.Xml.XmlNodeType.EndElement &&
                reader.Depth == depth &&
                string.Equals(reader.LocalName, "ApplicableTradeTax", StringComparison.Ordinal)) {
                break;
            }
        }

        if (!tradeTaxHasTypeCode) {
            missingTypeCodeBreakdowns.Add("ApplicableTradeTax #" + breakdownIndex.ToString(System.Globalization.CultureInfo.InvariantCulture));
        }
    }
}
