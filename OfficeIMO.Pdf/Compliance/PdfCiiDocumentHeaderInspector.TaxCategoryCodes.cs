namespace OfficeIMO.Pdf;

internal static partial class PdfCiiDocumentHeaderInspector {
    internal static bool TryReadTaxCategoryCodes(PdfEmbeddedFile file, out PdfCiiTaxCategoryCodeEvidence? evidence, out string? diagnostic) {
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
                bool hasCategoryCode = false;
                var categoryCodes = new List<string>();
                var headerCategoryCodes = new List<string>();
                var lineCategoryCodes = new List<string>();
                var allowanceChargeCategoryCodes = new List<string>();

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
                        ReadLineTaxCategoryCodeBreakdowns(reader, ref hasApplicableTradeTax, ref hasCategoryCode, categoryCodes, lineCategoryCodes);
                        continue;
                    }

                    if (string.Equals(reader.LocalName, "ApplicableHeaderTradeSettlement", StringComparison.Ordinal)) {
                        ReadHeaderTaxCategoryCodeBreakdowns(reader, ref hasApplicableTradeTax, ref hasCategoryCode, categoryCodes, headerCategoryCodes, allowanceChargeCategoryCodes);
                    }
                }

                if (!sawRoot) {
                    diagnostic = "Attach non-empty UN/CEFACT CrossIndustryInvoice XML in factur-x.xml.";
                    return false;
                }

                int headerNotSubjectTaxBreakdownCount = 0;
                var nonNotSubjectHeaderTaxBreakdownCategories = new List<string>();
                for (int i = 0; i < headerCategoryCodes.Count; i++) {
                    string categoryCode = headerCategoryCodes[i];
                    if (string.Equals(categoryCode, "O", StringComparison.Ordinal)) {
                        headerNotSubjectTaxBreakdownCount++;
                    } else if (!nonNotSubjectHeaderTaxBreakdownCategories.Contains(categoryCode)) {
                        nonNotSubjectHeaderTaxBreakdownCategories.Add(categoryCode);
                    }
                }

                bool hasLineNotSubjectTaxCategory = false;
                for (int i = 0; i < lineCategoryCodes.Count; i++) {
                    if (string.Equals(lineCategoryCodes[i], "O", StringComparison.Ordinal)) {
                        hasLineNotSubjectTaxCategory = true;
                        break;
                    }
                }

                bool hasAllowanceChargeNotSubjectTaxCategory = false;
                for (int i = 0; i < allowanceChargeCategoryCodes.Count; i++) {
                    if (string.Equals(allowanceChargeCategoryCodes[i], "O", StringComparison.Ordinal)) {
                        hasAllowanceChargeNotSubjectTaxCategory = true;
                        break;
                    }
                }

                evidence = new PdfCiiTaxCategoryCodeEvidence(
                    hasApplicableTradeTax,
                    hasCategoryCode,
                    categoryCodes.Distinct(StringComparer.Ordinal).ToArray(),
                    headerNotSubjectTaxBreakdownCount,
                    nonNotSubjectHeaderTaxBreakdownCategories.ToArray(),
                    hasLineNotSubjectTaxCategory,
                    hasAllowanceChargeNotSubjectTaxCategory);
                diagnostic = null;
                return true;
            }
        } catch (System.Xml.XmlException ex) {
            diagnostic = "Attach parseable XML in factur-x.xml: " + ex.Message;
            return false;
        }
    }

    private static void ReadHeaderTaxCategoryCodeBreakdowns(
        System.Xml.XmlReader reader,
        ref bool hasApplicableTradeTax,
        ref bool hasCategoryCode,
        List<string> categoryCodes,
        List<string> headerCategoryCodes,
        List<string> allowanceChargeCategoryCodes) {
        if (reader.IsEmptyElement) {
            return;
        }

        int depth = reader.Depth;
        while (reader.Read()) {
            if (reader.NodeType == System.Xml.XmlNodeType.Element &&
                string.Equals(reader.LocalName, "ApplicableTradeTax", StringComparison.Ordinal)) {
                hasApplicableTradeTax = true;
                ReadTaxCategoryCodeValue(reader, ref hasCategoryCode, categoryCodes, headerCategoryCodes);
                continue;
            }

            if (reader.NodeType == System.Xml.XmlNodeType.Element &&
                string.Equals(reader.LocalName, "SpecifiedTradeAllowanceCharge", StringComparison.Ordinal)) {
                ReadAllowanceChargeTaxCategoryCodeValue(reader, ref hasCategoryCode, categoryCodes, allowanceChargeCategoryCodes);
                continue;
            }

            if (reader.NodeType == System.Xml.XmlNodeType.EndElement &&
                reader.Depth == depth &&
                string.Equals(reader.LocalName, "ApplicableHeaderTradeSettlement", StringComparison.Ordinal)) {
                break;
            }
        }
    }

    private static void ReadAllowanceChargeTaxCategoryCodeValue(
        System.Xml.XmlReader reader,
        ref bool hasCategoryCode,
        List<string> categoryCodes,
        List<string> allowanceChargeCategoryCodes) {
        if (reader.IsEmptyElement) {
            return;
        }

        int depth = reader.Depth;
        while (reader.Read()) {
            if (reader.NodeType == System.Xml.XmlNodeType.Element &&
                reader.Depth == depth + 1 &&
                string.Equals(reader.LocalName, "CategoryTradeTax", StringComparison.Ordinal)) {
                ReadCategoryTradeTaxCategoryCodeValue(reader, ref hasCategoryCode, categoryCodes, allowanceChargeCategoryCodes);
                continue;
            }

            if (reader.NodeType == System.Xml.XmlNodeType.EndElement &&
                reader.Depth == depth &&
                string.Equals(reader.LocalName, "SpecifiedTradeAllowanceCharge", StringComparison.Ordinal)) {
                break;
            }
        }
    }

    private static void ReadCategoryTradeTaxCategoryCodeValue(
        System.Xml.XmlReader reader,
        ref bool hasCategoryCode,
        List<string> categoryCodes,
        List<string> allowanceChargeCategoryCodes) {
        if (reader.IsEmptyElement) {
            return;
        }

        int depth = reader.Depth;
        while (reader.Read()) {
            if (reader.NodeType == System.Xml.XmlNodeType.Element &&
                reader.Depth == depth + 1 &&
                string.Equals(reader.LocalName, "CategoryCode", StringComparison.Ordinal)) {
                string value = ReadElementText(reader);
                if (!string.IsNullOrWhiteSpace(value)) {
                    hasCategoryCode = true;
                    string categoryCode = value.Trim().ToUpperInvariant();
                    categoryCodes.Add(categoryCode);
                    allowanceChargeCategoryCodes.Add(categoryCode);
                }

                continue;
            }

            if (reader.NodeType == System.Xml.XmlNodeType.EndElement &&
                reader.Depth == depth &&
                string.Equals(reader.LocalName, "CategoryTradeTax", StringComparison.Ordinal)) {
                break;
            }
        }
    }

    private static void ReadLineTaxCategoryCodeBreakdowns(
        System.Xml.XmlReader reader,
        ref bool hasApplicableTradeTax,
        ref bool hasCategoryCode,
        List<string> categoryCodes,
        List<string> lineCategoryCodes) {
        if (reader.IsEmptyElement) {
            return;
        }

        int depth = reader.Depth;
        while (reader.Read()) {
            if (reader.NodeType == System.Xml.XmlNodeType.Element &&
                string.Equals(reader.LocalName, "SpecifiedLineTradeSettlement", StringComparison.Ordinal)) {
                ReadLineSettlementTaxCategoryCodeBreakdowns(reader, ref hasApplicableTradeTax, ref hasCategoryCode, categoryCodes, lineCategoryCodes);
                continue;
            }

            if (reader.NodeType == System.Xml.XmlNodeType.EndElement &&
                reader.Depth == depth &&
                string.Equals(reader.LocalName, "IncludedSupplyChainTradeLineItem", StringComparison.Ordinal)) {
                break;
            }
        }
    }

    private static void ReadLineSettlementTaxCategoryCodeBreakdowns(
        System.Xml.XmlReader reader,
        ref bool hasApplicableTradeTax,
        ref bool hasCategoryCode,
        List<string> categoryCodes,
        List<string> lineCategoryCodes) {
        if (reader.IsEmptyElement) {
            return;
        }

        int depth = reader.Depth;
        while (reader.Read()) {
            if (reader.NodeType == System.Xml.XmlNodeType.Element &&
                string.Equals(reader.LocalName, "ApplicableTradeTax", StringComparison.Ordinal)) {
                hasApplicableTradeTax = true;
                ReadTaxCategoryCodeValue(reader, ref hasCategoryCode, categoryCodes, lineCategoryCodes);
                continue;
            }

            if (reader.NodeType == System.Xml.XmlNodeType.EndElement &&
                reader.Depth == depth &&
                string.Equals(reader.LocalName, "SpecifiedLineTradeSettlement", StringComparison.Ordinal)) {
                break;
            }
        }
    }

    private static void ReadTaxCategoryCodeValue(System.Xml.XmlReader reader, ref bool hasCategoryCode, List<string> categoryCodes, List<string> scopedCategoryCodes) {
        if (reader.IsEmptyElement) {
            return;
        }

        int depth = reader.Depth;
        while (reader.Read()) {
            if (reader.NodeType == System.Xml.XmlNodeType.Element &&
                reader.Depth == depth + 1 &&
                string.Equals(reader.LocalName, "CategoryCode", StringComparison.Ordinal)) {
                string value = ReadElementText(reader);
                if (!string.IsNullOrWhiteSpace(value)) {
                    hasCategoryCode = true;
                    string categoryCode = value.Trim().ToUpperInvariant();
                    categoryCodes.Add(categoryCode);
                    scopedCategoryCodes.Add(categoryCode);
                }

                continue;
            }

            if (reader.NodeType == System.Xml.XmlNodeType.EndElement &&
                reader.Depth == depth &&
                string.Equals(reader.LocalName, "ApplicableTradeTax", StringComparison.Ordinal)) {
                break;
            }
        }
    }
}
