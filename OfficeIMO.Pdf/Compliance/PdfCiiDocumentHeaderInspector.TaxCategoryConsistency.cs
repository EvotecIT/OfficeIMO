namespace OfficeIMO.Pdf;

internal static partial class PdfCiiDocumentHeaderInspector {
    internal static bool TryReadTaxCategoryConsistency(PdfEmbeddedFile file, out PdfCiiTaxCategoryConsistencyEvidence? evidence, out string? diagnostic) {
        Guard.NotNull(file, nameof(file));
        evidence = null;

        try {
            using (var stream = new MemoryStream(file.DataSnapshot))
            using (var reader = System.Xml.XmlReader.Create(stream, new System.Xml.XmlReaderSettings {
                DtdProcessing = System.Xml.DtdProcessing.Prohibit,
                XmlResolver = null
            })) {
                bool sawRoot = false;
                bool hasHeaderApplicableTradeTax = false;
                bool hasLineApplicableTradeTax = false;
                var headerCategoryRates = new List<string>();
                var lineCategoryRates = new List<string>();
                var allowanceChargeCategoryRates = new List<string>();

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
                        ReadLineTaxCategoryRates(reader, lineCategoryRates, ref hasLineApplicableTradeTax);
                        continue;
                    }

                    if (string.Equals(reader.LocalName, "ApplicableHeaderTradeSettlement", StringComparison.Ordinal)) {
                        ReadHeaderTaxCategoryRates(reader, headerCategoryRates, allowanceChargeCategoryRates, ref hasHeaderApplicableTradeTax);
                    }
                }

                if (!sawRoot) {
                    diagnostic = "Attach non-empty UN/CEFACT CrossIndustryInvoice XML in factur-x.xml.";
                    return false;
                }

                var headerSet = new HashSet<string>(headerCategoryRates, StringComparer.Ordinal);
                var unmatched = new List<string>();
                for (int i = 0; i < lineCategoryRates.Count; i++) {
                    string lineCategoryRate = lineCategoryRates[i];
                    if (!headerSet.Contains(lineCategoryRate) && !unmatched.Contains(lineCategoryRate)) {
                        unmatched.Add(lineCategoryRate);
                    }
                }

                var unmatchedAllowanceChargeCategoryRates = new List<string>();
                for (int i = 0; i < allowanceChargeCategoryRates.Count; i++) {
                    string allowanceChargeCategoryRate = allowanceChargeCategoryRates[i];
                    if (!headerSet.Contains(allowanceChargeCategoryRate) && !unmatchedAllowanceChargeCategoryRates.Contains(allowanceChargeCategoryRate)) {
                        unmatchedAllowanceChargeCategoryRates.Add(allowanceChargeCategoryRate);
                    }
                }

                evidence = new PdfCiiTaxCategoryConsistencyEvidence(
                    hasHeaderApplicableTradeTax,
                    hasLineApplicableTradeTax,
                    headerCategoryRates.Count > 0,
                    lineCategoryRates.Count > 0,
                    unmatched.Count == 0,
                    unmatched,
                    unmatchedAllowanceChargeCategoryRates.Count == 0,
                    unmatchedAllowanceChargeCategoryRates);
                diagnostic = null;
                return true;
            }
        } catch (System.Xml.XmlException ex) {
            diagnostic = "Attach parseable XML in factur-x.xml: " + ex.Message;
            return false;
        }
    }

    private static void ReadHeaderTaxCategoryRates(
        System.Xml.XmlReader reader,
        List<string> categoryRates,
        List<string> allowanceChargeCategoryRates,
        ref bool hasHeaderApplicableTradeTax) {
        if (reader.IsEmptyElement) {
            return;
        }

        int depth = reader.Depth;
        while (reader.Read()) {
            if (reader.NodeType == System.Xml.XmlNodeType.Element) {
                if (string.Equals(reader.LocalName, "ApplicableTradeTax", StringComparison.Ordinal)) {
                    hasHeaderApplicableTradeTax = true;
                    if (TryReadTaxCategoryRate(reader, "ApplicableTradeTax", out string? categoryRate)) {
                        categoryRates.Add(categoryRate!);
                    }

                    continue;
                }

                if (string.Equals(reader.LocalName, "SpecifiedTradeAllowanceCharge", StringComparison.Ordinal)) {
                    ReadAllowanceChargeTaxCategoryRates(reader, allowanceChargeCategoryRates);
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

    private static void ReadAllowanceChargeTaxCategoryRates(System.Xml.XmlReader reader, List<string> categoryRates) {
        if (reader.IsEmptyElement) {
            return;
        }

        int depth = reader.Depth;
        while (reader.Read()) {
            if (reader.NodeType == System.Xml.XmlNodeType.Element &&
                reader.Depth == depth + 1 &&
                string.Equals(reader.LocalName, "CategoryTradeTax", StringComparison.Ordinal)) {
                if (TryReadTaxCategoryRate(reader, "CategoryTradeTax", out string? categoryRate)) {
                    categoryRates.Add(categoryRate!);
                }

                continue;
            }

            if (reader.NodeType == System.Xml.XmlNodeType.EndElement &&
                reader.Depth == depth &&
                string.Equals(reader.LocalName, "SpecifiedTradeAllowanceCharge", StringComparison.Ordinal)) {
                break;
            }
        }
    }

    private static void ReadLineTaxCategoryRates(System.Xml.XmlReader reader, List<string> categoryRates, ref bool hasLineApplicableTradeTax) {
        if (reader.IsEmptyElement) {
            return;
        }

        int depth = reader.Depth;
        while (reader.Read()) {
            if (reader.NodeType == System.Xml.XmlNodeType.Element) {
                if (string.Equals(reader.LocalName, "SpecifiedLineTradeSettlement", StringComparison.Ordinal)) {
                    ReadLineSettlementTaxCategoryRates(reader, categoryRates, ref hasLineApplicableTradeTax);
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

    private static void ReadLineSettlementTaxCategoryRates(System.Xml.XmlReader reader, List<string> categoryRates, ref bool hasLineApplicableTradeTax) {
        if (reader.IsEmptyElement) {
            return;
        }

        int depth = reader.Depth;
        while (reader.Read()) {
            if (reader.NodeType == System.Xml.XmlNodeType.Element) {
                if (string.Equals(reader.LocalName, "ApplicableTradeTax", StringComparison.Ordinal)) {
                    hasLineApplicableTradeTax = true;
                    if (TryReadTaxCategoryRate(reader, "ApplicableTradeTax", out string? categoryRate)) {
                        categoryRates.Add(categoryRate!);
                    }

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

    private static bool TryReadTaxCategoryRate(System.Xml.XmlReader reader, string containerName, out string? categoryRate) {
        categoryRate = null;
        if (reader.IsEmptyElement) {
            return false;
        }

        string? categoryCode = null;
        string? rate = null;
        int depth = reader.Depth;
        while (reader.Read()) {
            if (reader.NodeType == System.Xml.XmlNodeType.Element && reader.Depth == depth + 1) {
                if (string.Equals(reader.LocalName, "CategoryCode", StringComparison.Ordinal)) {
                    categoryCode = NormalizeTaxCategoryCode(ReadElementText(reader));
                    continue;
                }

                if (string.Equals(reader.LocalName, "RateApplicablePercent", StringComparison.Ordinal)) {
                    rate = NormalizeTaxRate(ReadElementText(reader));
                    continue;
                }
            }

            if (reader.NodeType == System.Xml.XmlNodeType.EndElement &&
                reader.Depth == depth &&
                string.Equals(reader.LocalName, containerName, StringComparison.Ordinal)) {
                break;
            }
        }

        string normalizedCategoryCode = categoryCode ?? string.Empty;
        if (string.IsNullOrWhiteSpace(normalizedCategoryCode)) {
            return false;
        }

        if (string.Equals(normalizedCategoryCode, "O", StringComparison.Ordinal)) {
            categoryRate = normalizedCategoryCode;
            return true;
        }

        if (string.IsNullOrWhiteSpace(rate)) {
            return false;
        }

        categoryRate = normalizedCategoryCode + "/" + rate;
        return true;
    }

    private static string NormalizeTaxCategoryCode(string value) => value.Trim().ToUpperInvariant();

    private static string NormalizeTaxRate(string value) {
        string trimmed = value.Trim();
        if (TryParseCiiDecimal(trimmed, out decimal rate)) {
            return rate.ToString("0.#############################", System.Globalization.CultureInfo.InvariantCulture);
        }

        return trimmed;
    }
}
