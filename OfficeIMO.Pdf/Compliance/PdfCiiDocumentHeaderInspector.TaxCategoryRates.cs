namespace OfficeIMO.Pdf;

internal static partial class PdfCiiDocumentHeaderInspector {
    internal static bool TryReadTaxCategoryRates(PdfEmbeddedFile file, out PdfCiiTaxCategoryRateEvidence? evidence, out string? diagnostic) {
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
                bool hasTaxCategoryRate = false;
                bool hasRateRequirementCoverage = false;
                var nonZeroRatedCategoryRates = new List<string>();
                var forbiddenRateCategoryCodes = new List<string>();
                string? parseDiagnostic = null;

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

                    if (string.Equals(reader.LocalName, "ApplicableTradeTax", StringComparison.Ordinal)) {
                        hasApplicableTradeTax = true;
                        ReadTaxCategoryRateSemantics(reader, ref hasTaxCategoryRate, ref hasRateRequirementCoverage, nonZeroRatedCategoryRates, forbiddenRateCategoryCodes, ref parseDiagnostic);
                        continue;
                    }

                    if (string.Equals(reader.LocalName, "SpecifiedTradeAllowanceCharge", StringComparison.Ordinal)) {
                        ReadAllowanceChargeTaxCategoryRateSemantics(reader, ref hasRateRequirementCoverage, forbiddenRateCategoryCodes, ref parseDiagnostic);
                    }
                }

                if (!sawRoot) {
                    diagnostic = "Attach non-empty UN/CEFACT CrossIndustryInvoice XML in factur-x.xml.";
                    return false;
                }

                evidence = new PdfCiiTaxCategoryRateEvidence(
                    hasApplicableTradeTax,
                    hasTaxCategoryRate,
                    hasRateRequirementCoverage,
                    nonZeroRatedCategoryRates.Count == 0 && parseDiagnostic == null,
                    nonZeroRatedCategoryRates.Distinct(StringComparer.Ordinal).ToArray(),
                    forbiddenRateCategoryCodes.Distinct(StringComparer.Ordinal).ToArray(),
                    parseDiagnostic);
                diagnostic = null;
                return true;
            }
        } catch (System.Xml.XmlException ex) {
            diagnostic = "Attach parseable XML in factur-x.xml: " + ex.Message;
            return false;
        }
    }

    private static void ReadAllowanceChargeTaxCategoryRateSemantics(
        System.Xml.XmlReader reader,
        ref bool hasRateRequirementCoverage,
        List<string> forbiddenRateCategoryCodes,
        ref string? parseDiagnostic) {
        if (reader.IsEmptyElement) {
            return;
        }

        bool? isCharge = null;
        string? categoryCode = null;
        string? rawRate = null;
        int depth = reader.Depth;
        while (reader.Read()) {
            if (reader.NodeType == System.Xml.XmlNodeType.Element) {
                if (reader.Depth == depth + 1 && string.Equals(reader.LocalName, "ChargeIndicator", StringComparison.Ordinal)) {
                    ReadChargeIndicator(reader, ref isCharge);
                    continue;
                }

                if (reader.Depth == depth + 1 && string.Equals(reader.LocalName, "CategoryTradeTax", StringComparison.Ordinal)) {
                    ReadAllowanceChargeCategoryTradeTaxRateSemantics(reader, ref categoryCode, ref rawRate, ref parseDiagnostic);
                    continue;
                }
            }

            if (reader.NodeType == System.Xml.XmlNodeType.EndElement &&
                reader.Depth == depth &&
                string.Equals(reader.LocalName, "SpecifiedTradeAllowanceCharge", StringComparison.Ordinal)) {
                break;
            }
        }

        string normalizedCategoryCode = categoryCode ?? string.Empty;
        if (!string.Equals(normalizedCategoryCode, "O", StringComparison.Ordinal)) {
            return;
        }

        hasRateRequirementCoverage = true;
        if (!string.IsNullOrWhiteSpace(rawRate)) {
            string allowanceChargeKind = isCharge == true ? "document-level charge" : "document-level allowance";
            forbiddenRateCategoryCodes.Add(normalizedCategoryCode + " " + allowanceChargeKind);
        }
    }

    private static void ReadAllowanceChargeCategoryTradeTaxRateSemantics(
        System.Xml.XmlReader reader,
        ref string? categoryCode,
        ref string? rawRate,
        ref string? parseDiagnostic) {
        if (reader.IsEmptyElement) {
            return;
        }

        int depth = reader.Depth;
        while (reader.Read()) {
            if (reader.NodeType == System.Xml.XmlNodeType.Element && reader.Depth == depth + 1) {
                if (string.Equals(reader.LocalName, "CategoryCode", StringComparison.Ordinal)) {
                    categoryCode = NormalizeTaxCategoryCode(ReadElementText(reader));
                    continue;
                }

                if (string.Equals(reader.LocalName, "RateApplicablePercent", StringComparison.Ordinal)) {
                    rawRate = ReadElementText(reader).Trim();
                    if (!decimal.TryParse(rawRate, System.Globalization.NumberStyles.Number, System.Globalization.CultureInfo.InvariantCulture, out _) &&
                        parseDiagnostic == null) {
                        parseDiagnostic = "Set factur-x.xml CategoryTradeTax RateApplicablePercent to a parseable decimal percentage. Found: " + rawRate + ".";
                    }

                    continue;
                }
            }

            if (reader.NodeType == System.Xml.XmlNodeType.EndElement &&
                reader.Depth == depth &&
                string.Equals(reader.LocalName, "CategoryTradeTax", StringComparison.Ordinal)) {
                break;
            }
        }
    }

    private static void ReadTaxCategoryRateSemantics(
        System.Xml.XmlReader reader,
        ref bool hasTaxCategoryRate,
        ref bool hasRateRequirementCoverage,
        List<string> nonZeroRatedCategoryRates,
        List<string> forbiddenRateCategoryCodes,
        ref string? parseDiagnostic) {
        if (reader.IsEmptyElement) {
            return;
        }

        string? categoryCode = null;
        decimal? rate = null;
        string? rawRate = null;
        int depth = reader.Depth;
        while (reader.Read()) {
            if (reader.NodeType == System.Xml.XmlNodeType.Element && reader.Depth == depth + 1) {
                if (string.Equals(reader.LocalName, "CategoryCode", StringComparison.Ordinal)) {
                    categoryCode = NormalizeTaxCategoryCode(ReadElementText(reader));
                    continue;
                }

                if (string.Equals(reader.LocalName, "RateApplicablePercent", StringComparison.Ordinal)) {
                    rawRate = ReadElementText(reader).Trim();
                    if (decimal.TryParse(rawRate, System.Globalization.NumberStyles.Number, System.Globalization.CultureInfo.InvariantCulture, out decimal parsedRate)) {
                        rate = parsedRate;
                    } else if (parseDiagnostic == null) {
                        parseDiagnostic = "Set factur-x.xml ApplicableTradeTax RateApplicablePercent to a parseable decimal percentage. Found: " + rawRate + ".";
                    }

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
            return;
        }

        hasTaxCategoryRate = true;
        hasRateRequirementCoverage = true;
        if (rate.HasValue && RequiresZeroTaxRate(normalizedCategoryCode) && rate.Value != 0m) {
            nonZeroRatedCategoryRates.Add(normalizedCategoryCode + "/" + rate.Value.ToString("0.#############################", System.Globalization.CultureInfo.InvariantCulture));
        }
    }

    private static bool RequiresZeroTaxRate(string categoryCode) =>
        string.Equals(categoryCode, "AE", StringComparison.Ordinal) ||
        string.Equals(categoryCode, "E", StringComparison.Ordinal) ||
        string.Equals(categoryCode, "G", StringComparison.Ordinal) ||
        string.Equals(categoryCode, "K", StringComparison.Ordinal) ||
        string.Equals(categoryCode, "Z", StringComparison.Ordinal);
}
