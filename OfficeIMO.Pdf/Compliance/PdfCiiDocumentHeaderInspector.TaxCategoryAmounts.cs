namespace OfficeIMO.Pdf;

internal static partial class PdfCiiDocumentHeaderInspector {
    internal static bool TryReadTaxCategoryAmounts(PdfEmbeddedFile file, out PdfCiiTaxCategoryAmountEvidence? evidence, out string? diagnostic) {
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
                bool hasTaxCategoryAmount = false;
                var nonZeroRatedCategoryAmounts = new List<string>();
                var mismatchedStandardRatedCategoryAmounts = new List<string>();
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

                    if (string.Equals(reader.LocalName, "ApplicableHeaderTradeSettlement", StringComparison.Ordinal)) {
                        ReadHeaderTaxCategoryAmounts(reader, ref hasApplicableTradeTax, ref hasTaxCategoryAmount, nonZeroRatedCategoryAmounts, mismatchedStandardRatedCategoryAmounts, ref parseDiagnostic);
                    }
                }

                if (!sawRoot) {
                    diagnostic = "Attach non-empty UN/CEFACT CrossIndustryInvoice XML in factur-x.xml.";
                    return false;
                }

                evidence = new PdfCiiTaxCategoryAmountEvidence(
                    hasApplicableTradeTax,
                    hasTaxCategoryAmount,
                    nonZeroRatedCategoryAmounts.Count == 0 && parseDiagnostic == null,
                    nonZeroRatedCategoryAmounts.Distinct(StringComparer.Ordinal).ToArray(),
                    mismatchedStandardRatedCategoryAmounts.Count == 0 && parseDiagnostic == null,
                    mismatchedStandardRatedCategoryAmounts.Distinct(StringComparer.Ordinal).ToArray(),
                    parseDiagnostic);
                diagnostic = null;
                return true;
            }
        } catch (System.Xml.XmlException ex) {
            diagnostic = "Attach parseable XML in factur-x.xml: " + ex.Message;
            return false;
        }
    }

    private static void ReadHeaderTaxCategoryAmounts(
        System.Xml.XmlReader reader,
        ref bool hasApplicableTradeTax,
        ref bool hasTaxCategoryAmount,
        List<string> nonZeroRatedCategoryAmounts,
        List<string> mismatchedStandardRatedCategoryAmounts,
        ref string? parseDiagnostic) {
        if (reader.IsEmptyElement) {
            return;
        }

        int depth = reader.Depth;
        while (reader.Read()) {
            if (reader.NodeType == System.Xml.XmlNodeType.Element &&
                string.Equals(reader.LocalName, "ApplicableTradeTax", StringComparison.Ordinal)) {
                hasApplicableTradeTax = true;
                ReadTaxCategoryAmountSemantics(reader, ref hasTaxCategoryAmount, nonZeroRatedCategoryAmounts, mismatchedStandardRatedCategoryAmounts, ref parseDiagnostic);
                continue;
            }

            if (reader.NodeType == System.Xml.XmlNodeType.EndElement &&
                reader.Depth == depth &&
                string.Equals(reader.LocalName, "ApplicableHeaderTradeSettlement", StringComparison.Ordinal)) {
                break;
            }
        }
    }

    private static void ReadTaxCategoryAmountSemantics(
        System.Xml.XmlReader reader,
        ref bool hasTaxCategoryAmount,
        List<string> nonZeroRatedCategoryAmounts,
        List<string> mismatchedStandardRatedCategoryAmounts,
        ref string? parseDiagnostic) {
        if (reader.IsEmptyElement) {
            return;
        }

        string? categoryCode = null;
        decimal? basisAmount = null;
        decimal? rate = null;
        decimal? calculatedAmount = null;
        string? rawCalculatedAmount = null;
        int depth = reader.Depth;
        while (reader.Read()) {
            if (reader.NodeType == System.Xml.XmlNodeType.Element && reader.Depth == depth + 1) {
                if (string.Equals(reader.LocalName, "CategoryCode", StringComparison.Ordinal)) {
                    categoryCode = NormalizeTaxCategoryCode(ReadElementText(reader));
                    continue;
                }

                if (string.Equals(reader.LocalName, "BasisAmount", StringComparison.Ordinal)) {
                    if (TryReadAmount(reader, "ApplicableTradeTax BasisAmount", ref parseDiagnostic, out decimal? amount)) {
                        basisAmount = amount;
                    }

                    continue;
                }

                if (string.Equals(reader.LocalName, "RateApplicablePercent", StringComparison.Ordinal)) {
                    string rawRate = ReadElementText(reader).Trim();
                    if (TryParseCiiDecimal(rawRate, out decimal parsedRate)) {
                        rate = parsedRate;
                    } else if (parseDiagnostic == null) {
                        parseDiagnostic = "Set factur-x.xml ApplicableTradeTax RateApplicablePercent to a parseable decimal percentage. Found: " + rawRate + ".";
                    }

                    continue;
                }

                if (string.Equals(reader.LocalName, "CalculatedAmount", StringComparison.Ordinal)) {
                    rawCalculatedAmount = ReadElementText(reader).Trim();
                    if (TryParseCiiDecimal(rawCalculatedAmount, out decimal parsedAmount)) {
                        calculatedAmount = parsedAmount;
                    } else if (parseDiagnostic == null) {
                        parseDiagnostic = "Set factur-x.xml ApplicableTradeTax CalculatedAmount to a parseable decimal amount. Found: " + rawCalculatedAmount + ".";
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

        if (string.IsNullOrWhiteSpace(categoryCode) || string.IsNullOrWhiteSpace(rawCalculatedAmount)) {
            return;
        }

        hasTaxCategoryAmount = true;
        if (calculatedAmount.HasValue && RequiresZeroTaxAmount(categoryCode!) && calculatedAmount.Value != 0m) {
            nonZeroRatedCategoryAmounts.Add(categoryCode + "/" + calculatedAmount.Value.ToString("0.#############################", System.Globalization.CultureInfo.InvariantCulture));
        }

        if (string.Equals(categoryCode, "S", StringComparison.Ordinal) &&
            basisAmount.HasValue &&
            rate.HasValue &&
            calculatedAmount.HasValue) {
            decimal expectedAmount = decimal.Round(basisAmount.Value * rate.Value / 100m, 2, MidpointRounding.AwayFromZero);
            if (System.Math.Abs(calculatedAmount.Value - expectedAmount) > 1.00m) {
                mismatchedStandardRatedCategoryAmounts.Add(
                    categoryCode +
                    "/" +
                    rate.Value.ToString("0.#############################", System.Globalization.CultureInfo.InvariantCulture) +
                    " expected " +
                    expectedAmount.ToString("0.00", System.Globalization.CultureInfo.InvariantCulture) +
                    " from taxable basis " +
                    basisAmount.Value.ToString("0.00", System.Globalization.CultureInfo.InvariantCulture) +
                    " but found " +
                    calculatedAmount.Value.ToString("0.00", System.Globalization.CultureInfo.InvariantCulture));
            }
        }
    }

    private static bool RequiresZeroTaxAmount(string categoryCode) =>
        string.Equals(categoryCode, "AE", StringComparison.Ordinal) ||
        string.Equals(categoryCode, "E", StringComparison.Ordinal) ||
        string.Equals(categoryCode, "G", StringComparison.Ordinal) ||
        string.Equals(categoryCode, "K", StringComparison.Ordinal) ||
        string.Equals(categoryCode, "O", StringComparison.Ordinal) ||
        string.Equals(categoryCode, "Z", StringComparison.Ordinal);
}
