namespace OfficeIMO.Pdf;

internal static partial class PdfCiiDocumentHeaderInspector {
    internal static bool TryReadTaxTotalConsistency(PdfEmbeddedFile file, out PdfCiiTaxTotalConsistencyEvidence? evidence, out string? diagnostic) {
        Guard.NotNull(file, nameof(file));
        evidence = null;

        try {
            using (var stream = new MemoryStream(file.DataSnapshot))
            using (var reader = System.Xml.XmlReader.Create(stream, new System.Xml.XmlReaderSettings {
                DtdProcessing = System.Xml.DtdProcessing.Prohibit,
                XmlResolver = null
            })) {
                bool sawRoot = false;
                decimal basisBreakdownSum = 0m;
                bool hasBasisBreakdownAmount = false;
                decimal calculatedBreakdownSum = 0m;
                bool hasCalculatedBreakdownAmount = false;
                decimal? taxBasisTotalAmount = null;
                decimal? taxTotalAmount = null;
                decimal? notSubjectHeaderBasisAmount = null;
                decimal notSubjectLineNetAmountSum = 0m;
                bool hasNotSubjectLineNetAmount = false;
                decimal notSubjectAllowanceAmountSum = 0m;
                bool hasNotSubjectAllowanceAmount = false;
                decimal notSubjectChargeAmountSum = 0m;
                bool hasNotSubjectChargeAmount = false;
                var headerBasisAmountsByCategoryRate = new Dictionary<string, decimal>(StringComparer.Ordinal);
                var lineNetAmountsByCategoryRate = new Dictionary<string, decimal>(StringComparer.Ordinal);
                var allowanceAmountsByCategoryRate = new Dictionary<string, decimal>(StringComparer.Ordinal);
                var chargeAmountsByCategoryRate = new Dictionary<string, decimal>(StringComparer.Ordinal);
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

                    if (string.Equals(reader.LocalName, "IncludedSupplyChainTradeLineItem", StringComparison.Ordinal)) {
                        ReadLineTaxBasisContribution(
                            reader,
                            lineNetAmountsByCategoryRate,
                            ref notSubjectLineNetAmountSum,
                            ref hasNotSubjectLineNetAmount,
                            ref parseDiagnostic);
                        continue;
                    }

                    if (string.Equals(reader.LocalName, "ApplicableHeaderTradeSettlement", StringComparison.Ordinal)) {
                        ReadHeaderTaxTotals(
                            reader,
                            ref basisBreakdownSum,
                            ref hasBasisBreakdownAmount,
                            ref calculatedBreakdownSum,
                            ref hasCalculatedBreakdownAmount,
                            ref taxBasisTotalAmount,
                            ref taxTotalAmount,
                            ref notSubjectHeaderBasisAmount,
                            ref notSubjectAllowanceAmountSum,
                            ref hasNotSubjectAllowanceAmount,
                            ref notSubjectChargeAmountSum,
                            ref hasNotSubjectChargeAmount,
                            headerBasisAmountsByCategoryRate,
                            allowanceAmountsByCategoryRate,
                            chargeAmountsByCategoryRate,
                            ref parseDiagnostic);
                    }
                }

                if (!sawRoot) {
                    diagnostic = "Attach non-empty UN/CEFACT CrossIndustryInvoice XML in factur-x.xml.";
                    return false;
                }

                evidence = new PdfCiiTaxTotalConsistencyEvidence(
                    hasBasisBreakdownAmount ? basisBreakdownSum : (decimal?)null,
                    hasCalculatedBreakdownAmount ? calculatedBreakdownSum : (decimal?)null,
                    taxBasisTotalAmount,
                    taxTotalAmount,
                    notSubjectHeaderBasisAmount,
                    hasNotSubjectLineNetAmount ? notSubjectLineNetAmountSum : (decimal?)null,
                    hasNotSubjectAllowanceAmount ? notSubjectAllowanceAmountSum : (decimal?)null,
                    hasNotSubjectChargeAmount ? notSubjectChargeAmountSum : (decimal?)null,
                    BuildAdjustedBasisMismatches(headerBasisAmountsByCategoryRate, lineNetAmountsByCategoryRate, allowanceAmountsByCategoryRate, chargeAmountsByCategoryRate),
                    parseDiagnostic);
                diagnostic = null;
                return true;
            }
        } catch (System.Xml.XmlException ex) {
            diagnostic = "Attach parseable XML in factur-x.xml: " + ex.Message;
            return false;
        }
    }

    private static void ReadHeaderTaxTotals(
        System.Xml.XmlReader reader,
        ref decimal basisBreakdownSum,
        ref bool hasBasisBreakdownAmount,
        ref decimal calculatedBreakdownSum,
        ref bool hasCalculatedBreakdownAmount,
        ref decimal? taxBasisTotalAmount,
        ref decimal? taxTotalAmount,
        ref decimal? notSubjectHeaderBasisAmount,
        ref decimal notSubjectAllowanceAmountSum,
        ref bool hasNotSubjectAllowanceAmount,
        ref decimal notSubjectChargeAmountSum,
        ref bool hasNotSubjectChargeAmount,
        Dictionary<string, decimal> headerBasisAmountsByCategoryRate,
        Dictionary<string, decimal> allowanceAmountsByCategoryRate,
        Dictionary<string, decimal> chargeAmountsByCategoryRate,
        ref string? parseDiagnostic) {
        if (reader.IsEmptyElement) {
            return;
        }

        int depth = reader.Depth;
        while (reader.Read()) {
            if (reader.NodeType == System.Xml.XmlNodeType.Element) {
                if (string.Equals(reader.LocalName, "ApplicableTradeTax", StringComparison.Ordinal)) {
                    ReadApplicableTradeTaxTotals(reader, ref basisBreakdownSum, ref hasBasisBreakdownAmount, ref calculatedBreakdownSum, ref hasCalculatedBreakdownAmount, ref notSubjectHeaderBasisAmount, headerBasisAmountsByCategoryRate, ref parseDiagnostic);
                    continue;
                }

                if (string.Equals(reader.LocalName, "SpecifiedTradeSettlementHeaderMonetarySummation", StringComparison.Ordinal)) {
                    ReadHeaderMonetaryTaxTotals(reader, ref taxBasisTotalAmount, ref taxTotalAmount, ref parseDiagnostic);
                    continue;
                }

                if (string.Equals(reader.LocalName, "SpecifiedTradeAllowanceCharge", StringComparison.Ordinal)) {
                    ReadHeaderAllowanceChargeTaxBasis(
                        reader,
                        ref notSubjectAllowanceAmountSum,
                        ref hasNotSubjectAllowanceAmount,
                        ref notSubjectChargeAmountSum,
                        ref hasNotSubjectChargeAmount,
                        allowanceAmountsByCategoryRate,
                        chargeAmountsByCategoryRate,
                        ref parseDiagnostic);
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

    private static void ReadHeaderAllowanceChargeTaxBasis(
        System.Xml.XmlReader reader,
        ref decimal notSubjectAllowanceAmountSum,
        ref bool hasNotSubjectAllowanceAmount,
        ref decimal notSubjectChargeAmountSum,
        ref bool hasNotSubjectChargeAmount,
        Dictionary<string, decimal> allowanceAmountsByCategoryRate,
        Dictionary<string, decimal> chargeAmountsByCategoryRate,
        ref string? parseDiagnostic) {
        if (reader.IsEmptyElement) {
            return;
        }

        bool? isCharge = null;
        decimal? actualAmount = null;
        string? categoryCode = null;
        string? taxCategoryRateKey = null;
        int depth = reader.Depth;
        while (reader.Read()) {
            if (reader.NodeType == System.Xml.XmlNodeType.Element) {
                if (reader.Depth == depth + 1 && string.Equals(reader.LocalName, "ChargeIndicator", StringComparison.Ordinal)) {
                    ReadChargeIndicator(reader, ref isCharge);
                    continue;
                }

                if (reader.Depth == depth + 1 && string.Equals(reader.LocalName, "ActualAmount", StringComparison.Ordinal)) {
                    if (TryReadAmount(reader, "SpecifiedTradeAllowanceCharge ActualAmount", ref parseDiagnostic, out decimal? amount)) {
                        actualAmount = amount;
                    }

                    continue;
                }

                if (reader.Depth == depth + 1 && string.Equals(reader.LocalName, "CategoryTradeTax", StringComparison.Ordinal)) {
                    ReadAllowanceChargeTaxCategory(reader, ref categoryCode, ref taxCategoryRateKey);
                    continue;
                }
            }

            if (reader.NodeType == System.Xml.XmlNodeType.EndElement &&
                reader.Depth == depth &&
                string.Equals(reader.LocalName, "SpecifiedTradeAllowanceCharge", StringComparison.Ordinal)) {
                break;
            }
        }

        if (!actualAmount.HasValue) {
            return;
        }

        if (!string.IsNullOrWhiteSpace(taxCategoryRateKey)) {
            if (isCharge == true) {
                AddAmount(chargeAmountsByCategoryRate, taxCategoryRateKey!, actualAmount.Value);
            } else {
                AddAmount(allowanceAmountsByCategoryRate, taxCategoryRateKey!, actualAmount.Value);
            }
        }

        if (string.Equals(categoryCode, "O", StringComparison.Ordinal)) {
            if (isCharge == true) {
                notSubjectChargeAmountSum += actualAmount.Value;
                hasNotSubjectChargeAmount = true;
            } else {
                notSubjectAllowanceAmountSum += actualAmount.Value;
                hasNotSubjectAllowanceAmount = true;
            }
        }
    }

    private static void AddAmount(Dictionary<string, decimal> amounts, string key, decimal amount) {
        decimal existing;
        if (amounts.TryGetValue(key, out existing)) {
            amounts[key] = existing + amount;
        } else {
            amounts.Add(key, amount);
        }
    }

    private static void ReadChargeIndicator(System.Xml.XmlReader reader, ref bool? isCharge) {
        if (reader.IsEmptyElement) {
            return;
        }

        int depth = reader.Depth;
        while (reader.Read()) {
            if (reader.NodeType == System.Xml.XmlNodeType.Element &&
                reader.Depth == depth + 1 &&
                string.Equals(reader.LocalName, "Indicator", StringComparison.Ordinal)) {
                string value = ReadElementText(reader).Trim();
                if (string.Equals(value, "true", StringComparison.OrdinalIgnoreCase) ||
                    string.Equals(value, "1", StringComparison.Ordinal)) {
                    isCharge = true;
                } else if (string.Equals(value, "false", StringComparison.OrdinalIgnoreCase) ||
                    string.Equals(value, "0", StringComparison.Ordinal)) {
                    isCharge = false;
                }

                continue;
            }

            if (reader.NodeType == System.Xml.XmlNodeType.Text || reader.NodeType == System.Xml.XmlNodeType.CDATA) {
                string value = reader.Value.Trim();
                if (string.Equals(value, "true", StringComparison.OrdinalIgnoreCase) ||
                    string.Equals(value, "1", StringComparison.Ordinal)) {
                    isCharge = true;
                } else if (string.Equals(value, "false", StringComparison.OrdinalIgnoreCase) ||
                    string.Equals(value, "0", StringComparison.Ordinal)) {
                    isCharge = false;
                }
            }

            if (reader.NodeType == System.Xml.XmlNodeType.EndElement &&
                reader.Depth == depth &&
                string.Equals(reader.LocalName, "ChargeIndicator", StringComparison.Ordinal)) {
                break;
            }
        }
    }

    private static void ReadAllowanceChargeTaxCategory(System.Xml.XmlReader reader, ref string? categoryCode, ref string? taxCategoryRateKey) {
        if (reader.IsEmptyElement) {
            return;
        }

        string? rate = null;
        int depth = reader.Depth;
        while (reader.Read()) {
            if (reader.NodeType == System.Xml.XmlNodeType.Element &&
                reader.Depth == depth + 1) {
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
                string.Equals(reader.LocalName, "CategoryTradeTax", StringComparison.Ordinal)) {
                break;
            }
        }

        taxCategoryRateKey = BuildTaxCategoryRateKey(categoryCode, rate);
    }

    private static void ReadApplicableTradeTaxTotals(
        System.Xml.XmlReader reader,
        ref decimal basisBreakdownSum,
        ref bool hasBasisBreakdownAmount,
        ref decimal calculatedBreakdownSum,
        ref bool hasCalculatedBreakdownAmount,
        ref decimal? notSubjectHeaderBasisAmount,
        Dictionary<string, decimal> headerBasisAmountsByCategoryRate,
        ref string? parseDiagnostic) {
        if (reader.IsEmptyElement) {
            return;
        }

        string? categoryCode = null;
        string? rate = null;
        decimal? basisAmount = null;
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

                if (string.Equals(reader.LocalName, "BasisAmount", StringComparison.Ordinal)) {
                    if (TryReadAmount(reader, "ApplicableTradeTax BasisAmount", ref parseDiagnostic, out decimal? amount)) {
                        basisAmount = amount;
                        basisBreakdownSum += amount!.Value;
                        hasBasisBreakdownAmount = true;
                    }

                    continue;
                }

                if (string.Equals(reader.LocalName, "CalculatedAmount", StringComparison.Ordinal)) {
                    if (TryReadAmount(reader, "ApplicableTradeTax CalculatedAmount", ref parseDiagnostic, out decimal? amount)) {
                        calculatedBreakdownSum += amount!.Value;
                        hasCalculatedBreakdownAmount = true;
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

        if (string.Equals(categoryCode, "O", StringComparison.Ordinal) && basisAmount.HasValue) {
            notSubjectHeaderBasisAmount = notSubjectHeaderBasisAmount.HasValue
                ? notSubjectHeaderBasisAmount.Value + basisAmount.Value
                : basisAmount.Value;
        }

        if (basisAmount.HasValue) {
            string? taxCategoryRateKey = BuildTaxCategoryRateKey(categoryCode, rate);
            if (!string.IsNullOrWhiteSpace(taxCategoryRateKey)) {
                AddAmount(headerBasisAmountsByCategoryRate, taxCategoryRateKey!, basisAmount.Value);
            }
        }
    }

    private static void ReadHeaderMonetaryTaxTotals(System.Xml.XmlReader reader, ref decimal? taxBasisTotalAmount, ref decimal? taxTotalAmount, ref string? parseDiagnostic) {
        if (reader.IsEmptyElement) {
            return;
        }

        int depth = reader.Depth;
        while (reader.Read()) {
            if (reader.NodeType == System.Xml.XmlNodeType.Element && reader.Depth == depth + 1) {
                if (string.Equals(reader.LocalName, "TaxBasisTotalAmount", StringComparison.Ordinal)) {
                    if (TryReadAmount(reader, "TaxBasisTotalAmount", ref parseDiagnostic, out decimal? amount)) {
                        taxBasisTotalAmount = amount;
                    }

                    continue;
                }

                if (string.Equals(reader.LocalName, "TaxTotalAmount", StringComparison.Ordinal)) {
                    if (TryReadAmount(reader, "TaxTotalAmount", ref parseDiagnostic, out decimal? amount)) {
                        taxTotalAmount = amount;
                    }

                    continue;
                }
            }

            if (reader.NodeType == System.Xml.XmlNodeType.EndElement &&
                reader.Depth == depth &&
                string.Equals(reader.LocalName, "SpecifiedTradeSettlementHeaderMonetarySummation", StringComparison.Ordinal)) {
                break;
            }
        }
    }

    private static void ReadLineTaxBasisContribution(
        System.Xml.XmlReader reader,
        Dictionary<string, decimal> lineNetAmountsByCategoryRate,
        ref decimal notSubjectLineNetAmountSum,
        ref bool hasNotSubjectLineNetAmount,
        ref string? parseDiagnostic) {
        if (reader.IsEmptyElement) {
            return;
        }

        string? categoryCode = null;
        string? taxCategoryRateKey = null;
        decimal? lineTotalAmount = null;
        int depth = reader.Depth;
        while (reader.Read()) {
            if (reader.NodeType == System.Xml.XmlNodeType.Element) {
                if (string.Equals(reader.LocalName, "SpecifiedLineTradeSettlement", StringComparison.Ordinal)) {
                    ReadLineSettlementTaxAndNetAmount(reader, ref categoryCode, ref taxCategoryRateKey, ref lineTotalAmount, ref parseDiagnostic);
                    continue;
                }
            }

            if (reader.NodeType == System.Xml.XmlNodeType.EndElement &&
                reader.Depth == depth &&
                string.Equals(reader.LocalName, "IncludedSupplyChainTradeLineItem", StringComparison.Ordinal)) {
                break;
            }
        }

        if (!lineTotalAmount.HasValue) {
            return;
        }

        if (!string.IsNullOrWhiteSpace(taxCategoryRateKey)) {
            AddAmount(lineNetAmountsByCategoryRate, taxCategoryRateKey!, lineTotalAmount.Value);
        }

        if (string.Equals(categoryCode, "O", StringComparison.Ordinal)) {
            notSubjectLineNetAmountSum += lineTotalAmount.Value;
            hasNotSubjectLineNetAmount = true;
        }
    }

    private static void ReadLineSettlementTaxAndNetAmount(
        System.Xml.XmlReader reader,
        ref string? categoryCode,
        ref string? taxCategoryRateKey,
        ref decimal? lineTotalAmount,
        ref string? parseDiagnostic) {
        if (reader.IsEmptyElement) {
            return;
        }

        int depth = reader.Depth;
        while (reader.Read()) {
            if (reader.NodeType == System.Xml.XmlNodeType.Element) {
                if (string.Equals(reader.LocalName, "ApplicableTradeTax", StringComparison.Ordinal)) {
                    ReadLineTaxCategory(reader, ref categoryCode, ref taxCategoryRateKey);
                    continue;
                }

                if (string.Equals(reader.LocalName, "SpecifiedTradeSettlementLineMonetarySummation", StringComparison.Ordinal)) {
                    ReadLineNetAmount(reader, ref lineTotalAmount, ref parseDiagnostic);
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

    private static void ReadLineTaxCategory(System.Xml.XmlReader reader, ref string? categoryCode, ref string? taxCategoryRateKey) {
        if (reader.IsEmptyElement) {
            return;
        }

        string? rate = null;
        int depth = reader.Depth;
        while (reader.Read()) {
            if (reader.NodeType == System.Xml.XmlNodeType.Element &&
                reader.Depth == depth + 1) {
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
                string.Equals(reader.LocalName, "ApplicableTradeTax", StringComparison.Ordinal)) {
                break;
            }
        }

        taxCategoryRateKey = BuildTaxCategoryRateKey(categoryCode, rate);
    }

    private static void ReadLineNetAmount(System.Xml.XmlReader reader, ref decimal? lineTotalAmount, ref string? parseDiagnostic) {
        if (reader.IsEmptyElement) {
            return;
        }

        int depth = reader.Depth;
        while (reader.Read()) {
            if (reader.NodeType == System.Xml.XmlNodeType.Element &&
                reader.Depth == depth + 1 &&
                string.Equals(reader.LocalName, "LineTotalAmount", StringComparison.Ordinal)) {
                if (TryReadAmount(reader, "LineTotalAmount", ref parseDiagnostic, out decimal? amount)) {
                    lineTotalAmount = amount;
                }

                continue;
            }

            if (reader.NodeType == System.Xml.XmlNodeType.EndElement &&
                reader.Depth == depth &&
                string.Equals(reader.LocalName, "SpecifiedTradeSettlementLineMonetarySummation", StringComparison.Ordinal)) {
                break;
            }
        }
    }

    private static List<string> BuildAdjustedBasisMismatches(
        Dictionary<string, decimal> headerBasisAmountsByCategoryRate,
        Dictionary<string, decimal> lineNetAmountsByCategoryRate,
        Dictionary<string, decimal> allowanceAmountsByCategoryRate,
        Dictionary<string, decimal> chargeAmountsByCategoryRate) {
        var mismatches = new List<string>();
        foreach (KeyValuePair<string, decimal> headerBasis in headerBasisAmountsByCategoryRate) {
            decimal lineNetAmount;
            lineNetAmountsByCategoryRate.TryGetValue(headerBasis.Key, out lineNetAmount);

            decimal allowanceAmount;
            allowanceAmountsByCategoryRate.TryGetValue(headerBasis.Key, out allowanceAmount);
            decimal chargeAmount;
            chargeAmountsByCategoryRate.TryGetValue(headerBasis.Key, out chargeAmount);

            decimal expectedBasis = lineNetAmount - allowanceAmount + chargeAmount;
            if (System.Math.Abs(headerBasis.Value - expectedBasis) > 0.01m) {
                mismatches.Add(
                    headerBasis.Key +
                    " expected " +
                    expectedBasis.ToString("0.00", System.Globalization.CultureInfo.InvariantCulture) +
                    " from line net " +
                    lineNetAmount.ToString("0.00", System.Globalization.CultureInfo.InvariantCulture) +
                    " minus allowances " +
                    allowanceAmount.ToString("0.00", System.Globalization.CultureInfo.InvariantCulture) +
                    " plus charges " +
                    chargeAmount.ToString("0.00", System.Globalization.CultureInfo.InvariantCulture) +
                    " but found " +
                    headerBasis.Value.ToString("0.00", System.Globalization.CultureInfo.InvariantCulture));
            }
        }

        return mismatches;
    }

    private static string? BuildTaxCategoryRateKey(string? categoryCode, string? rate) {
        if (string.IsNullOrWhiteSpace(categoryCode)) {
            return null;
        }

        string normalizedCategoryCode = categoryCode!.Trim().ToUpperInvariant();
        if (string.Equals(normalizedCategoryCode, "O", StringComparison.Ordinal)) {
            return normalizedCategoryCode;
        }

        if (string.IsNullOrWhiteSpace(rate)) {
            return null;
        }

        return normalizedCategoryCode + "/" + rate!.Trim();
    }
}
