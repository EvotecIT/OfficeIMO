namespace OfficeIMO.Pdf;

internal static partial class PdfCiiDocumentHeaderInspector {
    private static void ReadAmountConsistencyLineItem(
        System.Xml.XmlReader reader,
        ref decimal lineTotalAmountSum,
        ref bool hasLineTotalAmount,
        ref string? parseDiagnostic) {
        if (reader.IsEmptyElement) {
            return;
        }

        int depth = reader.Depth;
        while (reader.Read()) {
            if (reader.NodeType == System.Xml.XmlNodeType.Element &&
                string.Equals(reader.LocalName, "LineTotalAmount", StringComparison.Ordinal)) {
                if (TryReadAmount(reader, "LineTotalAmount", ref parseDiagnostic, out decimal? amount)) {
                    lineTotalAmountSum += amount!.Value;
                    hasLineTotalAmount = true;
                }

                continue;
            }

            if (reader.NodeType == System.Xml.XmlNodeType.EndElement &&
                reader.Depth == depth &&
                string.Equals(reader.LocalName, "IncludedSupplyChainTradeLineItem", StringComparison.Ordinal)) {
                break;
            }
        }
    }

    private static void ReadAmountConsistencyHeaderSettlement(
        System.Xml.XmlReader reader,
        ref decimal? allowanceTotalAmount,
        ref decimal? chargeTotalAmount,
        ref decimal? taxBasisTotalAmount,
        ref decimal? taxTotalAmount,
        ref decimal? grandTotalAmount,
        ref decimal? duePayableAmount,
        ref decimal? paidAmount,
        ref decimal? roundingAmount,
        ref decimal documentLevelAllowanceAmountSum,
        ref bool hasDocumentLevelAllowanceAmount,
        ref decimal documentLevelChargeAmountSum,
        ref bool hasDocumentLevelChargeAmount,
        ref string? parseDiagnostic) {
        if (reader.IsEmptyElement) {
            return;
        }

        int depth = reader.Depth;
        while (reader.Read()) {
            if (reader.NodeType == System.Xml.XmlNodeType.Element) {
                if (reader.Depth == depth + 1 &&
                    string.Equals(reader.LocalName, "SpecifiedTradeAllowanceCharge", StringComparison.Ordinal)) {
                    ReadAmountConsistencyHeaderAllowanceCharge(
                        reader,
                        ref documentLevelAllowanceAmountSum,
                        ref hasDocumentLevelAllowanceAmount,
                        ref documentLevelChargeAmountSum,
                        ref hasDocumentLevelChargeAmount,
                        ref parseDiagnostic);
                    continue;
                }

                if (reader.Depth == depth + 1 &&
                    string.Equals(reader.LocalName, "SpecifiedTradeSettlementHeaderMonetarySummation", StringComparison.Ordinal)) {
                    ReadAmountConsistencyHeaderMonetarySummation(
                        reader,
                        ref allowanceTotalAmount,
                        ref chargeTotalAmount,
                        ref taxBasisTotalAmount,
                        ref taxTotalAmount,
                        ref grandTotalAmount,
                        ref duePayableAmount,
                        ref paidAmount,
                        ref roundingAmount,
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

    private static void ReadAmountConsistencyHeaderAllowanceCharge(
        System.Xml.XmlReader reader,
        ref decimal documentLevelAllowanceAmountSum,
        ref bool hasDocumentLevelAllowanceAmount,
        ref decimal documentLevelChargeAmountSum,
        ref bool hasDocumentLevelChargeAmount,
        ref string? parseDiagnostic) {
        if (reader.IsEmptyElement) {
            return;
        }

        bool? isCharge = null;
        decimal? actualAmount = null;
        int depth = reader.Depth;
        while (reader.Read()) {
            if (reader.NodeType == System.Xml.XmlNodeType.Element &&
                reader.Depth == depth + 1) {
                if (string.Equals(reader.LocalName, "ChargeIndicator", StringComparison.Ordinal)) {
                    ReadChargeIndicator(reader, ref isCharge);
                    continue;
                }

                if (string.Equals(reader.LocalName, "ActualAmount", StringComparison.Ordinal)) {
                    if (TryReadAmount(reader, "SpecifiedTradeAllowanceCharge ActualAmount", ref parseDiagnostic, out decimal? amount)) {
                        actualAmount = amount;
                    }

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

        if (isCharge == true) {
            documentLevelChargeAmountSum += actualAmount.Value;
            hasDocumentLevelChargeAmount = true;
        } else {
            documentLevelAllowanceAmountSum += actualAmount.Value;
            hasDocumentLevelAllowanceAmount = true;
        }
    }

    private static void ReadAmountConsistencyHeaderMonetarySummation(
        System.Xml.XmlReader reader,
        ref decimal? allowanceTotalAmount,
        ref decimal? chargeTotalAmount,
        ref decimal? taxBasisTotalAmount,
        ref decimal? taxTotalAmount,
        ref decimal? grandTotalAmount,
        ref decimal? duePayableAmount,
        ref decimal? paidAmount,
        ref decimal? roundingAmount,
        ref string? parseDiagnostic) {
        if (reader.IsEmptyElement) {
            return;
        }

        int depth = reader.Depth;
        while (reader.Read()) {
            if (reader.NodeType == System.Xml.XmlNodeType.Element &&
                reader.Depth == depth + 1) {
                if (string.Equals(reader.LocalName, "TaxBasisTotalAmount", StringComparison.Ordinal)) {
                    if (TryReadAmount(reader, "TaxBasisTotalAmount", ref parseDiagnostic, out decimal? amount)) {
                        taxBasisTotalAmount = amount;
                    }

                    continue;
                }

                if (string.Equals(reader.LocalName, "AllowanceTotalAmount", StringComparison.Ordinal)) {
                    if (TryReadAmount(reader, "AllowanceTotalAmount", ref parseDiagnostic, out decimal? amount)) {
                        allowanceTotalAmount = amount;
                    }

                    continue;
                }

                if (string.Equals(reader.LocalName, "ChargeTotalAmount", StringComparison.Ordinal)) {
                    if (TryReadAmount(reader, "ChargeTotalAmount", ref parseDiagnostic, out decimal? amount)) {
                        chargeTotalAmount = amount;
                    }

                    continue;
                }

                if (string.Equals(reader.LocalName, "TaxTotalAmount", StringComparison.Ordinal)) {
                    if (TryReadAmount(reader, "TaxTotalAmount", ref parseDiagnostic, out decimal? amount)) {
                        taxTotalAmount = amount;
                    }

                    continue;
                }

                if (string.Equals(reader.LocalName, "GrandTotalAmount", StringComparison.Ordinal)) {
                    if (TryReadAmount(reader, "GrandTotalAmount", ref parseDiagnostic, out decimal? amount)) {
                        grandTotalAmount = amount;
                    }

                    continue;
                }

                if (string.Equals(reader.LocalName, "DuePayableAmount", StringComparison.Ordinal)) {
                    if (TryReadAmount(reader, "DuePayableAmount", ref parseDiagnostic, out decimal? amount)) {
                        duePayableAmount = amount;
                    }

                    continue;
                }

                if (string.Equals(reader.LocalName, "PaidAmount", StringComparison.Ordinal)) {
                    if (TryReadAmount(reader, "PaidAmount", ref parseDiagnostic, out decimal? amount)) {
                        paidAmount = amount;
                    }

                    continue;
                }

                if (string.Equals(reader.LocalName, "RoundingAmount", StringComparison.Ordinal)) {
                    if (TryReadAmount(reader, "RoundingAmount", ref parseDiagnostic, out decimal? amount)) {
                        roundingAmount = amount;
                    }
                }
            }

            if (reader.NodeType == System.Xml.XmlNodeType.EndElement &&
                reader.Depth == depth &&
                string.Equals(reader.LocalName, "SpecifiedTradeSettlementHeaderMonetarySummation", StringComparison.Ordinal)) {
                break;
            }
        }
    }
}
