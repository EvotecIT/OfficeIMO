namespace OfficeIMO.Pdf;

internal static partial class PdfCiiDocumentHeaderInspector {
    internal static bool TryReadLineAmountConsistency(PdfEmbeddedFile file, out PdfCiiLineAmountConsistencyEvidence? evidence, out string? diagnostic) {
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
                bool hasBilledQuantity = false;
                bool hasPriceChargeAmount = false;
                bool hasLineTotalAmount = false;
                var mismatchedLineIds = new List<string>();
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
                        hasLineItem = true;
                        ReadLineAmountConsistency(reader, ref hasBilledQuantity, ref hasPriceChargeAmount, ref hasLineTotalAmount, mismatchedLineIds, ref parseDiagnostic);
                    }
                }

                if (!sawRoot) {
                    diagnostic = "Attach non-empty UN/CEFACT CrossIndustryInvoice XML in factur-x.xml.";
                    return false;
                }

                evidence = new PdfCiiLineAmountConsistencyEvidence(
                    hasLineItem,
                    hasBilledQuantity,
                    hasPriceChargeAmount,
                    hasLineTotalAmount,
                    mismatchedLineIds.Count == 0,
                    mismatchedLineIds,
                    parseDiagnostic);
                diagnostic = null;
                return true;
            }
        } catch (System.Xml.XmlException ex) {
            diagnostic = "Attach parseable XML in factur-x.xml: " + ex.Message;
            return false;
        }
    }

    private static void ReadLineAmountConsistency(
        System.Xml.XmlReader reader,
        ref bool hasBilledQuantity,
        ref bool hasPriceChargeAmount,
        ref bool hasLineTotalAmount,
        List<string> mismatchedLineIds,
        ref string? parseDiagnostic) {
        if (reader.IsEmptyElement) {
            return;
        }

        string? lineId = null;
        decimal? quantity = null;
        decimal? price = null;
        decimal? lineTotal = null;

        int depth = reader.Depth;
        while (reader.Read()) {
            if (reader.NodeType == System.Xml.XmlNodeType.Element) {
                if (string.Equals(reader.LocalName, "AssociatedDocumentLineDocument", StringComparison.Ordinal)) {
                    ReadLineDocumentId(reader, ref lineId);
                    continue;
                }

                if (string.Equals(reader.LocalName, "SpecifiedLineTradeDelivery", StringComparison.Ordinal)) {
                    ReadLineDeliveryQuantity(reader, ref quantity, ref hasBilledQuantity, ref parseDiagnostic);
                    continue;
                }

                if (string.Equals(reader.LocalName, "SpecifiedLineTradeAgreement", StringComparison.Ordinal)) {
                    ReadLineAgreementPrice(reader, ref price, ref hasPriceChargeAmount, ref parseDiagnostic);
                    continue;
                }

                if (string.Equals(reader.LocalName, "SpecifiedLineTradeSettlement", StringComparison.Ordinal)) {
                    ReadLineSettlementAmount(reader, ref lineTotal, ref hasLineTotalAmount, ref parseDiagnostic);
                    continue;
                }
            }

            if (reader.NodeType == System.Xml.XmlNodeType.EndElement &&
                reader.Depth == depth &&
                string.Equals(reader.LocalName, "IncludedSupplyChainTradeLineItem", StringComparison.Ordinal)) {
                break;
            }
        }

        if (quantity.HasValue && price.HasValue && lineTotal.HasValue &&
            System.Math.Abs(quantity.Value * price.Value - lineTotal.Value) > 0.01m) {
            mismatchedLineIds.Add(string.IsNullOrWhiteSpace(lineId) ? "(unknown)" : lineId!);
        }
    }

    private static void ReadLineDocumentId(System.Xml.XmlReader reader, ref string? lineId) {
        if (reader.IsEmptyElement) {
            return;
        }

        int depth = reader.Depth;
        while (reader.Read()) {
            if (reader.NodeType == System.Xml.XmlNodeType.Element &&
                reader.Depth == depth + 1 &&
                string.Equals(reader.LocalName, "LineID", StringComparison.Ordinal)) {
                string value = ReadElementText(reader);
                if (!string.IsNullOrWhiteSpace(value)) {
                    lineId = value.Trim();
                }

                continue;
            }

            if (reader.NodeType == System.Xml.XmlNodeType.EndElement &&
                reader.Depth == depth &&
                string.Equals(reader.LocalName, "AssociatedDocumentLineDocument", StringComparison.Ordinal)) {
                break;
            }
        }
    }

    private static void ReadLineDeliveryQuantity(System.Xml.XmlReader reader, ref decimal? quantity, ref bool hasBilledQuantity, ref string? parseDiagnostic) {
        if (reader.IsEmptyElement) {
            return;
        }

        int depth = reader.Depth;
        while (reader.Read()) {
            if (reader.NodeType == System.Xml.XmlNodeType.Element &&
                reader.Depth == depth + 1 &&
                string.Equals(reader.LocalName, "BilledQuantity", StringComparison.Ordinal)) {
                if (TryReadAmount(reader, "BilledQuantity", ref parseDiagnostic, out decimal? amount)) {
                    quantity = amount;
                    hasBilledQuantity = true;
                }

                continue;
            }

            if (reader.NodeType == System.Xml.XmlNodeType.EndElement &&
                reader.Depth == depth &&
                string.Equals(reader.LocalName, "SpecifiedLineTradeDelivery", StringComparison.Ordinal)) {
                break;
            }
        }
    }

    private static void ReadLineAgreementPrice(System.Xml.XmlReader reader, ref decimal? price, ref bool hasPriceChargeAmount, ref string? parseDiagnostic) {
        if (reader.IsEmptyElement) {
            return;
        }

        int depth = reader.Depth;
        while (reader.Read()) {
            if (reader.NodeType == System.Xml.XmlNodeType.Element &&
                (string.Equals(reader.LocalName, "GrossPriceProductTradePrice", StringComparison.Ordinal) ||
                 string.Equals(reader.LocalName, "NetPriceProductTradePrice", StringComparison.Ordinal))) {
                ReadProductTradePriceAmount(reader, ref price, ref hasPriceChargeAmount, ref parseDiagnostic);
                continue;
            }

            if (reader.NodeType == System.Xml.XmlNodeType.EndElement &&
                reader.Depth == depth &&
                string.Equals(reader.LocalName, "SpecifiedLineTradeAgreement", StringComparison.Ordinal)) {
                break;
            }
        }
    }

    private static void ReadProductTradePriceAmount(System.Xml.XmlReader reader, ref decimal? price, ref bool hasPriceChargeAmount, ref string? parseDiagnostic) {
        if (reader.IsEmptyElement) {
            return;
        }

        int depth = reader.Depth;
        while (reader.Read()) {
            if (reader.NodeType == System.Xml.XmlNodeType.Element &&
                reader.Depth == depth + 1 &&
                string.Equals(reader.LocalName, "ChargeAmount", StringComparison.Ordinal)) {
                if (TryReadAmount(reader, "ProductTradePrice ChargeAmount", ref parseDiagnostic, out decimal? amount)) {
                    price = amount;
                    hasPriceChargeAmount = true;
                }

                continue;
            }

            if (reader.NodeType == System.Xml.XmlNodeType.EndElement &&
                reader.Depth == depth &&
                (string.Equals(reader.LocalName, "GrossPriceProductTradePrice", StringComparison.Ordinal) ||
                 string.Equals(reader.LocalName, "NetPriceProductTradePrice", StringComparison.Ordinal))) {
                break;
            }
        }
    }

    private static void ReadLineSettlementAmount(System.Xml.XmlReader reader, ref decimal? lineTotal, ref bool hasLineTotalAmount, ref string? parseDiagnostic) {
        if (reader.IsEmptyElement) {
            return;
        }

        int depth = reader.Depth;
        while (reader.Read()) {
            if (reader.NodeType == System.Xml.XmlNodeType.Element &&
                string.Equals(reader.LocalName, "LineTotalAmount", StringComparison.Ordinal)) {
                if (TryReadAmount(reader, "LineTotalAmount", ref parseDiagnostic, out decimal? amount)) {
                    lineTotal = amount;
                    hasLineTotalAmount = true;
                }

                continue;
            }

            if (reader.NodeType == System.Xml.XmlNodeType.EndElement &&
                reader.Depth == depth &&
                string.Equals(reader.LocalName, "SpecifiedLineTradeSettlement", StringComparison.Ordinal)) {
                break;
            }
        }
    }
}
