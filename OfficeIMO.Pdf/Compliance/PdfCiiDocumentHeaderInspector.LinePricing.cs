namespace OfficeIMO.Pdf;

internal static partial class PdfCiiDocumentHeaderInspector {
    internal static bool TryReadLinePricing(PdfEmbeddedFile file, out PdfCiiLinePricingEvidence? evidence, out string? diagnostic) {
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
                bool hasAgreement = false;
                bool hasProductTradePrice = false;
                bool hasPriceChargeAmount = false;

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
                        ReadLinePricing(reader, ref hasAgreement, ref hasProductTradePrice, ref hasPriceChargeAmount);
                    }
                }

                if (!sawRoot) {
                    diagnostic = "Attach non-empty UN/CEFACT CrossIndustryInvoice XML in factur-x.xml.";
                    return false;
                }

                evidence = new PdfCiiLinePricingEvidence(hasLineItem, hasAgreement, hasProductTradePrice, hasPriceChargeAmount);
                diagnostic = null;
                return true;
            }
        } catch (System.Xml.XmlException ex) {
            diagnostic = "Attach parseable XML in factur-x.xml: " + ex.Message;
            return false;
        }
    }

    private static void ReadLinePricing(System.Xml.XmlReader reader, ref bool hasAgreement, ref bool hasProductTradePrice, ref bool hasPriceChargeAmount) {
        if (reader.IsEmptyElement) {
            return;
        }

        int depth = reader.Depth;
        while (reader.Read()) {
            if (reader.NodeType == System.Xml.XmlNodeType.Element) {
                if (string.Equals(reader.LocalName, "SpecifiedLineTradeAgreement", StringComparison.Ordinal)) {
                    hasAgreement = true;
                    ReadLineTradeAgreement(reader, ref hasProductTradePrice, ref hasPriceChargeAmount);
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

    private static void ReadLineTradeAgreement(System.Xml.XmlReader reader, ref bool hasProductTradePrice, ref bool hasPriceChargeAmount) {
        if (reader.IsEmptyElement) {
            return;
        }

        int depth = reader.Depth;
        while (reader.Read()) {
            if (reader.NodeType == System.Xml.XmlNodeType.Element) {
                if (string.Equals(reader.LocalName, "NetPriceProductTradePrice", StringComparison.Ordinal)) {
                    hasProductTradePrice = true;
                    ReadProductTradePrice(reader, ref hasPriceChargeAmount);
                    continue;
                }
            }

            if (reader.NodeType == System.Xml.XmlNodeType.EndElement &&
                reader.Depth == depth &&
                string.Equals(reader.LocalName, "SpecifiedLineTradeAgreement", StringComparison.Ordinal)) {
                break;
            }
        }
    }

    private static void ReadProductTradePrice(System.Xml.XmlReader reader, ref bool hasPriceChargeAmount) {
        if (reader.IsEmptyElement) {
            return;
        }

        int depth = reader.Depth;
        while (reader.Read()) {
            if (reader.NodeType == System.Xml.XmlNodeType.Element &&
                reader.Depth == depth + 1 &&
                string.Equals(reader.LocalName, "ChargeAmount", StringComparison.Ordinal)) {
                hasPriceChargeAmount = hasPriceChargeAmount || !string.IsNullOrWhiteSpace(ReadElementText(reader));
                continue;
            }

            if (reader.NodeType == System.Xml.XmlNodeType.EndElement &&
                reader.Depth == depth &&
                string.Equals(reader.LocalName, "NetPriceProductTradePrice", StringComparison.Ordinal)) {
                break;
            }
        }
    }
}
