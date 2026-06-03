namespace OfficeIMO.Pdf;

internal static partial class PdfCiiDocumentHeaderInspector {
    internal static bool TryReadElectronicAddresses(PdfEmbeddedFile file, out PdfCiiElectronicAddressEvidence? evidence, out string? diagnostic) {
        Guard.NotNull(file, nameof(file));
        evidence = null;

        try {
            using (var stream = new MemoryStream(file.DataSnapshot))
            using (var reader = System.Xml.XmlReader.Create(stream, new System.Xml.XmlReaderSettings {
                DtdProcessing = System.Xml.DtdProcessing.Prohibit,
                XmlResolver = null
            })) {
                bool sawRoot = false;
                bool hasSellerCommunication = false;
                bool hasSellerUriId = false;
                bool hasSellerSchemeId = false;
                bool hasBuyerCommunication = false;
                bool hasBuyerUriId = false;
                bool hasBuyerSchemeId = false;
                var sellerSchemeIds = new List<string>();
                var buyerSchemeIds = new List<string>();

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

                    if (string.Equals(reader.LocalName, "SellerTradeParty", StringComparison.Ordinal)) {
                        ReadPartyElectronicAddress(reader, ref hasSellerCommunication, ref hasSellerUriId, ref hasSellerSchemeId, sellerSchemeIds);
                        continue;
                    }

                    if (string.Equals(reader.LocalName, "BuyerTradeParty", StringComparison.Ordinal)) {
                        ReadPartyElectronicAddress(reader, ref hasBuyerCommunication, ref hasBuyerUriId, ref hasBuyerSchemeId, buyerSchemeIds);
                    }
                }

                if (!sawRoot) {
                    diagnostic = "Attach non-empty UN/CEFACT CrossIndustryInvoice XML in factur-x.xml.";
                    return false;
                }

                evidence = new PdfCiiElectronicAddressEvidence(
                    hasSellerCommunication,
                    hasSellerUriId,
                    hasSellerSchemeId,
                    sellerSchemeIds.Distinct(StringComparer.Ordinal).ToArray(),
                    hasBuyerCommunication,
                    hasBuyerUriId,
                    hasBuyerSchemeId,
                    buyerSchemeIds.Distinct(StringComparer.Ordinal).ToArray());
                diagnostic = null;
                return true;
            }
        } catch (System.Xml.XmlException ex) {
            diagnostic = "Attach parseable XML in factur-x.xml: " + ex.Message;
            return false;
        }
    }

    private static void ReadPartyElectronicAddress(
        System.Xml.XmlReader reader,
        ref bool hasCommunication,
        ref bool hasUriId,
        ref bool hasSchemeId,
        List<string> schemeIds) {
        if (reader.IsEmptyElement) {
            return;
        }

        string partyElementName = reader.LocalName;
        int depth = reader.Depth;
        while (reader.Read()) {
            if (reader.NodeType == System.Xml.XmlNodeType.Element &&
                string.Equals(reader.LocalName, "URIUniversalCommunication", StringComparison.Ordinal)) {
                hasCommunication = true;
                ReadUriUniversalCommunication(reader, ref hasUriId, ref hasSchemeId, schemeIds);
                continue;
            }

            if (reader.NodeType == System.Xml.XmlNodeType.EndElement &&
                reader.Depth == depth &&
                string.Equals(reader.LocalName, partyElementName, StringComparison.Ordinal)) {
                break;
            }
        }
    }

    private static void ReadUriUniversalCommunication(System.Xml.XmlReader reader, ref bool hasUriId, ref bool hasSchemeId, List<string> schemeIds) {
        if (reader.IsEmptyElement) {
            return;
        }

        int depth = reader.Depth;
        while (reader.Read()) {
            if (reader.NodeType == System.Xml.XmlNodeType.Element &&
                reader.Depth == depth + 1 &&
                string.Equals(reader.LocalName, "URIID", StringComparison.Ordinal)) {
                string? schemeId = reader.GetAttribute("schemeID");
                if (!string.IsNullOrWhiteSpace(schemeId)) {
                    hasSchemeId = true;
                    schemeIds.Add(schemeId.Trim());
                }

                hasUriId = hasUriId || !string.IsNullOrWhiteSpace(ReadElementText(reader));
                continue;
            }

            if (reader.NodeType == System.Xml.XmlNodeType.EndElement &&
                reader.Depth == depth &&
                string.Equals(reader.LocalName, "URIUniversalCommunication", StringComparison.Ordinal)) {
                break;
            }
        }
    }
}
