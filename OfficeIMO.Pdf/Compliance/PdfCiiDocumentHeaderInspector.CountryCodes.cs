namespace OfficeIMO.Pdf;

internal static partial class PdfCiiDocumentHeaderInspector {
    internal static bool TryReadCountryCodes(PdfEmbeddedFile file, out PdfCiiCountryCodeEvidence? evidence, out string? diagnostic) {
        Guard.NotNull(file, nameof(file));
        evidence = null;

        try {
            using (var stream = new MemoryStream(file.DataSnapshot))
            using (var reader = System.Xml.XmlReader.Create(stream, new System.Xml.XmlReaderSettings {
                DtdProcessing = System.Xml.DtdProcessing.Prohibit,
                XmlResolver = null
            })) {
                bool sawRoot = false;
                string? sellerCountryId = null;
                string? buyerCountryId = null;

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
                        sellerCountryId = ReadPartyCountryId(reader);
                        continue;
                    }

                    if (string.Equals(reader.LocalName, "BuyerTradeParty", StringComparison.Ordinal)) {
                        buyerCountryId = ReadPartyCountryId(reader);
                    }
                }

                if (!sawRoot) {
                    diagnostic = "Attach non-empty UN/CEFACT CrossIndustryInvoice XML in factur-x.xml.";
                    return false;
                }

                evidence = new PdfCiiCountryCodeEvidence(
                    !string.IsNullOrWhiteSpace(sellerCountryId),
                    !string.IsNullOrWhiteSpace(buyerCountryId),
                    sellerCountryId,
                    buyerCountryId);
                diagnostic = null;
                return true;
            }
        } catch (System.Xml.XmlException ex) {
            diagnostic = "Attach parseable XML in factur-x.xml: " + ex.Message;
            return false;
        }
    }

    private static string? ReadPartyCountryId(System.Xml.XmlReader reader) {
        if (reader.IsEmptyElement) {
            return null;
        }

        int depth = reader.Depth;
        while (reader.Read()) {
            if (reader.NodeType == System.Xml.XmlNodeType.Element &&
                string.Equals(reader.LocalName, "PostalTradeAddress", StringComparison.Ordinal)) {
                string? countryId = ReadPostalTradeAddressCountryId(reader);
                if (!string.IsNullOrWhiteSpace(countryId)) {
                    return countryId;
                }

                continue;
            }

            if (reader.NodeType == System.Xml.XmlNodeType.EndElement &&
                reader.Depth == depth) {
                break;
            }
        }

        return null;
    }

    private static string? ReadPostalTradeAddressCountryId(System.Xml.XmlReader reader) {
        if (reader.IsEmptyElement) {
            return null;
        }

        int depth = reader.Depth;
        while (reader.Read()) {
            if (reader.NodeType == System.Xml.XmlNodeType.Element &&
                string.Equals(reader.LocalName, "CountryID", StringComparison.Ordinal)) {
                return ReadElementText(reader);
            }

            if (reader.NodeType == System.Xml.XmlNodeType.EndElement &&
                reader.Depth == depth &&
                string.Equals(reader.LocalName, "PostalTradeAddress", StringComparison.Ordinal)) {
                break;
            }
        }

        return null;
    }
}
