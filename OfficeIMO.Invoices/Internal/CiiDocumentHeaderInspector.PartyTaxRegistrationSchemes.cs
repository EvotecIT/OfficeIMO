namespace OfficeIMO.Invoices.Internal;

internal static partial class CiiDocumentHeaderInspector {
    internal static bool TryReadPartyTaxRegistrationSchemes(byte[] data, out CiiPartyTaxRegistrationSchemeEvidence? evidence, out string? diagnostic) {
        if (data == null) throw new ArgumentNullException(nameof(data), "Invoice XML is required.");
        evidence = null;

        try {
            using (var stream = new MemoryStream(data))
            using (var reader = System.Xml.XmlReader.Create(stream, new System.Xml.XmlReaderSettings {
                DtdProcessing = System.Xml.DtdProcessing.Prohibit,
                XmlResolver = null
            })) {
                bool sawRoot = false;
                bool hasSellerTaxRegistrationId = false;
                bool hasSellerTaxRegistrationSchemeId = false;
                bool hasBuyerTaxRegistrationId = false;
                bool hasBuyerTaxRegistrationSchemeId = false;

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
                        ReadPartyTaxRegistrationScheme(reader, ref hasSellerTaxRegistrationId, ref hasSellerTaxRegistrationSchemeId);
                        continue;
                    }

                    if (string.Equals(reader.LocalName, "BuyerTradeParty", StringComparison.Ordinal)) {
                        ReadPartyTaxRegistrationScheme(reader, ref hasBuyerTaxRegistrationId, ref hasBuyerTaxRegistrationSchemeId);
                    }
                }

                if (!sawRoot) {
                    diagnostic = "Attach non-empty UN/CEFACT CrossIndustryInvoice XML in factur-x.xml.";
                    return false;
                }

                evidence = new CiiPartyTaxRegistrationSchemeEvidence(
                    hasSellerTaxRegistrationId,
                    hasSellerTaxRegistrationSchemeId,
                    hasBuyerTaxRegistrationId,
                    hasBuyerTaxRegistrationSchemeId);
                diagnostic = null;
                return true;
            }
        } catch (System.Xml.XmlException ex) {
            diagnostic = "Attach parseable XML in factur-x.xml: " + ex.Message;
            return false;
        }
    }

    private static void ReadPartyTaxRegistrationScheme(System.Xml.XmlReader reader, ref bool hasTaxRegistrationId, ref bool hasTaxRegistrationSchemeId) {
        if (reader.IsEmptyElement) {
            return;
        }

        string partyElementName = reader.LocalName;
        int depth = reader.Depth;
        while (reader.Read()) {
            if (reader.NodeType == System.Xml.XmlNodeType.Element &&
                string.Equals(reader.LocalName, "SpecifiedTaxRegistration", StringComparison.Ordinal)) {
                ReadTaxRegistrationScheme(reader, ref hasTaxRegistrationId, ref hasTaxRegistrationSchemeId);
                continue;
            }

            if (reader.NodeType == System.Xml.XmlNodeType.EndElement &&
                reader.Depth == depth &&
                string.Equals(reader.LocalName, partyElementName, StringComparison.Ordinal)) {
                break;
            }
        }
    }

    private static void ReadTaxRegistrationScheme(System.Xml.XmlReader reader, ref bool hasTaxRegistrationId, ref bool hasTaxRegistrationSchemeId) {
        if (reader.IsEmptyElement) {
            return;
        }

        int depth = reader.Depth;
        while (reader.Read()) {
            if (reader.NodeType == System.Xml.XmlNodeType.Element &&
                reader.Depth == depth + 1 &&
                string.Equals(reader.LocalName, "ID", StringComparison.Ordinal)) {
                hasTaxRegistrationSchemeId = hasTaxRegistrationSchemeId || !string.IsNullOrWhiteSpace(reader.GetAttribute("schemeID"));
                hasTaxRegistrationId = hasTaxRegistrationId || !string.IsNullOrWhiteSpace(ReadElementText(reader));
                continue;
            }

            if (reader.NodeType == System.Xml.XmlNodeType.EndElement &&
                reader.Depth == depth &&
                string.Equals(reader.LocalName, "SpecifiedTaxRegistration", StringComparison.Ordinal)) {
                break;
            }
        }
    }
}
