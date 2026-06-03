namespace OfficeIMO.Pdf;

internal static partial class PdfCiiDocumentHeaderInspector {
    internal static bool TryReadPartyTaxRegistration(PdfEmbeddedFile file, out PdfCiiPartyTaxRegistrationEvidence? evidence, out string? diagnostic) {
        Guard.NotNull(file, nameof(file));
        evidence = null;

        try {
            using (var stream = new MemoryStream(file.DataSnapshot))
            using (var reader = System.Xml.XmlReader.Create(stream, new System.Xml.XmlReaderSettings {
                DtdProcessing = System.Xml.DtdProcessing.Prohibit,
                XmlResolver = null
            })) {
                bool sawRoot = false;
                bool hasSellerTaxRegistrationId = false;
                bool hasBuyerTaxRegistrationId = false;

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
                        ReadPartyTaxRegistration(reader, ref hasSellerTaxRegistrationId);
                        continue;
                    }

                    if (string.Equals(reader.LocalName, "BuyerTradeParty", StringComparison.Ordinal)) {
                        ReadPartyTaxRegistration(reader, ref hasBuyerTaxRegistrationId);
                    }
                }

                if (!sawRoot) {
                    diagnostic = "Attach non-empty UN/CEFACT CrossIndustryInvoice XML in factur-x.xml.";
                    return false;
                }

                evidence = new PdfCiiPartyTaxRegistrationEvidence(hasSellerTaxRegistrationId, hasBuyerTaxRegistrationId);
                diagnostic = null;
                return true;
            }
        } catch (System.Xml.XmlException ex) {
            diagnostic = "Attach parseable XML in factur-x.xml: " + ex.Message;
            return false;
        }
    }

    private static void ReadPartyTaxRegistration(System.Xml.XmlReader reader, ref bool hasTaxRegistrationId) {
        if (reader.IsEmptyElement) {
            return;
        }

        string partyElementName = reader.LocalName;
        int depth = reader.Depth;
        while (reader.Read()) {
            if (reader.NodeType == System.Xml.XmlNodeType.Element &&
                string.Equals(reader.LocalName, "SpecifiedTaxRegistration", StringComparison.Ordinal)) {
                ReadTaxRegistration(reader, ref hasTaxRegistrationId);
                continue;
            }

            if (reader.NodeType == System.Xml.XmlNodeType.EndElement &&
                reader.Depth == depth &&
                string.Equals(reader.LocalName, partyElementName, StringComparison.Ordinal)) {
                break;
            }
        }
    }

    private static void ReadTaxRegistration(System.Xml.XmlReader reader, ref bool hasTaxRegistrationId) {
        if (reader.IsEmptyElement) {
            return;
        }

        int depth = reader.Depth;
        while (reader.Read()) {
            if (reader.NodeType == System.Xml.XmlNodeType.Element &&
                reader.Depth == depth + 1 &&
                string.Equals(reader.LocalName, "ID", StringComparison.Ordinal)) {
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
