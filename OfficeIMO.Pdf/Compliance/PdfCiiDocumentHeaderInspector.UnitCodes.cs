namespace OfficeIMO.Pdf;

internal static partial class PdfCiiDocumentHeaderInspector {
    internal static bool TryReadUnitCodes(PdfEmbeddedFile file, out PdfCiiUnitCodeEvidence? evidence, out string? diagnostic) {
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
                bool hasBilledQuantityUnitCode = false;
                var unitCodes = new List<string>();

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
                        ReadLineUnitCodes(reader, ref hasBilledQuantity, ref hasBilledQuantityUnitCode, unitCodes);
                    }
                }

                if (!sawRoot) {
                    diagnostic = "Attach non-empty UN/CEFACT CrossIndustryInvoice XML in factur-x.xml.";
                    return false;
                }

                evidence = new PdfCiiUnitCodeEvidence(
                    hasLineItem,
                    hasBilledQuantity,
                    hasBilledQuantityUnitCode,
                    unitCodes.Distinct(StringComparer.Ordinal).ToArray());
                diagnostic = null;
                return true;
            }
        } catch (System.Xml.XmlException ex) {
            diagnostic = "Attach parseable XML in factur-x.xml: " + ex.Message;
            return false;
        }
    }

    private static void ReadLineUnitCodes(System.Xml.XmlReader reader, ref bool hasBilledQuantity, ref bool hasBilledQuantityUnitCode, List<string> unitCodes) {
        if (reader.IsEmptyElement) {
            return;
        }

        int depth = reader.Depth;
        while (reader.Read()) {
            if (reader.NodeType == System.Xml.XmlNodeType.Element &&
                string.Equals(reader.LocalName, "BilledQuantity", StringComparison.Ordinal)) {
                string? unitCode = reader.GetAttribute("unitCode");
                if (!string.IsNullOrWhiteSpace(unitCode)) {
                    hasBilledQuantityUnitCode = true;
                    unitCodes.Add(unitCode.Trim());
                }

                hasBilledQuantity = hasBilledQuantity || !string.IsNullOrWhiteSpace(ReadElementText(reader));
                continue;
            }

            if (reader.NodeType == System.Xml.XmlNodeType.EndElement &&
                reader.Depth == depth &&
                string.Equals(reader.LocalName, "IncludedSupplyChainTradeLineItem", StringComparison.Ordinal)) {
                break;
            }
        }
    }
}
