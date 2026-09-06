namespace OfficeIMO.Pdf;

internal static partial class PdfCiiDocumentHeaderInspector {
    internal static bool TryReadPaymentTerms(PdfEmbeddedFile file, out PdfCiiPaymentTermsEvidence? evidence, out string? diagnostic) {
        Guard.NotNull(file, nameof(file));
        evidence = null;

        try {
            using (var stream = new MemoryStream(file.DataSnapshot))
            using (var reader = System.Xml.XmlReader.Create(stream, new System.Xml.XmlReaderSettings {
                DtdProcessing = System.Xml.DtdProcessing.Prohibit,
                XmlResolver = null
            })) {
                bool sawRoot = false;
                bool hasPaymentTerms = false;
                bool hasDescription = false;
                bool hasDueDateDateTime = false;

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

                    if (string.Equals(reader.LocalName, "SpecifiedTradePaymentTerms", StringComparison.Ordinal)) {
                        hasPaymentTerms = true;
                        ReadPaymentTerms(reader, ref hasDescription, ref hasDueDateDateTime);
                    }
                }

                if (!sawRoot) {
                    diagnostic = "Attach non-empty UN/CEFACT CrossIndustryInvoice XML in factur-x.xml.";
                    return false;
                }

                evidence = new PdfCiiPaymentTermsEvidence(hasPaymentTerms, hasDescription, hasDueDateDateTime);
                diagnostic = null;
                return true;
            }
        } catch (System.Xml.XmlException ex) {
            diagnostic = "Attach parseable XML in factur-x.xml: " + ex.Message;
            return false;
        }
    }

    private static void ReadPaymentTerms(System.Xml.XmlReader reader, ref bool hasDescription, ref bool hasDueDateDateTime) {
        if (reader.IsEmptyElement) {
            return;
        }

        int depth = reader.Depth;
        while (reader.Read()) {
            if (reader.NodeType == System.Xml.XmlNodeType.Element) {
                if (reader.Depth == depth + 1 && string.Equals(reader.LocalName, "Description", StringComparison.Ordinal)) {
                    hasDescription = hasDescription || !string.IsNullOrWhiteSpace(ReadElementText(reader));
                    continue;
                }

                if (reader.Depth == depth + 1 && string.Equals(reader.LocalName, "DueDateDateTime", StringComparison.Ordinal)) {
                    hasDueDateDateTime = hasDueDateDateTime || !string.IsNullOrWhiteSpace(ReadElementText(reader));
                    continue;
                }
            }

            if (reader.NodeType == System.Xml.XmlNodeType.EndElement &&
                reader.Depth == depth &&
                string.Equals(reader.LocalName, "SpecifiedTradePaymentTerms", StringComparison.Ordinal)) {
                break;
            }
        }
    }
}
