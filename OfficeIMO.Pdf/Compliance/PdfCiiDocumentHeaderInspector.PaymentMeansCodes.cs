namespace OfficeIMO.Pdf;

internal static partial class PdfCiiDocumentHeaderInspector {
    internal static bool TryReadPaymentMeansCodes(PdfEmbeddedFile file, out PdfCiiPaymentMeansCodeEvidence? evidence, out string? diagnostic) {
        Guard.NotNull(file, nameof(file));
        evidence = null;

        try {
            using (var stream = new MemoryStream(file.DataSnapshot))
            using (var reader = System.Xml.XmlReader.Create(stream, new System.Xml.XmlReaderSettings {
                DtdProcessing = System.Xml.DtdProcessing.Prohibit,
                XmlResolver = null
            })) {
                bool sawRoot = false;
                bool hasPaymentMeans = false;
                bool hasTypeCode = false;
                var typeCodes = new List<string>();
                var missingTypeCodePaymentMeans = new List<string>();
                int paymentMeansIndex = 0;

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

                    if (string.Equals(reader.LocalName, "SpecifiedTradeSettlementPaymentMeans", StringComparison.Ordinal)) {
                        hasPaymentMeans = true;
                        paymentMeansIndex++;
                        bool paymentMeansHasTypeCode = ReadPaymentMeansTypeCodes(reader, typeCodes);
                        hasTypeCode = hasTypeCode || paymentMeansHasTypeCode;
                        if (!paymentMeansHasTypeCode) {
                            missingTypeCodePaymentMeans.Add("SpecifiedTradeSettlementPaymentMeans #" + paymentMeansIndex.ToString(System.Globalization.CultureInfo.InvariantCulture));
                        }
                    }
                }

                if (!sawRoot) {
                    diagnostic = "Attach non-empty UN/CEFACT CrossIndustryInvoice XML in factur-x.xml.";
                    return false;
                }

                evidence = new PdfCiiPaymentMeansCodeEvidence(
                    hasPaymentMeans,
                    hasTypeCode,
                    typeCodes.Distinct(StringComparer.Ordinal).ToArray(),
                    missingTypeCodePaymentMeans);
                diagnostic = null;
                return true;
            }
        } catch (System.Xml.XmlException ex) {
            diagnostic = "Attach parseable XML in factur-x.xml: " + ex.Message;
            return false;
        }
    }

    private static bool ReadPaymentMeansTypeCodes(System.Xml.XmlReader reader, List<string> typeCodes) {
        if (reader.IsEmptyElement) {
            return false;
        }

        bool hasTypeCode = false;
        int depth = reader.Depth;
        while (reader.Read()) {
            if (reader.NodeType == System.Xml.XmlNodeType.Element &&
                reader.Depth == depth + 1 &&
                string.Equals(reader.LocalName, "TypeCode", StringComparison.Ordinal)) {
                string value = ReadElementText(reader);
                if (!string.IsNullOrWhiteSpace(value)) {
                    hasTypeCode = true;
                    typeCodes.Add(value.Trim());
                }

                continue;
            }

            if (reader.NodeType == System.Xml.XmlNodeType.EndElement &&
                reader.Depth == depth &&
                string.Equals(reader.LocalName, "SpecifiedTradeSettlementPaymentMeans", StringComparison.Ordinal)) {
                break;
            }
        }

        return hasTypeCode;
    }
}
