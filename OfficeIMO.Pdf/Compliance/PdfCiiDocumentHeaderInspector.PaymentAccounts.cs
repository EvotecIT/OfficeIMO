namespace OfficeIMO.Pdf;

internal static partial class PdfCiiDocumentHeaderInspector {
    internal static bool TryReadPaymentAccounts(PdfEmbeddedFile file, out PdfCiiPaymentAccountEvidence? evidence, out string? diagnostic) {
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
                bool hasCreditorAccount = false;
                bool hasAccountId = false;
                bool hasIbanId = false;
                var invalidIbanIds = new List<string>();
                var typeCodes = new List<string>();

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
                        ReadPaymentAccountMeans(reader, typeCodes, ref hasCreditorAccount, ref hasAccountId, ref hasIbanId, invalidIbanIds);
                    }
                }

                if (!sawRoot) {
                    diagnostic = "Attach non-empty UN/CEFACT CrossIndustryInvoice XML in factur-x.xml.";
                    return false;
                }

                evidence = new PdfCiiPaymentAccountEvidence(
                    hasPaymentMeans,
                    hasCreditorAccount,
                    hasAccountId,
                    hasIbanId,
                    invalidIbanIds.Count == 0,
                    invalidIbanIds.Distinct(StringComparer.Ordinal).ToArray(),
                    typeCodes.Distinct(StringComparer.Ordinal).ToArray());
                diagnostic = null;
                return true;
            }
        } catch (System.Xml.XmlException ex) {
            diagnostic = "Attach parseable XML in factur-x.xml: " + ex.Message;
            return false;
        }
    }

    private static void ReadPaymentAccountMeans(System.Xml.XmlReader reader, List<string> typeCodes, ref bool hasCreditorAccount, ref bool hasAccountId, ref bool hasIbanId, List<string> invalidIbanIds) {
        if (reader.IsEmptyElement) {
            return;
        }

        int depth = reader.Depth;
        while (reader.Read()) {
            if (reader.NodeType == System.Xml.XmlNodeType.Element) {
                if (reader.Depth == depth + 1 && string.Equals(reader.LocalName, "TypeCode", StringComparison.Ordinal)) {
                    string typeCode = ReadElementText(reader);
                    if (!string.IsNullOrWhiteSpace(typeCode)) {
                        typeCodes.Add(typeCode.Trim());
                    }

                    continue;
                }

                if (string.Equals(reader.LocalName, "PayeePartyCreditorFinancialAccount", StringComparison.Ordinal)) {
                    hasCreditorAccount = true;
                    ReadPaymentAccountValues(reader, ref hasAccountId, ref hasIbanId, invalidIbanIds);
                    continue;
                }
            }

            if (reader.NodeType == System.Xml.XmlNodeType.EndElement &&
                reader.Depth == depth &&
                string.Equals(reader.LocalName, "SpecifiedTradeSettlementPaymentMeans", StringComparison.Ordinal)) {
                break;
            }
        }
    }

    private static void ReadPaymentAccountValues(System.Xml.XmlReader reader, ref bool hasAccountId, ref bool hasIbanId, List<string> invalidIbanIds) {
        if (reader.IsEmptyElement) {
            return;
        }

        int depth = reader.Depth;
        while (reader.Read()) {
            if (reader.NodeType == System.Xml.XmlNodeType.Element) {
                if (string.Equals(reader.LocalName, "IBANID", StringComparison.Ordinal)) {
                    string value = ReadElementText(reader);
                    if (!string.IsNullOrWhiteSpace(value)) {
                        hasAccountId = true;
                        hasIbanId = true;
                        if (!OfficeIMO.Internal.Invoicing.InvoiceBankAccountIdentity.IsValidIban(value)) {
                            invalidIbanIds.Add(value.Trim());
                        }
                    }

                    continue;
                }

                if (string.Equals(reader.LocalName, "ProprietaryID", StringComparison.Ordinal)) {
                    hasAccountId = hasAccountId || !string.IsNullOrWhiteSpace(ReadElementText(reader));
                    continue;
                }
            }

            if (reader.NodeType == System.Xml.XmlNodeType.EndElement &&
                reader.Depth == depth &&
                string.Equals(reader.LocalName, "PayeePartyCreditorFinancialAccount", StringComparison.Ordinal)) {
                break;
            }
        }
    }

}
