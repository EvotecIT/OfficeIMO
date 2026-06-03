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
                        if (!IsValidIban(value)) {
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

    private static bool IsValidIban(string value) {
        if (!TryNormalizeIban(value, out string? normalized)) {
            return false;
        }

        int modulo = 0;
        for (int i = 4; i < normalized!.Length; i++) {
            modulo = AppendIbanCharacterModulo(modulo, normalized[i]);
        }

        for (int i = 0; i < 4; i++) {
            modulo = AppendIbanCharacterModulo(modulo, normalized[i]);
        }

        return modulo == 1;
    }

    private static bool TryNormalizeIban(string value, out string? normalized) {
        var builder = new System.Text.StringBuilder(value.Length);
        for (int i = 0; i < value.Length; i++) {
            char current = value[i];
            if (char.IsWhiteSpace(current)) {
                continue;
            }

            char upper = char.ToUpperInvariant(current);
            if (!((upper >= '0' && upper <= '9') || (upper >= 'A' && upper <= 'Z'))) {
                normalized = null;
                return false;
            }

            builder.Append(upper);
        }

        if (builder.Length < 15 || builder.Length > 34) {
            normalized = null;
            return false;
        }

        normalized = builder.ToString();
        if (!(normalized[0] >= 'A' && normalized[0] <= 'Z') ||
            !(normalized[1] >= 'A' && normalized[1] <= 'Z') ||
            !(normalized[2] >= '0' && normalized[2] <= '9') ||
            !(normalized[3] >= '0' && normalized[3] <= '9')) {
            normalized = null;
            return false;
        }

        return true;
    }

    private static int AppendIbanCharacterModulo(int modulo, char value) {
        if (value >= '0' && value <= '9') {
            return AppendIbanDigitModulo(modulo, value - '0');
        }

        int letterValue = value - 'A' + 10;
        modulo = AppendIbanDigitModulo(modulo, letterValue / 10);
        return AppendIbanDigitModulo(modulo, letterValue % 10);
    }

    private static int AppendIbanDigitModulo(int modulo, int digit) => (modulo * 10 + digit) % 97;
}
