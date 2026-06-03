namespace OfficeIMO.Pdf;

internal static partial class PdfCiiDocumentHeaderInspector {
    internal static bool TryReadTaxPartyIdentifiers(PdfEmbeddedFile file, out PdfCiiTaxPartyIdentifierEvidence? evidence, out string? diagnostic) {
        Guard.NotNull(file, nameof(file));
        evidence = null;

        try {
            using (var stream = new MemoryStream(file.DataSnapshot))
            using (var reader = System.Xml.XmlReader.Create(stream, new System.Xml.XmlReaderSettings {
                DtdProcessing = System.Xml.DtdProcessing.Prohibit,
                XmlResolver = null
            })) {
                bool sawRoot = false;
                bool hasApplicableTradeTax = false;
                bool hasSellerTaxRegistrationId = false;
                bool hasSellerVatRegistrationId = false;
                bool hasBuyerTaxRegistrationId = false;
                bool hasBuyerVatRegistrationId = false;
                var taxCategoryCodes = new List<string>();
                var lineTaxCategoryCodes = new List<string>();

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
                        ReadPartyTaxIdentifier(reader, ref hasSellerTaxRegistrationId, ref hasSellerVatRegistrationId);
                        continue;
                    }

                    if (string.Equals(reader.LocalName, "BuyerTradeParty", StringComparison.Ordinal)) {
                        ReadPartyTaxIdentifier(reader, ref hasBuyerTaxRegistrationId, ref hasBuyerVatRegistrationId);
                        continue;
                    }

                    if (string.Equals(reader.LocalName, "IncludedSupplyChainTradeLineItem", StringComparison.Ordinal)) {
                        ReadLineTaxIdentifierCategories(reader, lineTaxCategoryCodes);
                        continue;
                    }

                    if (string.Equals(reader.LocalName, "ApplicableHeaderTradeSettlement", StringComparison.Ordinal)) {
                        ReadHeaderTaxIdentifierCategories(reader, ref hasApplicableTradeTax, taxCategoryCodes);
                    }
                }

                if (!sawRoot) {
                    diagnostic = "Attach non-empty UN/CEFACT CrossIndustryInvoice XML in factur-x.xml.";
                    return false;
                }

                var missingSellerIdentifierCategories = new List<string>();
                var missingBuyerIdentifierCategories = new List<string>();
                var forbiddenSellerVatIdentifierCategories = new List<string>();
                var forbiddenBuyerVatIdentifierCategories = new List<string>();
                string[] distinctCategoryCodes = taxCategoryCodes.Distinct(StringComparer.Ordinal).ToArray();
                for (int i = 0; i < distinctCategoryCodes.Length; i++) {
                    string categoryCode = distinctCategoryCodes[i];
                    if (RequiresSellerTaxRegistrationIdentifier(categoryCode) && !hasSellerTaxRegistrationId) {
                        missingSellerIdentifierCategories.Add(categoryCode);
                    } else if (RequiresSellerVatRegistrationIdentifier(categoryCode) && !hasSellerVatRegistrationId) {
                        missingSellerIdentifierCategories.Add(categoryCode);
                    }

                    if (RequiresBuyerTaxRegistrationIdentifier(categoryCode) && !hasBuyerTaxRegistrationId) {
                        missingBuyerIdentifierCategories.Add(categoryCode);
                    } else if (RequiresBuyerVatRegistrationIdentifier(categoryCode) && !hasBuyerVatRegistrationId) {
                        missingBuyerIdentifierCategories.Add(categoryCode);
                    }
                }

                string[] distinctLineCategoryCodes = lineTaxCategoryCodes.Distinct(StringComparer.Ordinal).ToArray();
                for (int i = 0; i < distinctLineCategoryCodes.Length; i++) {
                    string categoryCode = distinctLineCategoryCodes[i];
                    if (!ForbidsVatRegistrationIdentifier(categoryCode)) {
                        continue;
                    }

                    if (hasSellerVatRegistrationId) {
                        forbiddenSellerVatIdentifierCategories.Add(categoryCode);
                    }

                    if (hasBuyerVatRegistrationId) {
                        forbiddenBuyerVatIdentifierCategories.Add(categoryCode);
                    }
                }

                evidence = new PdfCiiTaxPartyIdentifierEvidence(
                    hasApplicableTradeTax,
                    hasSellerTaxRegistrationId,
                    hasSellerVatRegistrationId,
                    hasBuyerTaxRegistrationId,
                    hasBuyerVatRegistrationId,
                    missingSellerIdentifierCategories.ToArray(),
                    missingBuyerIdentifierCategories.ToArray(),
                    forbiddenSellerVatIdentifierCategories.ToArray(),
                    forbiddenBuyerVatIdentifierCategories.ToArray());
                diagnostic = null;
                return true;
            }
        } catch (System.Xml.XmlException ex) {
            diagnostic = "Attach parseable XML in factur-x.xml: " + ex.Message;
            return false;
        }
    }

    private static void ReadPartyTaxIdentifier(System.Xml.XmlReader reader, ref bool hasTaxRegistrationId, ref bool hasVatRegistrationId) {
        if (reader.IsEmptyElement) {
            return;
        }

        string partyElementName = reader.LocalName;
        int depth = reader.Depth;
        while (reader.Read()) {
            if (reader.NodeType == System.Xml.XmlNodeType.Element &&
                string.Equals(reader.LocalName, "SpecifiedTaxRegistration", StringComparison.Ordinal)) {
                ReadTaxIdentifier(reader, ref hasTaxRegistrationId, ref hasVatRegistrationId);
                continue;
            }

            if (reader.NodeType == System.Xml.XmlNodeType.EndElement &&
                reader.Depth == depth &&
                string.Equals(reader.LocalName, partyElementName, StringComparison.Ordinal)) {
                break;
            }
        }
    }

    private static void ReadTaxIdentifier(System.Xml.XmlReader reader, ref bool hasTaxRegistrationId, ref bool hasVatRegistrationId) {
        if (reader.IsEmptyElement) {
            return;
        }

        int depth = reader.Depth;
        while (reader.Read()) {
            if (reader.NodeType == System.Xml.XmlNodeType.Element &&
                reader.Depth == depth + 1 &&
                string.Equals(reader.LocalName, "ID", StringComparison.Ordinal)) {
                string? schemeId = reader.GetAttribute("schemeID");
                bool isVatScheme = string.Equals(schemeId, "VA", StringComparison.OrdinalIgnoreCase);
                bool hasValue = !string.IsNullOrWhiteSpace(ReadElementText(reader));
                hasTaxRegistrationId = hasTaxRegistrationId || hasValue;
                hasVatRegistrationId = hasVatRegistrationId || (hasValue && isVatScheme);
                continue;
            }

            if (reader.NodeType == System.Xml.XmlNodeType.EndElement &&
                reader.Depth == depth &&
                string.Equals(reader.LocalName, "SpecifiedTaxRegistration", StringComparison.Ordinal)) {
                break;
            }
        }
    }

    private static void ReadHeaderTaxIdentifierCategories(System.Xml.XmlReader reader, ref bool hasApplicableTradeTax, List<string> taxCategoryCodes) {
        if (reader.IsEmptyElement) {
            return;
        }

        int depth = reader.Depth;
        while (reader.Read()) {
            if (reader.NodeType == System.Xml.XmlNodeType.Element &&
                string.Equals(reader.LocalName, "ApplicableTradeTax", StringComparison.Ordinal)) {
                hasApplicableTradeTax = true;
                ReadTaxIdentifierCategory(reader, taxCategoryCodes);
                continue;
            }

            if (reader.NodeType == System.Xml.XmlNodeType.EndElement &&
                reader.Depth == depth &&
                string.Equals(reader.LocalName, "ApplicableHeaderTradeSettlement", StringComparison.Ordinal)) {
                break;
            }
        }
    }

    private static void ReadLineTaxIdentifierCategories(System.Xml.XmlReader reader, List<string> taxCategoryCodes) {
        if (reader.IsEmptyElement) {
            return;
        }

        int depth = reader.Depth;
        while (reader.Read()) {
            if (reader.NodeType == System.Xml.XmlNodeType.Element &&
                string.Equals(reader.LocalName, "SpecifiedLineTradeSettlement", StringComparison.Ordinal)) {
                ReadLineSettlementTaxIdentifierCategories(reader, taxCategoryCodes);
                continue;
            }

            if (reader.NodeType == System.Xml.XmlNodeType.EndElement &&
                reader.Depth == depth &&
                string.Equals(reader.LocalName, "IncludedSupplyChainTradeLineItem", StringComparison.Ordinal)) {
                break;
            }
        }
    }

    private static void ReadLineSettlementTaxIdentifierCategories(System.Xml.XmlReader reader, List<string> taxCategoryCodes) {
        if (reader.IsEmptyElement) {
            return;
        }

        int depth = reader.Depth;
        while (reader.Read()) {
            if (reader.NodeType == System.Xml.XmlNodeType.Element &&
                string.Equals(reader.LocalName, "ApplicableTradeTax", StringComparison.Ordinal)) {
                ReadTaxIdentifierCategory(reader, taxCategoryCodes);
                continue;
            }

            if (reader.NodeType == System.Xml.XmlNodeType.EndElement &&
                reader.Depth == depth &&
                string.Equals(reader.LocalName, "SpecifiedLineTradeSettlement", StringComparison.Ordinal)) {
                break;
            }
        }
    }

    private static void ReadTaxIdentifierCategory(System.Xml.XmlReader reader, List<string> taxCategoryCodes) {
        if (reader.IsEmptyElement) {
            return;
        }

        int depth = reader.Depth;
        while (reader.Read()) {
            if (reader.NodeType == System.Xml.XmlNodeType.Element &&
                reader.Depth == depth + 1 &&
                string.Equals(reader.LocalName, "CategoryCode", StringComparison.Ordinal)) {
                string categoryCode = NormalizeTaxCategoryCode(ReadElementText(reader));
                if (!string.IsNullOrWhiteSpace(categoryCode)) {
                    taxCategoryCodes.Add(categoryCode);
                }

                continue;
            }

            if (reader.NodeType == System.Xml.XmlNodeType.EndElement &&
                reader.Depth == depth &&
                string.Equals(reader.LocalName, "ApplicableTradeTax", StringComparison.Ordinal)) {
                break;
            }
        }
    }

    private static bool RequiresSellerTaxRegistrationIdentifier(string categoryCode) =>
        string.Equals(categoryCode, "AE", StringComparison.Ordinal) ||
        string.Equals(categoryCode, "E", StringComparison.Ordinal);

    private static bool RequiresSellerVatRegistrationIdentifier(string categoryCode) =>
        string.Equals(categoryCode, "G", StringComparison.Ordinal) ||
        string.Equals(categoryCode, "K", StringComparison.Ordinal);

    private static bool RequiresBuyerTaxRegistrationIdentifier(string categoryCode) =>
        string.Equals(categoryCode, "AE", StringComparison.Ordinal);

    private static bool RequiresBuyerVatRegistrationIdentifier(string categoryCode) =>
        string.Equals(categoryCode, "K", StringComparison.Ordinal);

    private static bool ForbidsVatRegistrationIdentifier(string categoryCode) =>
        string.Equals(categoryCode, "O", StringComparison.Ordinal);
}
