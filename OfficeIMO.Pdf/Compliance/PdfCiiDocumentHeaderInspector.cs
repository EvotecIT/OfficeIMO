namespace OfficeIMO.Pdf;

internal static partial class PdfCiiDocumentHeaderInspector {
    internal static bool TryRead(PdfEmbeddedFile file, out PdfCiiDocumentHeaderEvidence? evidence, out string? diagnostic) {
        Guard.NotNull(file, nameof(file));
        evidence = null;

        try {
            using (var stream = new MemoryStream(file.DataSnapshot))
            using (var reader = System.Xml.XmlReader.Create(stream, new System.Xml.XmlReaderSettings {
                DtdProcessing = System.Xml.DtdProcessing.Prohibit,
                XmlResolver = null
            })) {
                bool sawRoot = false;
                bool hasTradeTransaction = false;
                string? documentId = null;
                string? typeCode = null;
                string? issueDateTime = null;

                while (reader.Read()) {
                    if (reader.NodeType == System.Xml.XmlNodeType.Element) {
                        if (!sawRoot) {
                            sawRoot = true;
                            if (!IsCiiRoot(reader)) {
                                diagnostic = "Attach UN/CEFACT CrossIndustryInvoice XML in factur-x.xml.";
                                return false;
                            }
                        }

                        if (string.Equals(reader.LocalName, "ExchangedDocument", StringComparison.Ordinal)) {
                            ReadExchangedDocument(reader, ref documentId, ref typeCode, ref issueDateTime);
                            continue;
                        }

                        if (string.Equals(reader.LocalName, "SupplyChainTradeTransaction", StringComparison.Ordinal)) {
                            hasTradeTransaction = true;
                            continue;
                        }

                    }
                }

                if (!sawRoot) {
                    diagnostic = "Attach non-empty UN/CEFACT CrossIndustryInvoice XML in factur-x.xml.";
                    return false;
                }

                evidence = new PdfCiiDocumentHeaderEvidence(documentId, typeCode, issueDateTime, hasTradeTransaction);
                diagnostic = null;
                return true;
            }
        } catch (System.Xml.XmlException ex) {
            diagnostic = "Attach parseable XML in factur-x.xml: " + ex.Message;
            return false;
        }
    }

    internal static bool TryReadTradeTransaction(PdfEmbeddedFile file, out PdfCiiTradeTransactionEvidence? evidence, out string? diagnostic) {
        Guard.NotNull(file, nameof(file));
        evidence = null;

        try {
            using (var stream = new MemoryStream(file.DataSnapshot))
            using (var reader = System.Xml.XmlReader.Create(stream, new System.Xml.XmlReaderSettings {
                DtdProcessing = System.Xml.DtdProcessing.Prohibit,
                XmlResolver = null
            })) {
                bool sawRoot = false;

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

                    if (string.Equals(reader.LocalName, "SupplyChainTradeTransaction", StringComparison.Ordinal)) {
                        evidence = ReadTradeTransaction(reader);
                        diagnostic = null;
                        return true;
                    }
                }

                if (!sawRoot) {
                    diagnostic = "Attach non-empty UN/CEFACT CrossIndustryInvoice XML in factur-x.xml.";
                    return false;
                }

                evidence = new PdfCiiTradeTransactionEvidence(false, false, false, false, false, false, false);
                diagnostic = null;
                return true;
            }
        } catch (System.Xml.XmlException ex) {
            diagnostic = "Attach parseable XML in factur-x.xml: " + ex.Message;
            return false;
        }
    }

    internal static bool TryReadLineItems(PdfEmbeddedFile file, out PdfCiiLineItemEvidence? evidence, out string? diagnostic) {
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
                bool hasLineId = true;
                bool hasProductName = true;
                bool hasBilledQuantity = true;
                bool hasBilledQuantityUnitCode = true;
                bool hasLineTotalAmount = true;
                int lineItemNumber = 0;
                var missingLineItemFields = new List<string>();

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
                        lineItemNumber++;
                        bool lineHasLineId = false;
                        bool lineHasProductName = false;
                        bool lineHasBilledQuantity = false;
                        bool lineHasBilledQuantityUnitCode = false;
                        bool lineHasLineTotalAmount = false;
                        ReadLineItem(reader, ref lineHasLineId, ref lineHasProductName, ref lineHasBilledQuantity, ref lineHasBilledQuantityUnitCode, ref lineHasLineTotalAmount);
                        hasLineId = hasLineId && lineHasLineId;
                        hasProductName = hasProductName && lineHasProductName;
                        hasBilledQuantity = hasBilledQuantity && lineHasBilledQuantity;
                        hasBilledQuantityUnitCode = hasBilledQuantityUnitCode && lineHasBilledQuantityUnitCode;
                        hasLineTotalAmount = hasLineTotalAmount && lineHasLineTotalAmount;
                        AddMissingLineItemFields(missingLineItemFields, lineItemNumber, lineHasLineId, lineHasProductName, lineHasBilledQuantity, lineHasBilledQuantityUnitCode, lineHasLineTotalAmount);
                    }
                }

                if (!sawRoot) {
                    diagnostic = "Attach non-empty UN/CEFACT CrossIndustryInvoice XML in factur-x.xml.";
                    return false;
                }

                if (!hasLineItem) {
                    hasLineId = false;
                    hasProductName = false;
                    hasBilledQuantity = false;
                    hasBilledQuantityUnitCode = false;
                    hasLineTotalAmount = false;
                }

                evidence = new PdfCiiLineItemEvidence(hasLineItem, hasLineId, hasProductName, hasBilledQuantity, hasBilledQuantityUnitCode, hasLineTotalAmount, missingLineItemFields.ToArray());
                diagnostic = null;
                return true;
            }
        } catch (System.Xml.XmlException ex) {
            diagnostic = "Attach parseable XML in factur-x.xml: " + ex.Message;
            return false;
        }
    }

    internal static bool TryReadPartyIdentification(PdfEmbeddedFile file, out PdfCiiPartyIdentificationEvidence? evidence, out string? diagnostic) {
        Guard.NotNull(file, nameof(file));
        evidence = null;

        try {
            using (var stream = new MemoryStream(file.DataSnapshot))
            using (var reader = System.Xml.XmlReader.Create(stream, new System.Xml.XmlReaderSettings {
                DtdProcessing = System.Xml.DtdProcessing.Prohibit,
                XmlResolver = null
            })) {
                bool sawRoot = false;
                bool hasSellerName = false;
                bool hasSellerCountryId = false;
                bool hasBuyerName = false;
                bool hasBuyerCountryId = false;

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
                        ReadPartyIdentification(reader, ref hasSellerName, ref hasSellerCountryId);
                        continue;
                    }

                    if (string.Equals(reader.LocalName, "BuyerTradeParty", StringComparison.Ordinal)) {
                        ReadPartyIdentification(reader, ref hasBuyerName, ref hasBuyerCountryId);
                    }
                }

                if (!sawRoot) {
                    diagnostic = "Attach non-empty UN/CEFACT CrossIndustryInvoice XML in factur-x.xml.";
                    return false;
                }

                evidence = new PdfCiiPartyIdentificationEvidence(hasSellerName, hasSellerCountryId, hasBuyerName, hasBuyerCountryId);
                diagnostic = null;
                return true;
            }
        } catch (System.Xml.XmlException ex) {
            diagnostic = "Attach parseable XML in factur-x.xml: " + ex.Message;
            return false;
        }
    }

    internal static bool TryReadSettlementSummary(PdfEmbeddedFile file, out PdfCiiSettlementSummaryEvidence? evidence, out string? diagnostic) {
        Guard.NotNull(file, nameof(file));
        evidence = null;

        try {
            using (var stream = new MemoryStream(file.DataSnapshot))
            using (var reader = System.Xml.XmlReader.Create(stream, new System.Xml.XmlReaderSettings {
                DtdProcessing = System.Xml.DtdProcessing.Prohibit,
                XmlResolver = null
            })) {
                bool sawRoot = false;
                bool hasSettlement = false;
                bool hasCurrencyCode = false;
                bool hasTradeTax = false;
                bool hasTaxBasisTotalAmount = false;
                bool hasTaxTotalAmount = false;

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

                    if (string.Equals(reader.LocalName, "ApplicableHeaderTradeSettlement", StringComparison.Ordinal)) {
                        hasSettlement = true;
                        ReadSettlementSummary(reader, ref hasCurrencyCode, ref hasTradeTax, ref hasTaxBasisTotalAmount, ref hasTaxTotalAmount);
                    }
                }

                if (!sawRoot) {
                    diagnostic = "Attach non-empty UN/CEFACT CrossIndustryInvoice XML in factur-x.xml.";
                    return false;
                }

                evidence = new PdfCiiSettlementSummaryEvidence(hasSettlement, hasCurrencyCode, hasTradeTax, hasTaxBasisTotalAmount, hasTaxTotalAmount);
                diagnostic = null;
                return true;
            }
        } catch (System.Xml.XmlException ex) {
            diagnostic = "Attach parseable XML in factur-x.xml: " + ex.Message;
            return false;
        }
    }

    internal static bool TryReadAmountConsistency(PdfEmbeddedFile file, out PdfCiiAmountConsistencyEvidence? evidence, out string? diagnostic) {
        Guard.NotNull(file, nameof(file));
        evidence = null;

        try {
            using (var stream = new MemoryStream(file.DataSnapshot))
            using (var reader = System.Xml.XmlReader.Create(stream, new System.Xml.XmlReaderSettings {
                DtdProcessing = System.Xml.DtdProcessing.Prohibit,
                XmlResolver = null
            })) {
                bool sawRoot = false;
                decimal lineTotalAmountSum = 0m;
                bool hasLineTotalAmount = false;
                decimal? allowanceTotalAmount = null;
                decimal? chargeTotalAmount = null;
                decimal documentLevelAllowanceAmountSum = 0m;
                bool hasDocumentLevelAllowanceAmount = false;
                decimal documentLevelChargeAmountSum = 0m;
                bool hasDocumentLevelChargeAmount = false;
                decimal? taxBasisTotalAmount = null;
                decimal? taxTotalAmount = null;
                decimal? grandTotalAmount = null;
                decimal? duePayableAmount = null;
                decimal? paidAmount = null;
                decimal? roundingAmount = null;
                string? parseDiagnostic = null;

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

                    if (string.Equals(reader.LocalName, "ApplicableHeaderTradeSettlement", StringComparison.Ordinal)) {
                        ReadAmountConsistencyHeaderSettlement(
                            reader,
                            ref allowanceTotalAmount,
                            ref chargeTotalAmount,
                            ref taxBasisTotalAmount,
                            ref taxTotalAmount,
                            ref grandTotalAmount,
                            ref duePayableAmount,
                            ref paidAmount,
                            ref roundingAmount,
                            ref documentLevelAllowanceAmountSum,
                            ref hasDocumentLevelAllowanceAmount,
                            ref documentLevelChargeAmountSum,
                            ref hasDocumentLevelChargeAmount,
                            ref parseDiagnostic);
                        continue;
                    }

                    if (string.Equals(reader.LocalName, "IncludedSupplyChainTradeLineItem", StringComparison.Ordinal)) {
                        ReadAmountConsistencyLineItem(
                            reader,
                            ref lineTotalAmountSum,
                            ref hasLineTotalAmount,
                            ref parseDiagnostic);
                        continue;
                    }

                    if (string.Equals(reader.LocalName, "TaxBasisTotalAmount", StringComparison.Ordinal)) {
                        if (TryReadAmount(reader, "TaxBasisTotalAmount", ref parseDiagnostic, out decimal? amount)) {
                            taxBasisTotalAmount = amount;
                        }

                        continue;
                    }

                    if (string.Equals(reader.LocalName, "AllowanceTotalAmount", StringComparison.Ordinal)) {
                        if (TryReadAmount(reader, "AllowanceTotalAmount", ref parseDiagnostic, out decimal? amount)) {
                            allowanceTotalAmount = amount;
                        }

                        continue;
                    }

                    if (string.Equals(reader.LocalName, "ChargeTotalAmount", StringComparison.Ordinal)) {
                        if (TryReadAmount(reader, "ChargeTotalAmount", ref parseDiagnostic, out decimal? amount)) {
                            chargeTotalAmount = amount;
                        }

                        continue;
                    }

                    if (string.Equals(reader.LocalName, "TaxTotalAmount", StringComparison.Ordinal)) {
                        if (TryReadAmount(reader, "TaxTotalAmount", ref parseDiagnostic, out decimal? amount)) {
                            taxTotalAmount = amount;
                        }

                        continue;
                    }

                    if (string.Equals(reader.LocalName, "GrandTotalAmount", StringComparison.Ordinal)) {
                        if (TryReadAmount(reader, "GrandTotalAmount", ref parseDiagnostic, out decimal? amount)) {
                            grandTotalAmount = amount;
                        }

                        continue;
                    }

                    if (string.Equals(reader.LocalName, "DuePayableAmount", StringComparison.Ordinal)) {
                        if (TryReadAmount(reader, "DuePayableAmount", ref parseDiagnostic, out decimal? amount)) {
                            duePayableAmount = amount;
                        }

                        continue;
                    }

                    if (string.Equals(reader.LocalName, "PaidAmount", StringComparison.Ordinal)) {
                        if (TryReadAmount(reader, "PaidAmount", ref parseDiagnostic, out decimal? amount)) {
                            paidAmount = amount;
                        }

                        continue;
                    }

                    if (string.Equals(reader.LocalName, "RoundingAmount", StringComparison.Ordinal)) {
                        if (TryReadAmount(reader, "RoundingAmount", ref parseDiagnostic, out decimal? amount)) {
                            roundingAmount = amount;
                        }
                    }
                }

                if (!sawRoot) {
                    diagnostic = "Attach non-empty UN/CEFACT CrossIndustryInvoice XML in factur-x.xml.";
                    return false;
                }

                evidence = new PdfCiiAmountConsistencyEvidence(
                    hasLineTotalAmount ? lineTotalAmountSum : (decimal?)null,
                    allowanceTotalAmount,
                    chargeTotalAmount,
                    hasDocumentLevelAllowanceAmount ? documentLevelAllowanceAmountSum : (decimal?)null,
                    hasDocumentLevelChargeAmount ? documentLevelChargeAmountSum : (decimal?)null,
                    taxBasisTotalAmount,
                    taxTotalAmount,
                    grandTotalAmount,
                    duePayableAmount,
                    paidAmount,
                    roundingAmount,
                    parseDiagnostic);
                diagnostic = null;
                return true;
            }
        } catch (System.Xml.XmlException ex) {
            diagnostic = "Attach parseable XML in factur-x.xml: " + ex.Message;
            return false;
        }
    }

    internal static bool TryReadPaymentInstructions(PdfEmbeddedFile file, out PdfCiiPaymentInstructionEvidence? evidence, out string? diagnostic) {
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
                bool hasCreditorAccount = false;
                bool hasCreditorAccountId = false;
                var typeCodes = new List<string>();
                var missingTypeCodePaymentMeans = new List<string>();
                var missingCreditorAccountPaymentMeans = new List<string>();
                var missingCreditorAccountIdPaymentMeans = new List<string>();
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
                        bool paymentMeansHasTypeCode = false;
                        bool paymentMeansHasCreditorAccount = false;
                        bool paymentMeansHasCreditorAccountId = false;
                        string? paymentMeansTypeCode = null;
                        ReadPaymentMeans(reader, ref paymentMeansHasTypeCode, typeCodes, ref paymentMeansHasCreditorAccount, ref paymentMeansHasCreditorAccountId, ref paymentMeansTypeCode);
                        hasTypeCode = hasTypeCode || paymentMeansHasTypeCode;
                        hasCreditorAccount = hasCreditorAccount || paymentMeansHasCreditorAccount;
                        hasCreditorAccountId = hasCreditorAccountId || paymentMeansHasCreditorAccountId;
                        AddMissingPaymentInstructionFields(paymentMeansIndex, paymentMeansHasTypeCode, paymentMeansTypeCode, paymentMeansHasCreditorAccount, paymentMeansHasCreditorAccountId, missingTypeCodePaymentMeans, missingCreditorAccountPaymentMeans, missingCreditorAccountIdPaymentMeans);
                    }
                }

                if (!sawRoot) {
                    diagnostic = "Attach non-empty UN/CEFACT CrossIndustryInvoice XML in factur-x.xml.";
                    return false;
                }

                evidence = new PdfCiiPaymentInstructionEvidence(
                    hasPaymentMeans,
                    hasTypeCode,
                    hasCreditorAccount,
                    hasCreditorAccountId,
                    typeCodes.Distinct(StringComparer.Ordinal).ToArray(),
                    missingTypeCodePaymentMeans,
                    missingCreditorAccountPaymentMeans,
                    missingCreditorAccountIdPaymentMeans);
                diagnostic = null;
                return true;
            }
        } catch (System.Xml.XmlException ex) {
            diagnostic = "Attach parseable XML in factur-x.xml: " + ex.Message;
            return false;
        }
    }

    private static bool IsCiiRoot(System.Xml.XmlReader reader) =>
        string.Equals(reader.LocalName, "CrossIndustryInvoice", StringComparison.Ordinal) &&
        reader.NamespaceURI.StartsWith("urn:un:unece:uncefact:data:standard:CrossIndustryInvoice:", StringComparison.Ordinal);

    private static void ReadExchangedDocument(System.Xml.XmlReader reader, ref string? documentId, ref string? typeCode, ref string? issueDateTime) {
        if (reader.IsEmptyElement) {
            return;
        }

        int depth = reader.Depth;
        while (reader.Read()) {
            if (reader.NodeType == System.Xml.XmlNodeType.Element) {
                if (string.Equals(reader.LocalName, "ID", StringComparison.Ordinal) && documentId == null) {
                    documentId = ReadElementText(reader);
                    continue;
                }

                if (string.Equals(reader.LocalName, "TypeCode", StringComparison.Ordinal) && typeCode == null) {
                    typeCode = ReadElementText(reader);
                    continue;
                }

                if (string.Equals(reader.LocalName, "IssueDateTime", StringComparison.Ordinal) && issueDateTime == null) {
                    issueDateTime = ReadElementText(reader);
                    continue;
                }
            }

            if (reader.NodeType == System.Xml.XmlNodeType.EndElement &&
                reader.Depth == depth &&
                string.Equals(reader.LocalName, "ExchangedDocument", StringComparison.Ordinal)) {
                break;
            }
        }
    }

    private static string ReadElementText(System.Xml.XmlReader reader) {
        if (reader.IsEmptyElement) {
            return string.Empty;
        }

        int depth = reader.Depth;
        var sb = new System.Text.StringBuilder();
        while (reader.Read()) {
            if (reader.NodeType == System.Xml.XmlNodeType.Text ||
                reader.NodeType == System.Xml.XmlNodeType.CDATA ||
                reader.NodeType == System.Xml.XmlNodeType.SignificantWhitespace) {
                sb.Append(reader.Value);
                continue;
            }

            if (reader.NodeType == System.Xml.XmlNodeType.EndElement && reader.Depth == depth) {
                break;
            }
        }

        return sb.ToString().Trim();
    }

    private static PdfCiiTradeTransactionEvidence ReadTradeTransaction(System.Xml.XmlReader reader) {
        bool hasAgreement = false;
        bool hasSeller = false;
        bool hasBuyer = false;
        bool hasSettlement = false;
        bool hasSummation = false;
        bool hasAmount = false;

        if (reader.IsEmptyElement) {
            return new PdfCiiTradeTransactionEvidence(true, hasAgreement, hasSeller, hasBuyer, hasSettlement, hasSummation, hasAmount);
        }

        int depth = reader.Depth;
        while (reader.Read()) {
            if (reader.NodeType == System.Xml.XmlNodeType.Element) {
                if (string.Equals(reader.LocalName, "ApplicableHeaderTradeAgreement", StringComparison.Ordinal)) {
                    hasAgreement = true;
                    continue;
                }

                if (string.Equals(reader.LocalName, "SellerTradeParty", StringComparison.Ordinal)) {
                    hasSeller = true;
                    continue;
                }

                if (string.Equals(reader.LocalName, "BuyerTradeParty", StringComparison.Ordinal)) {
                    hasBuyer = true;
                    continue;
                }

                if (string.Equals(reader.LocalName, "ApplicableHeaderTradeSettlement", StringComparison.Ordinal)) {
                    hasSettlement = true;
                    continue;
                }

                if (string.Equals(reader.LocalName, "SpecifiedTradeSettlementHeaderMonetarySummation", StringComparison.Ordinal)) {
                    hasSummation = true;
                    continue;
                }

                if (string.Equals(reader.LocalName, "GrandTotalAmount", StringComparison.Ordinal) ||
                    string.Equals(reader.LocalName, "DuePayableAmount", StringComparison.Ordinal)) {
                    hasAmount = hasAmount || !string.IsNullOrWhiteSpace(ReadElementText(reader));
                    continue;
                }
            }

            if (reader.NodeType == System.Xml.XmlNodeType.EndElement &&
                reader.Depth == depth &&
                string.Equals(reader.LocalName, "SupplyChainTradeTransaction", StringComparison.Ordinal)) {
                break;
            }
        }

        return new PdfCiiTradeTransactionEvidence(true, hasAgreement, hasSeller, hasBuyer, hasSettlement, hasSummation, hasAmount);
    }

    private static void ReadLineItem(System.Xml.XmlReader reader, ref bool hasLineId, ref bool hasProductName, ref bool hasBilledQuantity, ref bool hasBilledQuantityUnitCode, ref bool hasLineTotalAmount) {
        if (reader.IsEmptyElement) {
            return;
        }

        int depth = reader.Depth;
        while (reader.Read()) {
            if (reader.NodeType == System.Xml.XmlNodeType.Element) {
                if (string.Equals(reader.LocalName, "LineID", StringComparison.Ordinal)) {
                    hasLineId = hasLineId || !string.IsNullOrWhiteSpace(ReadElementText(reader));
                    continue;
                }

                if (string.Equals(reader.LocalName, "Name", StringComparison.Ordinal)) {
                    hasProductName = hasProductName || !string.IsNullOrWhiteSpace(ReadElementText(reader));
                    continue;
                }

                if (string.Equals(reader.LocalName, "BilledQuantity", StringComparison.Ordinal)) {
                    hasBilledQuantityUnitCode = hasBilledQuantityUnitCode || !string.IsNullOrWhiteSpace(reader.GetAttribute("unitCode"));
                    hasBilledQuantity = hasBilledQuantity || !string.IsNullOrWhiteSpace(ReadElementText(reader));
                    continue;
                }

                if (string.Equals(reader.LocalName, "LineTotalAmount", StringComparison.Ordinal)) {
                    hasLineTotalAmount = hasLineTotalAmount || !string.IsNullOrWhiteSpace(ReadElementText(reader));
                    continue;
                }
            }

            if (reader.NodeType == System.Xml.XmlNodeType.EndElement &&
                reader.Depth == depth &&
                string.Equals(reader.LocalName, "IncludedSupplyChainTradeLineItem", StringComparison.Ordinal)) {
                break;
            }
        }
    }

    private static void AddMissingLineItemFields(List<string> missingFields, int lineItemNumber, bool hasLineId, bool hasProductName, bool hasBilledQuantity, bool hasBilledQuantityUnitCode, bool hasLineTotalAmount) {
        string prefix = "line " + lineItemNumber.ToString(System.Globalization.CultureInfo.InvariantCulture) + " ";
        if (!hasLineId) {
            missingFields.Add(prefix + "AssociatedDocumentLineDocument LineID");
        }

        if (!hasProductName) {
            missingFields.Add(prefix + "SpecifiedTradeProduct Name");
        }

        if (!hasBilledQuantity) {
            missingFields.Add(prefix + "SpecifiedLineTradeDelivery BilledQuantity");
        }

        if (!hasBilledQuantityUnitCode) {
            missingFields.Add(prefix + "SpecifiedLineTradeDelivery BilledQuantity unitCode");
        }

        if (!hasLineTotalAmount) {
            missingFields.Add(prefix + "SpecifiedTradeSettlementLineMonetarySummation LineTotalAmount");
        }
    }

    private static void ReadPartyIdentification(System.Xml.XmlReader reader, ref bool hasName, ref bool hasCountryId) {
        if (reader.IsEmptyElement) {
            return;
        }

        string partyElementName = reader.LocalName;
        int depth = reader.Depth;
        while (reader.Read()) {
            if (reader.NodeType == System.Xml.XmlNodeType.Element) {
                if (reader.Depth == depth + 1 && string.Equals(reader.LocalName, "Name", StringComparison.Ordinal)) {
                    hasName = hasName || !string.IsNullOrWhiteSpace(ReadElementText(reader));
                    continue;
                }

                if (string.Equals(reader.LocalName, "PostalTradeAddress", StringComparison.Ordinal)) {
                    ReadPostalTradeAddress(reader, ref hasCountryId);
                    continue;
                }
            }

            if (reader.NodeType == System.Xml.XmlNodeType.EndElement &&
                reader.Depth == depth &&
                string.Equals(reader.LocalName, partyElementName, StringComparison.Ordinal)) {
                break;
            }
        }
    }

    private static void ReadPostalTradeAddress(System.Xml.XmlReader reader, ref bool hasCountryId) {
        if (reader.IsEmptyElement) {
            return;
        }

        int depth = reader.Depth;
        while (reader.Read()) {
            if (reader.NodeType == System.Xml.XmlNodeType.Element &&
                string.Equals(reader.LocalName, "CountryID", StringComparison.Ordinal)) {
                hasCountryId = hasCountryId || !string.IsNullOrWhiteSpace(ReadElementText(reader));
                continue;
            }

            if (reader.NodeType == System.Xml.XmlNodeType.EndElement &&
                reader.Depth == depth &&
                string.Equals(reader.LocalName, "PostalTradeAddress", StringComparison.Ordinal)) {
                break;
            }
        }
    }

    private static void ReadPaymentMeans(System.Xml.XmlReader reader, ref bool hasTypeCode, List<string> typeCodes, ref bool hasCreditorAccount, ref bool hasCreditorAccountId, ref string? paymentMeansTypeCode) {
        if (reader.IsEmptyElement) {
            return;
        }

        int depth = reader.Depth;
        while (reader.Read()) {
            if (reader.NodeType == System.Xml.XmlNodeType.Element) {
                if (reader.Depth == depth + 1 && string.Equals(reader.LocalName, "TypeCode", StringComparison.Ordinal)) {
                    string typeCode = ReadElementText(reader);
                    if (!string.IsNullOrWhiteSpace(typeCode)) {
                        hasTypeCode = true;
                        paymentMeansTypeCode = typeCode.Trim();
                        typeCodes.Add(typeCode.Trim());
                    }

                    continue;
                }

                if (string.Equals(reader.LocalName, "PayeePartyCreditorFinancialAccount", StringComparison.Ordinal)) {
                    hasCreditorAccount = true;
                    ReadCreditorFinancialAccount(reader, ref hasCreditorAccountId);
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

    private static void AddMissingPaymentInstructionFields(
        int paymentMeansIndex,
        bool hasTypeCode,
        string? paymentMeansTypeCode,
        bool hasCreditorAccount,
        bool hasCreditorAccountId,
        List<string> missingTypeCodePaymentMeans,
        List<string> missingCreditorAccountPaymentMeans,
        List<string> missingCreditorAccountIdPaymentMeans) {
        string paymentMeansLabel = "SpecifiedTradeSettlementPaymentMeans #" + paymentMeansIndex.ToString(System.Globalization.CultureInfo.InvariantCulture);
        if (!hasTypeCode) {
            missingTypeCodePaymentMeans.Add(paymentMeansLabel);
            return;
        }

        if (!PaymentMeansTypeCodeRequiresCreditorAccount(paymentMeansTypeCode)) {
            return;
        }

        if (!hasCreditorAccount) {
            missingCreditorAccountPaymentMeans.Add(paymentMeansLabel);
        }

        if (!hasCreditorAccountId) {
            missingCreditorAccountIdPaymentMeans.Add(paymentMeansLabel);
        }
    }

    private static bool PaymentMeansTypeCodeRequiresCreditorAccount(string? typeCode) {
        if (string.IsNullOrWhiteSpace(typeCode)) {
            return false;
        }

        string normalized = typeCode!.Trim();
        return string.Equals(normalized, "30", StringComparison.Ordinal) ||
               string.Equals(normalized, "31", StringComparison.Ordinal) ||
               string.Equals(normalized, "42", StringComparison.Ordinal) ||
               string.Equals(normalized, "45", StringComparison.Ordinal) ||
               string.Equals(normalized, "58", StringComparison.Ordinal) ||
               string.Equals(normalized, "59", StringComparison.Ordinal);
    }

    private static void ReadCreditorFinancialAccount(System.Xml.XmlReader reader, ref bool hasCreditorAccountId) {
        if (reader.IsEmptyElement) {
            return;
        }

        int depth = reader.Depth;
        while (reader.Read()) {
            if (reader.NodeType == System.Xml.XmlNodeType.Element &&
                (string.Equals(reader.LocalName, "IBANID", StringComparison.Ordinal) ||
                 string.Equals(reader.LocalName, "ProprietaryID", StringComparison.Ordinal))) {
                hasCreditorAccountId = hasCreditorAccountId || !string.IsNullOrWhiteSpace(ReadElementText(reader));
                continue;
            }

            if (reader.NodeType == System.Xml.XmlNodeType.EndElement &&
                reader.Depth == depth &&
                string.Equals(reader.LocalName, "PayeePartyCreditorFinancialAccount", StringComparison.Ordinal)) {
                break;
            }
        }
    }

    private static void ReadSettlementSummary(System.Xml.XmlReader reader, ref bool hasCurrencyCode, ref bool hasTradeTax, ref bool hasTaxBasisTotalAmount, ref bool hasTaxTotalAmount) {
        if (reader.IsEmptyElement) {
            return;
        }

        int depth = reader.Depth;
        while (reader.Read()) {
            if (reader.NodeType == System.Xml.XmlNodeType.Element) {
                if (string.Equals(reader.LocalName, "InvoiceCurrencyCode", StringComparison.Ordinal)) {
                    hasCurrencyCode = hasCurrencyCode || !string.IsNullOrWhiteSpace(ReadElementText(reader));
                    continue;
                }

                if (string.Equals(reader.LocalName, "ApplicableTradeTax", StringComparison.Ordinal)) {
                    hasTradeTax = true;
                    continue;
                }

                if (string.Equals(reader.LocalName, "TaxBasisTotalAmount", StringComparison.Ordinal)) {
                    hasTaxBasisTotalAmount = hasTaxBasisTotalAmount || !string.IsNullOrWhiteSpace(ReadElementText(reader));
                    continue;
                }

                if (string.Equals(reader.LocalName, "TaxTotalAmount", StringComparison.Ordinal)) {
                    hasTaxTotalAmount = hasTaxTotalAmount || !string.IsNullOrWhiteSpace(ReadElementText(reader));
                    continue;
                }
            }

            if (reader.NodeType == System.Xml.XmlNodeType.EndElement &&
                reader.Depth == depth &&
                string.Equals(reader.LocalName, "ApplicableHeaderTradeSettlement", StringComparison.Ordinal)) {
                break;
            }
        }
    }

    private static bool TryReadAmount(System.Xml.XmlReader reader, string fieldName, ref string? parseDiagnostic, out decimal? amount) {
        string text = ReadElementText(reader);
        if (TryParseCiiDecimal(text, out decimal value)) {
            amount = value;
            return true;
        }

        amount = null;
        if (parseDiagnostic == null) {
            parseDiagnostic = "Set factur-x.xml " + fieldName + " to a parseable decimal amount. Found: " + text + ".";
        }

        return false;
    }

    private static bool TryParseCiiDecimal(string? text, out decimal value) {
        string trimmed = text?.Trim() ?? string.Empty;
        if (trimmed.Length == 0 || trimmed.Contains(',')) {
            value = default;
            return false;
        }

        const System.Globalization.NumberStyles Styles =
            System.Globalization.NumberStyles.AllowLeadingSign |
            System.Globalization.NumberStyles.AllowDecimalPoint;
        return decimal.TryParse(trimmed, Styles, System.Globalization.CultureInfo.InvariantCulture, out value);
    }
}
