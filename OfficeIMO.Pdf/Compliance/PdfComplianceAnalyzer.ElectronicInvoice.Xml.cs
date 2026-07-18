namespace OfficeIMO.Pdf;

internal static partial class PdfComplianceAnalyzer {

    private static PdfComplianceRequirement BuildElectronicInvoiceXmlAttachmentRequirement(PdfOptions options) {
        IReadOnlyList<PdfEmbeddedFile> embeddedFiles = options.EmbeddedFiles;
        var diagnostics = new List<string>();
        for (int i = 0; i < embeddedFiles.Count; i++) {
            PdfEmbeddedFile file = embeddedFiles[i];
            if (IsFacturXCiiAttachment(file, diagnostics)) {
                return new PdfComplianceRequirement(
                    "einvoice-xml-attachment",
                    "EN 16931 XML associated file",
                    PdfComplianceRequirementStatus.Satisfied,
                    "A canonical factur-x.xml associated file contains parseable UN/CEFACT CrossIndustryInvoice XML.");
            }
        }

        string diagnostic = diagnostics.Count == 0
            ? "Attach the EN 16931 CII XML payload as factur-x.xml with an associated-file relationship and an XML MIME type."
            : string.Join(" ", diagnostics.Distinct(StringComparer.Ordinal).ToArray());
        return new PdfComplianceRequirement(
            "einvoice-xml-attachment",
            "EN 16931 XML associated file",
            PdfComplianceRequirementStatus.Missing,
            diagnostic);
    }

    private static PdfComplianceRequirement BuildElectronicInvoiceXmlProfileContextRequirement(PdfOptions options) {
        IReadOnlyList<PdfEmbeddedFile> embeddedFiles = options.EmbeddedFiles;
        var diagnostics = new List<string>();
        for (int i = 0; i < embeddedFiles.Count; i++) {
            PdfEmbeddedFile file = embeddedFiles[i];
            if (!IsFacturXCiiAttachment(file, diagnostics)) {
                continue;
            }

            if (!TryReadCiiProfileContext(file, out string? contextId, out string? profileDiagnostic)) {
                return new PdfComplianceRequirement(
                    "einvoice-xml-profile-context",
                    "EN 16931 XML profile context",
                    PdfComplianceRequirementStatus.Missing,
                    profileDiagnostic!);
            }

            if (!IsKnownElectronicInvoiceProfileContext(contextId!)) {
                return new PdfComplianceRequirement(
                    "einvoice-xml-profile-context",
                    "EN 16931 XML profile context",
                    PdfComplianceRequirementStatus.Missing,
                    "Set factur-x.xml ExchangedDocumentContext GuidelineSpecifiedDocumentContextParameter ID to a recognized Factur-X, ZUGFeRD, EN 16931, or XRechnung profile identifier. Found: " + contextId + ".");
            }

            return new PdfComplianceRequirement(
                "einvoice-xml-profile-context",
                "EN 16931 XML profile context",
                PdfComplianceRequirementStatus.Satisfied,
                "The factur-x.xml CrossIndustryInvoice declares a recognized EN 16931 profile context identifier.");
        }

        string diagnostic = diagnostics.Count == 0
            ? "Attach a canonical factur-x.xml CrossIndustryInvoice file before checking EN 16931 XML profile context."
            : string.Join(" ", diagnostics.Distinct(StringComparer.Ordinal).ToArray());
        return new PdfComplianceRequirement(
            "einvoice-xml-profile-context",
            "EN 16931 XML profile context",
            PdfComplianceRequirementStatus.Missing,
            diagnostic);
    }

    private static PdfComplianceRequirement BuildElectronicInvoiceXmlDocumentHeaderRequirement(PdfOptions options) {
        IReadOnlyList<PdfEmbeddedFile> embeddedFiles = options.EmbeddedFiles;
        var diagnostics = new List<string>();
        for (int i = 0; i < embeddedFiles.Count; i++) {
            PdfEmbeddedFile file = embeddedFiles[i];
            if (!IsFacturXCiiAttachment(file, diagnostics)) {
                continue;
            }

            if (!PdfCiiDocumentHeaderInspector.TryRead(file, out PdfCiiDocumentHeaderEvidence? evidence, out string? documentDiagnostic)) {
                return new PdfComplianceRequirement(
                    "einvoice-xml-document-header",
                    "EN 16931 XML document header",
                    PdfComplianceRequirementStatus.Missing,
                    documentDiagnostic!);
            }

            var missingFields = new List<string>();
            if (string.IsNullOrWhiteSpace(evidence!.Id)) {
                missingFields.Add("ExchangedDocument ID");
            }

            if (string.IsNullOrWhiteSpace(evidence.TypeCode)) {
                missingFields.Add("ExchangedDocument TypeCode");
            }

            if (string.IsNullOrWhiteSpace(evidence.IssueDateTime)) {
                missingFields.Add("ExchangedDocument IssueDateTime");
            }

            if (!evidence.HasSupplyChainTradeTransaction) {
                missingFields.Add("SupplyChainTradeTransaction");
            }

            if (missingFields.Count > 0) {
                return new PdfComplianceRequirement(
                    "einvoice-xml-document-header",
                    "EN 16931 XML document header",
                    PdfComplianceRequirementStatus.Missing,
                    "Set factur-x.xml document header essentials before Mustang validation: " + string.Join(", ", missingFields.ToArray()) + ".");
            }

            return new PdfComplianceRequirement(
                "einvoice-xml-document-header",
                "EN 16931 XML document header",
                PdfComplianceRequirementStatus.Satisfied,
                "The factur-x.xml CrossIndustryInvoice includes document ID, type code, issue date/time, and supply-chain transaction content for e-invoice readiness.");
        }

        string diagnostic = diagnostics.Count == 0
            ? "Attach a canonical factur-x.xml CrossIndustryInvoice file before checking EN 16931 XML document header essentials."
            : string.Join(" ", diagnostics.Distinct(StringComparer.Ordinal).ToArray());
        return new PdfComplianceRequirement(
            "einvoice-xml-document-header",
            "EN 16931 XML document header",
            PdfComplianceRequirementStatus.Missing,
            diagnostic);
    }

    private static PdfComplianceRequirement BuildElectronicInvoiceXmlTradeTransactionRequirement(PdfOptions options) {
        IReadOnlyList<PdfEmbeddedFile> embeddedFiles = options.EmbeddedFiles;
        var diagnostics = new List<string>();
        for (int i = 0; i < embeddedFiles.Count; i++) {
            PdfEmbeddedFile file = embeddedFiles[i];
            if (!IsFacturXCiiAttachment(file, diagnostics)) {
                continue;
            }

            if (!PdfCiiDocumentHeaderInspector.TryReadTradeTransaction(file, out PdfCiiTradeTransactionEvidence? evidence, out string? tradeDiagnostic)) {
                return new PdfComplianceRequirement(
                    "einvoice-xml-trade-transaction",
                    "EN 16931 XML trade transaction",
                    PdfComplianceRequirementStatus.Missing,
                    tradeDiagnostic!);
            }

            var missingFields = new List<string>();
            if (!evidence!.HasSupplyChainTradeTransaction) {
                missingFields.Add("SupplyChainTradeTransaction");
            }

            if (!evidence.HasApplicableHeaderTradeAgreement) {
                missingFields.Add("ApplicableHeaderTradeAgreement");
            }

            if (!evidence.HasSellerTradeParty) {
                missingFields.Add("SellerTradeParty");
            }

            if (!evidence.HasBuyerTradeParty) {
                missingFields.Add("BuyerTradeParty");
            }

            if (!evidence.HasApplicableHeaderTradeSettlement) {
                missingFields.Add("ApplicableHeaderTradeSettlement");
            }

            if (!evidence.HasSpecifiedTradeSettlementHeaderMonetarySummation) {
                missingFields.Add("SpecifiedTradeSettlementHeaderMonetarySummation");
            }

            if (!evidence.HasPayableOrGrandTotalAmount) {
                missingFields.Add("GrandTotalAmount or DuePayableAmount");
            }

            if (missingFields.Count > 0) {
                return new PdfComplianceRequirement(
                    "einvoice-xml-trade-transaction",
                    "EN 16931 XML trade transaction",
                    PdfComplianceRequirementStatus.Missing,
                    "Set factur-x.xml trade transaction essentials before Mustang validation: " + string.Join(", ", missingFields.ToArray()) + ".");
            }

            return new PdfComplianceRequirement(
                "einvoice-xml-trade-transaction",
                "EN 16931 XML trade transaction",
                PdfComplianceRequirementStatus.Satisfied,
                "The factur-x.xml CrossIndustryInvoice includes seller, buyer, settlement, monetary summation, and payable/total amount markers for e-invoice readiness.");
        }

        string diagnostic = diagnostics.Count == 0
            ? "Attach a canonical factur-x.xml CrossIndustryInvoice file before checking EN 16931 XML trade transaction essentials."
            : string.Join(" ", diagnostics.Distinct(StringComparer.Ordinal).ToArray());
        return new PdfComplianceRequirement(
            "einvoice-xml-trade-transaction",
            "EN 16931 XML trade transaction",
            PdfComplianceRequirementStatus.Missing,
            diagnostic);
    }

    private static PdfComplianceRequirement BuildElectronicInvoiceXmlPartyIdentificationRequirement(PdfOptions options) {
        IReadOnlyList<PdfEmbeddedFile> embeddedFiles = options.EmbeddedFiles;
        var diagnostics = new List<string>();
        for (int i = 0; i < embeddedFiles.Count; i++) {
            PdfEmbeddedFile file = embeddedFiles[i];
            if (!IsFacturXCiiAttachment(file, diagnostics)) {
                continue;
            }

            if (!PdfCiiDocumentHeaderInspector.TryReadPartyIdentification(file, out PdfCiiPartyIdentificationEvidence? evidence, out string? partyDiagnostic)) {
                return new PdfComplianceRequirement(
                    "einvoice-xml-party-identification",
                    "EN 16931 XML party identification",
                    PdfComplianceRequirementStatus.Missing,
                    partyDiagnostic!);
            }

            var missingFields = new List<string>();
            if (!evidence!.HasSellerName) {
                missingFields.Add("SellerTradeParty Name");
            }

            if (!evidence.HasSellerCountryId) {
                missingFields.Add("SellerTradeParty PostalTradeAddress CountryID");
            }

            if (!evidence.HasBuyerName) {
                missingFields.Add("BuyerTradeParty Name");
            }

            if (!evidence.HasBuyerCountryId) {
                missingFields.Add("BuyerTradeParty PostalTradeAddress CountryID");
            }

            if (missingFields.Count > 0) {
                return new PdfComplianceRequirement(
                    "einvoice-xml-party-identification",
                    "EN 16931 XML party identification",
                    PdfComplianceRequirementStatus.Missing,
                    "Set factur-x.xml seller and buyer identity essentials before Mustang validation: " + string.Join(", ", missingFields.ToArray()) + ".");
            }

            return new PdfComplianceRequirement(
                "einvoice-xml-party-identification",
                "EN 16931 XML party identification",
                PdfComplianceRequirementStatus.Satisfied,
                "The factur-x.xml CrossIndustryInvoice includes seller and buyer names plus postal country identifiers for e-invoice readiness.");
        }

        string diagnostic = diagnostics.Count == 0
            ? "Attach a canonical factur-x.xml CrossIndustryInvoice file before checking EN 16931 XML party identification essentials."
            : string.Join(" ", diagnostics.Distinct(StringComparer.Ordinal).ToArray());
        return new PdfComplianceRequirement(
            "einvoice-xml-party-identification",
            "EN 16931 XML party identification",
            PdfComplianceRequirementStatus.Missing,
            diagnostic);
    }

    private static PdfComplianceRequirement BuildElectronicInvoiceXmlLineItemRequirement(PdfOptions options) {
        IReadOnlyList<PdfEmbeddedFile> embeddedFiles = options.EmbeddedFiles;
        var diagnostics = new List<string>();
        for (int i = 0; i < embeddedFiles.Count; i++) {
            PdfEmbeddedFile file = embeddedFiles[i];
            if (!IsFacturXCiiAttachment(file, diagnostics)) {
                continue;
            }

            if (!PdfCiiDocumentHeaderInspector.TryReadLineItems(file, out PdfCiiLineItemEvidence? evidence, out string? lineDiagnostic)) {
                return new PdfComplianceRequirement(
                    "einvoice-xml-line-item",
                    "EN 16931 XML line item",
                    PdfComplianceRequirementStatus.Missing,
                    lineDiagnostic!);
            }

            var missingFields = new List<string>();
            if (!evidence!.HasIncludedSupplyChainTradeLineItem) {
                missingFields.Add("IncludedSupplyChainTradeLineItem");
            }

            if (evidence.MissingLineItemFields.Count > 0) {
                missingFields.AddRange(evidence.MissingLineItemFields);
            } else {
                if (!evidence.HasLineId) {
                    missingFields.Add("AssociatedDocumentLineDocument LineID");
                }

                if (!evidence.HasProductName) {
                    missingFields.Add("SpecifiedTradeProduct Name");
                }

                if (!evidence.HasBilledQuantity) {
                    missingFields.Add("SpecifiedLineTradeDelivery BilledQuantity");
                }

                if (!evidence.HasBilledQuantityUnitCode) {
                    missingFields.Add("SpecifiedLineTradeDelivery BilledQuantity unitCode");
                }

                if (!evidence.HasLineTotalAmount) {
                    missingFields.Add("SpecifiedTradeSettlementLineMonetarySummation LineTotalAmount");
                }
            }

            if (missingFields.Count > 0) {
                return new PdfComplianceRequirement(
                    "einvoice-xml-line-item",
                    "EN 16931 XML line item",
                    PdfComplianceRequirementStatus.Missing,
                    "Set factur-x.xml line item essentials before Mustang validation: " + string.Join(", ", missingFields.ToArray()) + ".");
            }

            return new PdfComplianceRequirement(
                "einvoice-xml-line-item",
                "EN 16931 XML line item",
                PdfComplianceRequirementStatus.Satisfied,
                "The factur-x.xml CrossIndustryInvoice includes at least one line item with line id, product name, billed quantity, quantity unit code, and line total amount markers for e-invoice readiness.");
        }

        string diagnostic = diagnostics.Count == 0
            ? "Attach a canonical factur-x.xml CrossIndustryInvoice file before checking EN 16931 XML line item essentials."
            : string.Join(" ", diagnostics.Distinct(StringComparer.Ordinal).ToArray());
        return new PdfComplianceRequirement(
            "einvoice-xml-line-item",
            "EN 16931 XML line item",
            PdfComplianceRequirementStatus.Missing,
            diagnostic);
    }

    private static PdfComplianceRequirement BuildElectronicInvoiceXmlPartyTaxRegistrationRequirement(PdfOptions options) {
        IReadOnlyList<PdfEmbeddedFile> embeddedFiles = options.EmbeddedFiles;
        var diagnostics = new List<string>();
        for (int i = 0; i < embeddedFiles.Count; i++) {
            PdfEmbeddedFile file = embeddedFiles[i];
            if (!IsFacturXCiiAttachment(file, diagnostics)) {
                continue;
            }

            if (!PdfCiiDocumentHeaderInspector.TryReadPartyTaxRegistration(file, out PdfCiiPartyTaxRegistrationEvidence? evidence, out string? taxDiagnostic)) {
                return new PdfComplianceRequirement(
                    "einvoice-xml-party-tax-registration",
                    "EN 16931 XML party tax registration",
                    PdfComplianceRequirementStatus.Missing,
                    taxDiagnostic!);
            }

            var missingFields = new List<string>();
            if (!evidence!.HasSellerTaxRegistrationId) {
                missingFields.Add("SellerTradeParty SpecifiedTaxRegistration ID");
            }

            if (missingFields.Count > 0) {
                return new PdfComplianceRequirement(
                    "einvoice-xml-party-tax-registration",
                    "EN 16931 XML party tax registration",
                    PdfComplianceRequirementStatus.Missing,
                    "Set factur-x.xml seller tax registration essentials before Mustang validation: " + string.Join(", ", missingFields.ToArray()) + ".");
            }

            return new PdfComplianceRequirement(
                "einvoice-xml-party-tax-registration",
                "EN 16931 XML party tax registration",
                PdfComplianceRequirementStatus.Satisfied,
                "The factur-x.xml CrossIndustryInvoice includes seller tax registration identifiers and leaves buyer tax registration to category-specific e-invoice readiness checks.");
        }

        string diagnostic = diagnostics.Count == 0
            ? "Attach a canonical factur-x.xml CrossIndustryInvoice file before checking EN 16931 XML party tax registration essentials."
            : string.Join(" ", diagnostics.Distinct(StringComparer.Ordinal).ToArray());
        return new PdfComplianceRequirement(
            "einvoice-xml-party-tax-registration",
            "EN 16931 XML party tax registration",
            PdfComplianceRequirementStatus.Missing,
            diagnostic);
    }

    private static PdfComplianceRequirement BuildElectronicInvoiceXmlSettlementSummaryRequirement(PdfOptions options) {
        IReadOnlyList<PdfEmbeddedFile> embeddedFiles = options.EmbeddedFiles;
        var diagnostics = new List<string>();
        for (int i = 0; i < embeddedFiles.Count; i++) {
            PdfEmbeddedFile file = embeddedFiles[i];
            if (!IsFacturXCiiAttachment(file, diagnostics)) {
                continue;
            }

            if (!PdfCiiDocumentHeaderInspector.TryReadSettlementSummary(file, out PdfCiiSettlementSummaryEvidence? evidence, out string? settlementDiagnostic)) {
                return new PdfComplianceRequirement(
                    "einvoice-xml-settlement-summary",
                    "EN 16931 XML settlement summary",
                    PdfComplianceRequirementStatus.Missing,
                    settlementDiagnostic!);
            }

            var missingFields = new List<string>();
            if (!evidence!.HasApplicableHeaderTradeSettlement) {
                missingFields.Add("ApplicableHeaderTradeSettlement");
            }

            if (!evidence.HasInvoiceCurrencyCode) {
                missingFields.Add("InvoiceCurrencyCode");
            }

            if (!evidence.HasApplicableTradeTax) {
                missingFields.Add("ApplicableTradeTax");
            }

            if (!evidence.HasTaxBasisTotalAmount) {
                missingFields.Add("SpecifiedTradeSettlementHeaderMonetarySummation TaxBasisTotalAmount");
            }

            if (!evidence.HasTaxTotalAmount) {
                missingFields.Add("SpecifiedTradeSettlementHeaderMonetarySummation TaxTotalAmount");
            }

            if (missingFields.Count > 0) {
                return new PdfComplianceRequirement(
                    "einvoice-xml-settlement-summary",
                    "EN 16931 XML settlement summary",
                    PdfComplianceRequirementStatus.Missing,
                    "Set factur-x.xml settlement summary essentials before Mustang validation: " + string.Join(", ", missingFields.ToArray()) + ".");
            }

            return new PdfComplianceRequirement(
                "einvoice-xml-settlement-summary",
                "EN 16931 XML settlement summary",
                PdfComplianceRequirementStatus.Satisfied,
                "The factur-x.xml CrossIndustryInvoice includes invoice currency, trade tax, tax basis total, and tax total amount markers for e-invoice readiness.");
        }

        string diagnostic = diagnostics.Count == 0
            ? "Attach a canonical factur-x.xml CrossIndustryInvoice file before checking EN 16931 XML settlement summary essentials."
            : string.Join(" ", diagnostics.Distinct(StringComparer.Ordinal).ToArray());
        return new PdfComplianceRequirement(
            "einvoice-xml-settlement-summary",
            "EN 16931 XML settlement summary",
            PdfComplianceRequirementStatus.Missing,
            diagnostic);
    }

    private static PdfComplianceRequirement BuildElectronicInvoiceXmlLinePricingRequirement(PdfOptions options) {
        IReadOnlyList<PdfEmbeddedFile> embeddedFiles = options.EmbeddedFiles;
        var diagnostics = new List<string>();
        for (int i = 0; i < embeddedFiles.Count; i++) {
            PdfEmbeddedFile file = embeddedFiles[i];
            if (!IsFacturXCiiAttachment(file, diagnostics)) {
                continue;
            }

            if (!PdfCiiDocumentHeaderInspector.TryReadLinePricing(file, out PdfCiiLinePricingEvidence? evidence, out string? pricingDiagnostic)) {
                return new PdfComplianceRequirement(
                    "einvoice-xml-line-pricing",
                    "EN 16931 XML line pricing",
                    PdfComplianceRequirementStatus.Missing,
                    pricingDiagnostic!);
            }

            var missingFields = new List<string>();
            if (!evidence!.HasIncludedSupplyChainTradeLineItem) {
                missingFields.Add("IncludedSupplyChainTradeLineItem");
            }

            if (!evidence.HasSpecifiedLineTradeAgreement) {
                missingFields.Add("SpecifiedLineTradeAgreement");
            }

            if (!evidence.HasProductTradePrice) {
                missingFields.Add("NetPriceProductTradePrice");
            }

            if (!evidence.HasPriceChargeAmount) {
                missingFields.Add("NetPriceProductTradePrice ChargeAmount");
            }

            if (evidence.MissingLinePricingFields.Count > 0) {
                missingFields.AddRange(evidence.MissingLinePricingFields);
            }

            if (missingFields.Count > 0) {
                return new PdfComplianceRequirement(
                    "einvoice-xml-line-pricing",
                    "EN 16931 XML line pricing",
                    PdfComplianceRequirementStatus.Missing,
                    "Set factur-x.xml line pricing essentials before Mustang validation: " + string.Join(", ", missingFields.ToArray()) + ".");
            }

            return new PdfComplianceRequirement(
                "einvoice-xml-line-pricing",
                "EN 16931 XML line pricing",
                PdfComplianceRequirementStatus.Satisfied,
                "The factur-x.xml CrossIndustryInvoice includes line trade agreement and product trade price charge amount markers for e-invoice readiness.");
        }

        string diagnostic = diagnostics.Count == 0
            ? "Attach a canonical factur-x.xml CrossIndustryInvoice file before checking EN 16931 XML line pricing essentials."
            : string.Join(" ", diagnostics.Distinct(StringComparer.Ordinal).ToArray());
        return new PdfComplianceRequirement(
            "einvoice-xml-line-pricing",
            "EN 16931 XML line pricing",
            PdfComplianceRequirementStatus.Missing,
            diagnostic);
    }

    private static PdfComplianceRequirement BuildElectronicInvoiceXmlAmountConsistencyRequirement(PdfOptions options) {
        IReadOnlyList<PdfEmbeddedFile> embeddedFiles = options.EmbeddedFiles;
        var diagnostics = new List<string>();
        for (int i = 0; i < embeddedFiles.Count; i++) {
            PdfEmbeddedFile file = embeddedFiles[i];
            if (!IsFacturXCiiAttachment(file, diagnostics)) {
                continue;
            }

            if (!PdfCiiDocumentHeaderInspector.TryReadAmountConsistency(file, out PdfCiiAmountConsistencyEvidence? evidence, out string? amountDiagnostic)) {
                return new PdfComplianceRequirement(
                    "einvoice-xml-amount-consistency",
                    "EN 16931 XML amount consistency",
                    PdfComplianceRequirementStatus.Missing,
                    amountDiagnostic!);
            }

            if (!string.IsNullOrWhiteSpace(evidence!.ParseDiagnostic)) {
                return new PdfComplianceRequirement(
                    "einvoice-xml-amount-consistency",
                    "EN 16931 XML amount consistency",
                    PdfComplianceRequirementStatus.Missing,
                    evidence.ParseDiagnostic!);
            }

            var missingFields = new List<string>();
            if (!evidence.LineTotalAmountSum.HasValue) {
                missingFields.Add("LineTotalAmount");
            }

            if (!evidence.TaxBasisTotalAmount.HasValue) {
                missingFields.Add("TaxBasisTotalAmount");
            }

            if (!evidence.TaxTotalAmount.HasValue) {
                missingFields.Add("TaxTotalAmount");
            }

            if (!evidence.GrandOrDuePayableAmount.HasValue) {
                missingFields.Add("GrandTotalAmount or DuePayableAmount");
            }

            if (missingFields.Count > 0) {
                return new PdfComplianceRequirement(
                    "einvoice-xml-amount-consistency",
                    "EN 16931 XML amount consistency",
                    PdfComplianceRequirementStatus.Missing,
                    "Set factur-x.xml parseable amounts before Mustang validation: " + string.Join(", ", missingFields.ToArray()) + ".");
            }

            var mismatchDiagnostics = new List<string>();
            if (!evidence.LineTotalMatchesTaxBasis) {
                mismatchDiagnostics.Add("LineTotalAmount sum minus AllowanceTotalAmount plus ChargeTotalAmount must match TaxBasisTotalAmount.");
            }

            if (!evidence.AllowanceTotalMatchesDocumentLevelAllowances) {
                mismatchDiagnostics.Add("AllowanceTotalAmount must match the sum of document-level allowance ActualAmount values.");
            }

            if (!evidence.ChargeTotalMatchesDocumentLevelCharges) {
                mismatchDiagnostics.Add("ChargeTotalAmount must match the sum of document-level charge ActualAmount values.");
            }

            if (!evidence.GrandTotalMatchesBasisPlusTax) {
                mismatchDiagnostics.Add("GrandTotalAmount must match TaxBasisTotalAmount plus TaxTotalAmount.");
            }

            if (!evidence.DuePayableMatchesGrandTotal) {
                mismatchDiagnostics.Add("DuePayableAmount must match GrandTotalAmount minus PaidAmount plus RoundingAmount.");
            }

            if (mismatchDiagnostics.Count > 0) {
                return new PdfComplianceRequirement(
                    "einvoice-xml-amount-consistency",
                    "EN 16931 XML amount consistency",
                    PdfComplianceRequirementStatus.Missing,
                    "Fix factur-x.xml amount totals before Mustang validation: " + string.Join(" ", mismatchDiagnostics.ToArray()));
            }

            return new PdfComplianceRequirement(
                "einvoice-xml-amount-consistency",
                "EN 16931 XML amount consistency",
                PdfComplianceRequirementStatus.Satisfied,
                "The factur-x.xml CrossIndustryInvoice line, document-level allowance, document-level charge, tax basis, tax total, grand total, paid, rounding, and due payable amount markers are internally consistent for e-invoice readiness.");
        }

        string diagnostic = diagnostics.Count == 0
            ? "Attach a canonical factur-x.xml CrossIndustryInvoice file before checking EN 16931 XML amount consistency."
            : string.Join(" ", diagnostics.Distinct(StringComparer.Ordinal).ToArray());
        return new PdfComplianceRequirement(
            "einvoice-xml-amount-consistency",
            "EN 16931 XML amount consistency",
            PdfComplianceRequirementStatus.Missing,
            diagnostic);
    }

    private static PdfComplianceRequirement BuildElectronicInvoiceXmlTaxBreakdownRequirement(PdfOptions options) {
        IReadOnlyList<PdfEmbeddedFile> embeddedFiles = options.EmbeddedFiles;
        var diagnostics = new List<string>();
        for (int i = 0; i < embeddedFiles.Count; i++) {
            PdfEmbeddedFile file = embeddedFiles[i];
            if (!IsFacturXCiiAttachment(file, diagnostics)) {
                continue;
            }

            if (!PdfCiiDocumentHeaderInspector.TryReadTaxBreakdown(file, out PdfCiiTaxBreakdownEvidence? evidence, out string? taxDiagnostic)) {
                return new PdfComplianceRequirement(
                    "einvoice-xml-tax-breakdown",
                    "EN 16931 XML tax breakdown",
                    PdfComplianceRequirementStatus.Missing,
                    taxDiagnostic!);
            }

            var missingFields = new List<string>();
            if (!evidence!.HasApplicableTradeTax) {
                missingFields.Add("ApplicableTradeTax");
            }

            if (evidence.MissingTypeCodeBreakdowns.Count > 0) {
                missingFields.Add("ApplicableTradeTax TypeCode on " + string.Join(", ", evidence.MissingTypeCodeBreakdowns.ToArray()));
            } else if (!evidence.HasTypeCode) {
                missingFields.Add("ApplicableTradeTax TypeCode");
            }

            if (evidence.MissingCategoryCodeBreakdowns.Count > 0) {
                missingFields.Add("ApplicableTradeTax CategoryCode on " + string.Join(", ", evidence.MissingCategoryCodeBreakdowns.ToArray()));
            } else if (!evidence.HasCategoryCode) {
                missingFields.Add("ApplicableTradeTax CategoryCode");
            }

            if (!evidence.HasRateApplicablePercent) {
                missingFields.Add("ApplicableTradeTax RateApplicablePercent");
            }

            if (!evidence.HasBasisAmount) {
                missingFields.Add("ApplicableTradeTax BasisAmount");
            }

            if (!evidence.HasCalculatedAmount) {
                missingFields.Add("ApplicableTradeTax CalculatedAmount");
            }

            if (missingFields.Count > 0) {
                return new PdfComplianceRequirement(
                    "einvoice-xml-tax-breakdown",
                    "EN 16931 XML tax breakdown",
                    PdfComplianceRequirementStatus.Missing,
                    "Set factur-x.xml tax breakdown essentials before Mustang validation: " + string.Join(", ", missingFields.ToArray()) + ".");
            }

            var invalidTypeCodes = new List<string>();
            for (int j = 0; j < evidence.TypeCodes.Count; j++) {
                string typeCode = evidence.TypeCodes[j];
                if (!string.Equals(typeCode.Trim(), "VAT", StringComparison.Ordinal)) {
                    invalidTypeCodes.Add(typeCode);
                }
            }

            if (invalidTypeCodes.Count > 0) {
                return new PdfComplianceRequirement(
                    "einvoice-xml-tax-breakdown",
                    "EN 16931 XML tax breakdown",
                    PdfComplianceRequirementStatus.Missing,
                    "Set factur-x.xml ApplicableTradeTax TypeCode to VAT before Mustang validation. Found: " + string.Join(", ", invalidTypeCodes.Distinct(StringComparer.Ordinal).ToArray()) + ".");
            }

            return new PdfComplianceRequirement(
                "einvoice-xml-tax-breakdown",
                "EN 16931 XML tax breakdown",
                PdfComplianceRequirementStatus.Satisfied,
                "The factur-x.xml CrossIndustryInvoice includes VAT trade tax type, category, rate, basis, and calculated amount markers for e-invoice readiness.");
        }

        string diagnostic = diagnostics.Count == 0
            ? "Attach a canonical factur-x.xml CrossIndustryInvoice file before checking EN 16931 XML tax breakdown essentials."
            : string.Join(" ", diagnostics.Distinct(StringComparer.Ordinal).ToArray());
        return new PdfComplianceRequirement(
            "einvoice-xml-tax-breakdown",
            "EN 16931 XML tax breakdown",
            PdfComplianceRequirementStatus.Missing,
            diagnostic);
    }

    private static PdfComplianceRequirement BuildElectronicInvoiceXmlPaymentInstructionsRequirement(PdfOptions options) {
        IReadOnlyList<PdfEmbeddedFile> embeddedFiles = options.EmbeddedFiles;
        var diagnostics = new List<string>();
        for (int i = 0; i < embeddedFiles.Count; i++) {
            PdfEmbeddedFile file = embeddedFiles[i];
            if (!IsFacturXCiiAttachment(file, diagnostics)) {
                continue;
            }

            if (!PdfCiiDocumentHeaderInspector.TryReadPaymentInstructions(file, out PdfCiiPaymentInstructionEvidence? evidence, out string? paymentDiagnostic)) {
                return new PdfComplianceRequirement(
                    "einvoice-xml-payment-instructions",
                    "EN 16931 XML payment instructions",
                    PdfComplianceRequirementStatus.Missing,
                    paymentDiagnostic!);
            }

            var missingFields = new List<string>();
            if (!evidence!.HasSpecifiedTradeSettlementPaymentMeans) {
                missingFields.Add("SpecifiedTradeSettlementPaymentMeans");
            }

            if (!evidence.HasTypeCode) {
                missingFields.Add("SpecifiedTradeSettlementPaymentMeans TypeCode");
            }

            if (evidence.MissingTypeCodePaymentMeans.Count > 0) {
                missingFields.Add("SpecifiedTradeSettlementPaymentMeans TypeCode on " + string.Join(", ", evidence.MissingTypeCodePaymentMeans.ToArray()));
            }

            bool requiresCreditorAccount = RequiresElectronicInvoiceCreditorAccount(evidence.TypeCodes);
            bool hasCreditorAccountData = evidence.HasCreditorFinancialAccount || evidence.HasCreditorAccountId;
            if (evidence.MissingCreditorAccountPaymentMeans.Count > 0) {
                missingFields.Add("PayeePartyCreditorFinancialAccount on " + string.Join(", ", evidence.MissingCreditorAccountPaymentMeans.ToArray()));
            } else if ((requiresCreditorAccount || hasCreditorAccountData) && !evidence.HasCreditorFinancialAccount) {
                missingFields.Add("PayeePartyCreditorFinancialAccount");
            }

            if (evidence.MissingCreditorAccountIdPaymentMeans.Count > 0) {
                missingFields.Add("PayeePartyCreditorFinancialAccount IBANID or ProprietaryID on " + string.Join(", ", evidence.MissingCreditorAccountIdPaymentMeans.ToArray()));
            } else if ((requiresCreditorAccount || hasCreditorAccountData) && !evidence.HasCreditorAccountId) {
                missingFields.Add("PayeePartyCreditorFinancialAccount IBANID or ProprietaryID");
            }

            if (missingFields.Count > 0) {
                return new PdfComplianceRequirement(
                    "einvoice-xml-payment-instructions",
                    "EN 16931 XML payment instructions",
                    PdfComplianceRequirementStatus.Missing,
                    "Set factur-x.xml payment instruction essentials before Mustang validation: " + string.Join(", ", missingFields.ToArray()) + ".");
            }

            return new PdfComplianceRequirement(
                "einvoice-xml-payment-instructions",
                "EN 16931 XML payment instructions",
                PdfComplianceRequirementStatus.Satisfied,
                requiresCreditorAccount
                    ? "The factur-x.xml CrossIndustryInvoice includes payment means type code and creditor account identifiers for e-invoice readiness."
                    : "The factur-x.xml CrossIndustryInvoice includes a payment means type code that does not require creditor account identifiers for e-invoice readiness.");
        }

        string diagnostic = diagnostics.Count == 0
            ? "Attach a canonical factur-x.xml CrossIndustryInvoice file before checking EN 16931 XML payment instruction essentials."
            : string.Join(" ", diagnostics.Distinct(StringComparer.Ordinal).ToArray());
        return new PdfComplianceRequirement(
            "einvoice-xml-payment-instructions",
            "EN 16931 XML payment instructions",
            PdfComplianceRequirementStatus.Missing,
            diagnostic);
    }

    private static PdfComplianceRequirement BuildElectronicInvoiceXmlPaymentTermsRequirement(PdfOptions options) {
        IReadOnlyList<PdfEmbeddedFile> embeddedFiles = options.EmbeddedFiles;
        var diagnostics = new List<string>();
        for (int i = 0; i < embeddedFiles.Count; i++) {
            PdfEmbeddedFile file = embeddedFiles[i];
            if (!IsFacturXCiiAttachment(file, diagnostics)) {
                continue;
            }

            if (!PdfCiiDocumentHeaderInspector.TryReadPaymentTerms(file, out PdfCiiPaymentTermsEvidence? evidence, out string? termsDiagnostic)) {
                return new PdfComplianceRequirement(
                    "einvoice-xml-payment-terms",
                    "EN 16931 XML payment terms",
                    PdfComplianceRequirementStatus.Missing,
                    termsDiagnostic!);
            }

            var missingFields = new List<string>();
            if (!evidence!.HasSpecifiedTradePaymentTerms) {
                missingFields.Add("SpecifiedTradePaymentTerms");
            }

            if (!evidence.HasDueDateOrDescription) {
                missingFields.Add("SpecifiedTradePaymentTerms DueDateDateTime or Description");
            }

            if (missingFields.Count > 0) {
                return new PdfComplianceRequirement(
                    "einvoice-xml-payment-terms",
                    "EN 16931 XML payment terms",
                    PdfComplianceRequirementStatus.Missing,
                    "Set factur-x.xml payment terms essentials before Mustang validation: " + string.Join(", ", missingFields.ToArray()) + ".");
            }

            return new PdfComplianceRequirement(
                "einvoice-xml-payment-terms",
                "EN 16931 XML payment terms",
                PdfComplianceRequirementStatus.Satisfied,
                "The factur-x.xml CrossIndustryInvoice includes payment terms with due date or description markers for e-invoice readiness.");
        }

        string diagnostic = diagnostics.Count == 0
            ? "Attach a canonical factur-x.xml CrossIndustryInvoice file before checking EN 16931 XML payment terms essentials."
            : string.Join(" ", diagnostics.Distinct(StringComparer.Ordinal).ToArray());
        return new PdfComplianceRequirement(
            "einvoice-xml-payment-terms",
            "EN 16931 XML payment terms",
            PdfComplianceRequirementStatus.Missing,
            diagnostic);
    }

    private static PdfComplianceRequirement BuildElectronicInvoiceXmlAttachmentParamsRequirement(PdfOptions options) {
        IReadOnlyList<PdfEmbeddedFile> embeddedFiles = options.EmbeddedFiles;
        for (int i = 0; i < embeddedFiles.Count; i++) {
            var diagnostics = new List<string>();
            if (IsFacturXCiiAttachment(embeddedFiles[i], diagnostics)) {
                return new PdfComplianceRequirement(
                    "einvoice-xml-attachment-params",
                    "EN 16931 embedded-file parameters",
                    PdfComplianceRequirementStatus.Satisfied,
                    "Generated factur-x.xml embedded-file streams will include deterministic /Params /Size and /CheckSum metadata.");
            }
        }

        return new PdfComplianceRequirement(
            "einvoice-xml-attachment-params",
            "EN 16931 embedded-file parameters",
            PdfComplianceRequirementStatus.Missing,
            "Attach a canonical factur-x.xml CrossIndustryInvoice associated file so generated embedded-file stream parameters can be emitted for the e-invoice payload.");
    }
}
