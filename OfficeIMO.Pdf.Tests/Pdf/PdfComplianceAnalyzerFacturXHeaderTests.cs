using System.Text;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public partial class PdfComplianceAnalyzerTests {
    [Fact]
    public void FacturXReadinessRequiresCiiDocumentHeaderEssentials() {
        var missingHeaderOptions = new PdfOptions()
            .SetPdfAIdentification(3, "B")
            .SetSrgbOutputIntent()
            .SetElectronicInvoiceMetadata(PdfElectronicInvoiceMetadata.FacturX("EN 16931"))
            .AddEmbeddedFile("factur-x.xml", CreateCiiXml(includeDocumentHeader: false), "application/xml", PdfAssociatedFileRelationship.Data);
        var missingTradeOptions = new PdfOptions()
            .SetPdfAIdentification(3, "B")
            .SetSrgbOutputIntent()
            .SetElectronicInvoiceMetadata(PdfElectronicInvoiceMetadata.FacturX("EN 16931"))
            .AddEmbeddedFile("factur-x.xml", CreateCiiXml(includeSupplyChainTradeTransaction: false), "application/xml", PdfAssociatedFileRelationship.Data);

        PdfComplianceRequirement missingHeader = AssertRequirement(
            PdfComplianceAnalyzer.Assess(PdfComplianceProfile.FacturX, missingHeaderOptions),
            "einvoice-xml-document-header",
            PdfComplianceRequirementStatus.Missing);
        PdfComplianceRequirement missingTrade = AssertRequirement(
            PdfComplianceAnalyzer.Assess(PdfComplianceProfile.FacturX, missingTradeOptions),
            "einvoice-xml-document-header",
            PdfComplianceRequirementStatus.Missing);

        Assert.Contains("ExchangedDocument ID", missingHeader.Diagnostic);
        Assert.Contains("ExchangedDocument TypeCode", missingHeader.Diagnostic);
        Assert.Contains("ExchangedDocument IssueDateTime", missingHeader.Diagnostic);
        Assert.Contains("SupplyChainTradeTransaction", missingTrade.Diagnostic);
    }

    [Fact]
    public void FacturXReadinessRequiresCiiDocumentTypeCodeListValue() {
        var invalidTypeCodeOptions = new PdfOptions()
            .SetPdfAIdentification(3, "B")
            .SetSrgbOutputIntent()
            .SetElectronicInvoiceMetadata(PdfElectronicInvoiceMetadata.FacturX("EN 16931"))
            .AddEmbeddedFile("factur-x.xml", CreateCiiXml(documentTypeCodeValue: "999"), "application/xml", PdfAssociatedFileRelationship.Data);
        var creditNoteTypeCodeOptions = new PdfOptions()
            .SetPdfAIdentification(3, "B")
            .SetSrgbOutputIntent()
            .SetElectronicInvoiceMetadata(PdfElectronicInvoiceMetadata.FacturX("EN 16931"))
            .AddEmbeddedFile("factur-x.xml", CreateCiiXml(documentTypeCodeValue: "381"), "application/xml", PdfAssociatedFileRelationship.Data);

        PdfComplianceRequirement invalidTypeCode = AssertRequirement(
            PdfComplianceAnalyzer.Assess(PdfComplianceProfile.FacturX, invalidTypeCodeOptions),
            "einvoice-xml-document-type-code",
            PdfComplianceRequirementStatus.Missing);
        PdfComplianceRequirement creditNoteTypeCode = AssertRequirement(
            PdfComplianceAnalyzer.Assess(PdfComplianceProfile.FacturX, creditNoteTypeCodeOptions),
            "einvoice-xml-document-type-code",
            PdfComplianceRequirementStatus.Satisfied);

        Assert.Contains("ExchangedDocument TypeCode", invalidTypeCode.Diagnostic);
        Assert.Contains("UNTDID 1001", invalidTypeCode.Diagnostic);
        Assert.Contains("999", invalidTypeCode.Diagnostic);
        Assert.Contains("UNTDID 1001", creditNoteTypeCode.Diagnostic);
    }

    [Fact]
    public void FacturXReadinessRequiresCiiDateFormats() {
        var invalidIssueDateOptions = new PdfOptions()
            .SetPdfAIdentification(3, "B")
            .SetSrgbOutputIntent()
            .SetElectronicInvoiceMetadata(PdfElectronicInvoiceMetadata.FacturX("EN 16931"))
            .AddEmbeddedFile("factur-x.xml", CreateCiiXml(issueDateTimeValue: "20261340"), "application/xml", PdfAssociatedFileRelationship.Data);
        var invalidDueDateOptions = new PdfOptions()
            .SetPdfAIdentification(3, "B")
            .SetSrgbOutputIntent()
            .SetElectronicInvoiceMetadata(PdfElectronicInvoiceMetadata.FacturX("EN 16931"))
            .AddEmbeddedFile("factur-x.xml", CreateCiiXml(dueDateTimeValue: "20260230"), "application/xml", PdfAssociatedFileRelationship.Data);
        var validDateTimeOptions = new PdfOptions()
            .SetPdfAIdentification(3, "B")
            .SetSrgbOutputIntent()
            .SetElectronicInvoiceMetadata(PdfElectronicInvoiceMetadata.FacturX("EN 16931"))
            .AddEmbeddedFile("factur-x.xml", CreateCiiXml(issueDateTimeFormat: "203", issueDateTimeValue: "202606031430"), "application/xml", PdfAssociatedFileRelationship.Data);
        var missingIssueFormatOptions = new PdfOptions()
            .SetPdfAIdentification(3, "B")
            .SetSrgbOutputIntent()
            .SetElectronicInvoiceMetadata(PdfElectronicInvoiceMetadata.FacturX("EN 16931"))
            .AddEmbeddedFile("factur-x.xml", CreateCiiXml(issueDateTimeFormat: ""), "application/xml", PdfAssociatedFileRelationship.Data);
        var unknownDueFormatOptions = new PdfOptions()
            .SetPdfAIdentification(3, "B")
            .SetSrgbOutputIntent()
            .SetElectronicInvoiceMetadata(PdfElectronicInvoiceMetadata.FacturX("EN 16931"))
            .AddEmbeddedFile("factur-x.xml", CreateCiiXml(dueDateTimeFormat: "999"), "application/xml", PdfAssociatedFileRelationship.Data);

        PdfComplianceRequirement invalidIssueDate = AssertRequirement(
            PdfComplianceAnalyzer.Assess(PdfComplianceProfile.FacturX, invalidIssueDateOptions),
            "einvoice-xml-date-format",
            PdfComplianceRequirementStatus.Missing);
        PdfComplianceRequirement invalidDueDate = AssertRequirement(
            PdfComplianceAnalyzer.Assess(PdfComplianceProfile.FacturX, invalidDueDateOptions),
            "einvoice-xml-date-format",
            PdfComplianceRequirementStatus.Missing);
        PdfComplianceRequirement validDateTime = AssertRequirement(
            PdfComplianceAnalyzer.Assess(PdfComplianceProfile.FacturX, validDateTimeOptions),
            "einvoice-xml-date-format",
            PdfComplianceRequirementStatus.Satisfied);
        PdfComplianceRequirement missingIssueFormat = AssertRequirement(
            PdfComplianceAnalyzer.Assess(PdfComplianceProfile.FacturX, missingIssueFormatOptions),
            "einvoice-xml-date-format",
            PdfComplianceRequirementStatus.Missing);
        PdfComplianceRequirement unknownDueFormat = AssertRequirement(
            PdfComplianceAnalyzer.Assess(PdfComplianceProfile.FacturX, unknownDueFormatOptions),
            "einvoice-xml-date-format",
            PdfComplianceRequirementStatus.Missing);

        Assert.Contains("ExchangedDocument IssueDateTime", invalidIssueDate.Diagnostic);
        Assert.Contains("SpecifiedTradePaymentTerms DueDateDateTime", invalidDueDate.Diagnostic);
        Assert.Contains("DateTimeString", invalidDueDate.Diagnostic);
        Assert.Contains("ExchangedDocument IssueDateTime", missingIssueFormat.Diagnostic);
        Assert.Contains("SpecifiedTradePaymentTerms DueDateDateTime", unknownDueFormat.Diagnostic);
        Assert.Contains("issue", validDateTime.Diagnostic);
    }

    [Fact]
    public void FacturXReadinessRequiresCiiTradeTransactionEssentials() {
        var emptyTransactionOptions = new PdfOptions()
            .SetPdfAIdentification(3, "B")
            .SetSrgbOutputIntent()
            .SetElectronicInvoiceMetadata(PdfElectronicInvoiceMetadata.FacturX("EN 16931"))
            .AddEmbeddedFile("factur-x.xml", CreateCiiXml(includeTradeTransactionEssentials: false), "application/xml", PdfAssociatedFileRelationship.Data);
        var missingSellerOptions = new PdfOptions()
            .SetPdfAIdentification(3, "B")
            .SetSrgbOutputIntent()
            .SetElectronicInvoiceMetadata(PdfElectronicInvoiceMetadata.FacturX("EN 16931"))
            .AddEmbeddedFile("factur-x.xml", CreateCiiXml(includeSellerTradeParty: false), "application/xml", PdfAssociatedFileRelationship.Data);
        var missingAmountOptions = new PdfOptions()
            .SetPdfAIdentification(3, "B")
            .SetSrgbOutputIntent()
            .SetElectronicInvoiceMetadata(PdfElectronicInvoiceMetadata.FacturX("EN 16931"))
            .AddEmbeddedFile("factur-x.xml", CreateCiiXml(includePayableAmount: false), "application/xml", PdfAssociatedFileRelationship.Data);

        PdfComplianceRequirement emptyTransaction = AssertRequirement(
            PdfComplianceAnalyzer.Assess(PdfComplianceProfile.FacturX, emptyTransactionOptions),
            "einvoice-xml-trade-transaction",
            PdfComplianceRequirementStatus.Missing);
        PdfComplianceRequirement missingSeller = AssertRequirement(
            PdfComplianceAnalyzer.Assess(PdfComplianceProfile.FacturX, missingSellerOptions),
            "einvoice-xml-trade-transaction",
            PdfComplianceRequirementStatus.Missing);
        PdfComplianceRequirement missingAmount = AssertRequirement(
            PdfComplianceAnalyzer.Assess(PdfComplianceProfile.FacturX, missingAmountOptions),
            "einvoice-xml-trade-transaction",
            PdfComplianceRequirementStatus.Missing);

        Assert.Contains("ApplicableHeaderTradeAgreement", emptyTransaction.Diagnostic);
        Assert.Contains("SellerTradeParty", emptyTransaction.Diagnostic);
        Assert.Contains("BuyerTradeParty", emptyTransaction.Diagnostic);
        Assert.Contains("ApplicableHeaderTradeSettlement", emptyTransaction.Diagnostic);
        Assert.Contains("SpecifiedTradeSettlementHeaderMonetarySummation", emptyTransaction.Diagnostic);
        Assert.Contains("SellerTradeParty", missingSeller.Diagnostic);
        Assert.Contains("GrandTotalAmount or DuePayableAmount", missingAmount.Diagnostic);
    }

    [Fact]
    public void FacturXReadinessPreservesGrandTotalWhenDuePayableAmountIsBlank() {
        var options = new PdfOptions()
            .SetPdfAIdentification(3, "B")
            .SetSrgbOutputIntent()
            .SetElectronicInvoiceMetadata(PdfElectronicInvoiceMetadata.FacturX("EN 16931"))
            .AddEmbeddedFile(
                "factur-x.xml",
                CreateCiiXml(includeDuePayableAmount: true, duePayableAmount: string.Empty),
                "application/xml",
                PdfAssociatedFileRelationship.Data);

        PdfComplianceReadinessReport report = PdfComplianceAnalyzer.Assess(PdfComplianceProfile.FacturX, options);

        AssertRequirement(report, "einvoice-xml-trade-transaction", PdfComplianceRequirementStatus.Satisfied);
    }

    [Fact]
    public void FacturXReadinessRequiresCiiPartyIdentificationEssentials() {
        var missingSellerCountryOptions = new PdfOptions()
            .SetPdfAIdentification(3, "B")
            .SetSrgbOutputIntent()
            .SetElectronicInvoiceMetadata(PdfElectronicInvoiceMetadata.FacturX("EN 16931"))
            .AddEmbeddedFile("factur-x.xml", CreateCiiXml(includeSellerCountryId: false), "application/xml", PdfAssociatedFileRelationship.Data);
        var missingBuyerNameOptions = new PdfOptions()
            .SetPdfAIdentification(3, "B")
            .SetSrgbOutputIntent()
            .SetElectronicInvoiceMetadata(PdfElectronicInvoiceMetadata.FacturX("EN 16931"))
            .AddEmbeddedFile("factur-x.xml", CreateCiiXml(includeBuyerName: false), "application/xml", PdfAssociatedFileRelationship.Data);
        var missingBuyerCountryOptions = new PdfOptions()
            .SetPdfAIdentification(3, "B")
            .SetSrgbOutputIntent()
            .SetElectronicInvoiceMetadata(PdfElectronicInvoiceMetadata.FacturX("EN 16931"))
            .AddEmbeddedFile("factur-x.xml", CreateCiiXml(includeBuyerCountryId: false), "application/xml", PdfAssociatedFileRelationship.Data);

        PdfComplianceRequirement missingSellerCountry = AssertRequirement(
            PdfComplianceAnalyzer.Assess(PdfComplianceProfile.FacturX, missingSellerCountryOptions),
            "einvoice-xml-party-identification",
            PdfComplianceRequirementStatus.Missing);
        PdfComplianceRequirement missingBuyerName = AssertRequirement(
            PdfComplianceAnalyzer.Assess(PdfComplianceProfile.FacturX, missingBuyerNameOptions),
            "einvoice-xml-party-identification",
            PdfComplianceRequirementStatus.Missing);
        PdfComplianceRequirement missingBuyerCountry = AssertRequirement(
            PdfComplianceAnalyzer.Assess(PdfComplianceProfile.FacturX, missingBuyerCountryOptions),
            "einvoice-xml-party-identification",
            PdfComplianceRequirementStatus.Missing);

        Assert.Contains("SellerTradeParty PostalTradeAddress CountryID", missingSellerCountry.Diagnostic);
        Assert.Contains("BuyerTradeParty Name", missingBuyerName.Diagnostic);
        Assert.Contains("BuyerTradeParty PostalTradeAddress CountryID", missingBuyerCountry.Diagnostic);
    }

    [Fact]
    public void FacturXReadinessRequiresCiiCountryCodeListValue() {
        var invalidSellerCountryOptions = new PdfOptions()
            .SetPdfAIdentification(3, "B")
            .SetSrgbOutputIntent()
            .SetElectronicInvoiceMetadata(PdfElectronicInvoiceMetadata.FacturX("EN 16931"))
            .AddEmbeddedFile("factur-x.xml", CreateCiiXml(sellerCountryIdValue: "POL"), "application/xml", PdfAssociatedFileRelationship.Data);
        var invalidBuyerCountryOptions = new PdfOptions()
            .SetPdfAIdentification(3, "B")
            .SetSrgbOutputIntent()
            .SetElectronicInvoiceMetadata(PdfElectronicInvoiceMetadata.FacturX("EN 16931"))
            .AddEmbeddedFile("factur-x.xml", CreateCiiXml(buyerCountryIdValue: "D"), "application/xml", PdfAssociatedFileRelationship.Data);
        var swissCountryOptions = new PdfOptions()
            .SetPdfAIdentification(3, "B")
            .SetSrgbOutputIntent()
            .SetElectronicInvoiceMetadata(PdfElectronicInvoiceMetadata.FacturX("EN 16931"))
            .AddEmbeddedFile("factur-x.xml", CreateCiiXml(sellerCountryIdValue: "CH", buyerCountryIdValue: "LI"), "application/xml", PdfAssociatedFileRelationship.Data);

        PdfComplianceRequirement invalidSellerCountry = AssertRequirement(
            PdfComplianceAnalyzer.Assess(PdfComplianceProfile.FacturX, invalidSellerCountryOptions),
            "einvoice-xml-country-code",
            PdfComplianceRequirementStatus.Missing);
        PdfComplianceRequirement invalidBuyerCountry = AssertRequirement(
            PdfComplianceAnalyzer.Assess(PdfComplianceProfile.FacturX, invalidBuyerCountryOptions),
            "einvoice-xml-country-code",
            PdfComplianceRequirementStatus.Missing);
        PdfComplianceRequirement swissCountry = AssertRequirement(
            PdfComplianceAnalyzer.Assess(PdfComplianceProfile.FacturX, swissCountryOptions),
            "einvoice-xml-country-code",
            PdfComplianceRequirementStatus.Satisfied);

        Assert.Contains("ISO 3166-1 alpha-2", invalidSellerCountry.Diagnostic);
        Assert.Contains("POL", invalidSellerCountry.Diagnostic);
        Assert.Contains("D", invalidBuyerCountry.Diagnostic);
        Assert.Contains("ISO 3166-1 alpha-2", swissCountry.Diagnostic);
    }

    [Fact]
    public void FacturXReadinessRequiresCiiElectronicAddressSchemeListValue() {
        var missingSellerAddressOptions = new PdfOptions()
            .SetPdfAIdentification(3, "B")
            .SetSrgbOutputIntent()
            .SetElectronicInvoiceMetadata(PdfElectronicInvoiceMetadata.FacturX("EN 16931"))
            .AddEmbeddedFile("factur-x.xml", CreateCiiXml(includeSellerElectronicAddress: false), "application/xml", PdfAssociatedFileRelationship.Data);
        var missingBuyerSchemeOptions = new PdfOptions()
            .SetPdfAIdentification(3, "B")
            .SetSrgbOutputIntent()
            .SetElectronicInvoiceMetadata(PdfElectronicInvoiceMetadata.FacturX("EN 16931"))
            .AddEmbeddedFile("factur-x.xml", CreateCiiXml(includeBuyerElectronicAddressSchemeId: false), "application/xml", PdfAssociatedFileRelationship.Data);
        var invalidSchemeOptions = new PdfOptions()
            .SetPdfAIdentification(3, "B")
            .SetSrgbOutputIntent()
            .SetElectronicInvoiceMetadata(PdfElectronicInvoiceMetadata.FacturX("EN 16931"))
            .AddEmbeddedFile("factur-x.xml", CreateCiiXml(sellerElectronicAddressSchemeIdValue: "9999"), "application/xml", PdfAssociatedFileRelationship.Data);
        var leitwegIdOptions = new PdfOptions()
            .SetPdfAIdentification(3, "B")
            .SetSrgbOutputIntent()
            .SetElectronicInvoiceMetadata(PdfElectronicInvoiceMetadata.FacturX("EN 16931"))
            .AddEmbeddedFile("factur-x.xml", CreateCiiXml(buyerElectronicAddressSchemeIdValue: "0204", buyerElectronicAddressValue: "991-12345-XX"), "application/xml", PdfAssociatedFileRelationship.Data);

        PdfComplianceRequirement missingSellerAddress = AssertRequirement(
            PdfComplianceAnalyzer.Assess(PdfComplianceProfile.FacturX, missingSellerAddressOptions),
            "einvoice-xml-electronic-address",
            PdfComplianceRequirementStatus.Missing);
        PdfComplianceRequirement missingBuyerScheme = AssertRequirement(
            PdfComplianceAnalyzer.Assess(PdfComplianceProfile.FacturX, missingBuyerSchemeOptions),
            "einvoice-xml-electronic-address",
            PdfComplianceRequirementStatus.Missing);
        PdfComplianceRequirement invalidScheme = AssertRequirement(
            PdfComplianceAnalyzer.Assess(PdfComplianceProfile.FacturX, invalidSchemeOptions),
            "einvoice-xml-electronic-address",
            PdfComplianceRequirementStatus.Missing);
        PdfComplianceRequirement leitwegId = AssertRequirement(
            PdfComplianceAnalyzer.Assess(PdfComplianceProfile.FacturX, leitwegIdOptions),
            "einvoice-xml-electronic-address",
            PdfComplianceRequirementStatus.Satisfied);

        Assert.Contains("SellerTradeParty URIUniversalCommunication", missingSellerAddress.Diagnostic);
        Assert.Contains("BuyerTradeParty URIUniversalCommunication URIID schemeID", missingBuyerScheme.Diagnostic);
        Assert.Contains("Electronic Address Scheme", invalidScheme.Diagnostic);
        Assert.Contains("9999", invalidScheme.Diagnostic);
        Assert.Contains("Electronic Address Scheme", leitwegId.Diagnostic);
    }

    [Fact]
    public void FacturXReadinessRequiresCiiPartyTaxRegistrationEssentials() {
        var missingSellerTaxOptions = new PdfOptions()
            .SetPdfAIdentification(3, "B")
            .SetSrgbOutputIntent()
            .SetElectronicInvoiceMetadata(PdfElectronicInvoiceMetadata.FacturX("EN 16931"))
            .AddEmbeddedFile("factur-x.xml", CreateCiiXml(includeSellerTaxRegistration: false), "application/xml", PdfAssociatedFileRelationship.Data);
        var missingBuyerTaxOptions = new PdfOptions()
            .SetPdfAIdentification(3, "B")
            .SetSrgbOutputIntent()
            .SetElectronicInvoiceMetadata(PdfElectronicInvoiceMetadata.FacturX("EN 16931"))
            .AddEmbeddedFile("factur-x.xml", CreateCiiXml(includeBuyerTaxRegistration: false), "application/xml", PdfAssociatedFileRelationship.Data);

        PdfComplianceRequirement missingSellerTax = AssertRequirement(
            PdfComplianceAnalyzer.Assess(PdfComplianceProfile.FacturX, missingSellerTaxOptions),
            "einvoice-xml-party-tax-registration",
            PdfComplianceRequirementStatus.Missing);
        PdfComplianceRequirement missingBuyerTax = AssertRequirement(
            PdfComplianceAnalyzer.Assess(PdfComplianceProfile.FacturX, missingBuyerTaxOptions),
            "einvoice-xml-party-tax-registration",
            PdfComplianceRequirementStatus.Satisfied);

        Assert.Contains("SellerTradeParty SpecifiedTaxRegistration ID", missingSellerTax.Diagnostic);
        Assert.Contains("category-specific", missingBuyerTax.Diagnostic);
    }

    [Fact]
    public void FacturXReadinessRequiresCiiPartyTaxRegistrationSchemeMetadata() {
        var missingSellerSchemeOptions = new PdfOptions()
            .SetPdfAIdentification(3, "B")
            .SetSrgbOutputIntent()
            .SetElectronicInvoiceMetadata(PdfElectronicInvoiceMetadata.FacturX("EN 16931"))
            .AddEmbeddedFile("factur-x.xml", CreateCiiXml(includeSellerTaxRegistrationSchemeId: false), "application/xml", PdfAssociatedFileRelationship.Data);
        var missingBuyerSchemeOptions = new PdfOptions()
            .SetPdfAIdentification(3, "B")
            .SetSrgbOutputIntent()
            .SetElectronicInvoiceMetadata(PdfElectronicInvoiceMetadata.FacturX("EN 16931"))
            .AddEmbeddedFile("factur-x.xml", CreateCiiXml(includeBuyerTaxRegistrationSchemeId: false), "application/xml", PdfAssociatedFileRelationship.Data);
        var validSchemeOptions = new PdfOptions()
            .SetPdfAIdentification(3, "B")
            .SetSrgbOutputIntent()
            .SetElectronicInvoiceMetadata(PdfElectronicInvoiceMetadata.FacturX("EN 16931"))
            .AddEmbeddedFile("factur-x.xml", CreateCiiXml(), "application/xml", PdfAssociatedFileRelationship.Data);

        PdfComplianceRequirement missingSellerScheme = AssertRequirement(
            PdfComplianceAnalyzer.Assess(PdfComplianceProfile.FacturX, missingSellerSchemeOptions),
            "einvoice-xml-party-tax-registration-scheme",
            PdfComplianceRequirementStatus.Missing);
        PdfComplianceRequirement missingBuyerScheme = AssertRequirement(
            PdfComplianceAnalyzer.Assess(PdfComplianceProfile.FacturX, missingBuyerSchemeOptions),
            "einvoice-xml-party-tax-registration-scheme",
            PdfComplianceRequirementStatus.Missing);
        PdfComplianceRequirement validScheme = AssertRequirement(
            PdfComplianceAnalyzer.Assess(PdfComplianceProfile.FacturX, validSchemeOptions),
            "einvoice-xml-party-tax-registration-scheme",
            PdfComplianceRequirementStatus.Satisfied);

        Assert.Contains("SellerTradeParty SpecifiedTaxRegistration ID schemeID", missingSellerScheme.Diagnostic);
        Assert.Contains("BuyerTradeParty SpecifiedTaxRegistration ID schemeID", missingBuyerScheme.Diagnostic);
        Assert.Contains("schemeID metadata", validScheme.Diagnostic);
    }


}
