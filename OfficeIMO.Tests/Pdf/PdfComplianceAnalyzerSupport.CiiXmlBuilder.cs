using System.Text;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public partial class PdfComplianceAnalyzerTests {
    private static byte[] CreateGrossLinePriceCiiXml() {
        string xml = Encoding.UTF8.GetString(CreateCiiXml());
        xml = xml.Replace("NetPriceProductTradePrice", "GrossPriceProductTradePrice");
        return Encoding.UTF8.GetBytes(xml);
    }

    private static byte[] CreateCiiXml(
        string? profileContextId = "urn:factur-x.eu:1p0:en16931",
        bool includeDocumentHeader = true,
        bool includeSupplyChainTradeTransaction = true,
        bool includeTradeTransactionEssentials = true,
        bool includeSellerTradeParty = true,
        bool includeBuyerTradeParty = true,
        bool includeSellerName = true,
        bool includeSellerCountryId = true,
        bool includeSellerTaxRegistration = true,
        bool includeSellerElectronicAddress = true,
        bool includeSellerElectronicAddressSchemeId = true,
        bool includeBuyerName = true,
        bool includeBuyerCountryId = true,
        bool includeBuyerTaxRegistration = true,
        bool includeSellerTaxRegistrationSchemeId = true,
        bool includeBuyerTaxRegistrationSchemeId = true,
        bool includeBuyerElectronicAddress = true,
        bool includeBuyerElectronicAddressSchemeId = true,
        bool includePayableAmount = true,
        bool includeLineItem = true,
        bool includeLineItemProductName = true,
        bool includeLineTradeAgreement = true,
        bool includeLinePriceChargeAmount = true,
        bool includeLineTradeTax = true,
        bool includeLineTradeTaxTypeCode = true,
        bool includeLineTradeTaxCategoryCode = true,
        bool includeLineTradeTaxRate = true,
        bool includeLineTotalAmount = true,
        bool includeLineBilledQuantityUnitCode = true,
        bool includeInvoiceCurrencyCode = true,
        bool includeApplicableTradeTax = true,
        bool includeTradeTaxTypeCode = true,
        bool includeTradeTaxCategoryCode = true,
        bool includeTradeTaxRate = true,
        bool includeTradeTaxBasisAmount = true,
        bool includeTradeTaxCalculatedAmount = true,
        bool includeTaxTotals = true,
        bool includeAllowanceTotalAmount = false,
        bool includeChargeTotalAmount = false,
        bool includeDuePayableAmount = false,
        bool includePaidAmount = false,
        bool includeRoundingAmount = false,
        bool includePaymentMeans = true,
        bool includePaymentMeansTypeCode = true,
        bool includeCreditorAccount = true,
        bool includeCreditorAccountId = true,
        bool useCreditorProprietaryAccountId = false,
        string paymentMeansTypeCodeValue = "58",
        bool includePaymentTerms = true,
        bool includePaymentTermsDescription = true,
        bool includePaymentTermsDueDate = true,
        string documentTypeCodeValue = "380",
        string issueDateTimeFormat = "102",
        string issueDateTimeValue = "20260603",
        string dueDateTimeFormat = "102",
        string dueDateTimeValue = "20260703",
        string lineTotalAmount = "100.00",
        string linePriceChargeAmountValue = "100.00",
        string? linePriceBasisQuantityValue = null,
        string lineBilledQuantityValue = "1",
        string lineBilledQuantityUnitCodeValue = "C62",
        string taxBasisTotalAmount = "100.00",
        string taxTotalAmount = "23.00",
        string grandTotalAmount = "123.00",
        string allowanceTotalAmount = "0.00",
        string chargeTotalAmount = "0.00",
        string duePayableAmount = "123.00",
        string paidAmount = "0.00",
        string roundingAmount = "0.00",
        string headerTradeTaxCategoryCodeValue = "S",
        string headerTradeTaxTypeCodeValue = "VAT",
        string headerTradeTaxRateValue = "23",
        string headerTradeTaxBasisAmountValue = "100.00",
        string headerTradeTaxCalculatedAmountValue = "23.00",
        string? headerTradeTaxExemptionReasonValue = null,
        string? headerTradeTaxExemptionReasonCodeValue = null,
        string lineTradeTaxTypeCodeValue = "VAT",
        string lineTradeTaxCategoryCodeValue = "S",
        string lineTradeTaxRateValue = "23",
        string invoiceCurrencyCodeValue = "EUR",
        string? amountCurrencyId = "EUR",
        string creditorAccountIban = "PL61109010140000071219812874",
        string creditorProprietaryAccountId = "ACCOUNT-001",
        string sellerCountryIdValue = "PL",
        string buyerCountryIdValue = "DE",
        string sellerElectronicAddressValue = "PL1234567890",
        string sellerElectronicAddressSchemeIdValue = "9945",
        string buyerElectronicAddressValue = "DE123456789",
        string buyerElectronicAddressSchemeIdValue = "9930") {
        string context = profileContextId == null
            ? "<rsm:ExchangedDocumentContext />"
            : "<rsm:ExchangedDocumentContext>" +
              "<ram:GuidelineSpecifiedDocumentContextParameter>" +
              "<ram:ID>" + profileContextId + "</ram:ID>" +
              "</ram:GuidelineSpecifiedDocumentContextParameter>" +
              "</rsm:ExchangedDocumentContext>";
        string document = includeDocumentHeader
            ? "<rsm:ExchangedDocument>" +
              "<ram:ID>INV-2026-0001</ram:ID>" +
              "<ram:TypeCode>" + documentTypeCodeValue + "</ram:TypeCode>" +
              "<ram:IssueDateTime><udt:DateTimeString format=\"" + issueDateTimeFormat + "\">" + issueDateTimeValue + "</udt:DateTimeString></ram:IssueDateTime>" +
              "</rsm:ExchangedDocument>"
            : "<rsm:ExchangedDocument />";
        string transaction = CreateSupplyChainTradeTransaction(
            includeSupplyChainTradeTransaction,
            includeTradeTransactionEssentials,
            includeSellerTradeParty,
            includeBuyerTradeParty,
            includeSellerName,
            includeSellerCountryId,
            includeSellerTaxRegistration,
            includeSellerTaxRegistrationSchemeId,
            includeSellerElectronicAddress,
            includeSellerElectronicAddressSchemeId,
            includeBuyerName,
            includeBuyerCountryId,
            includeBuyerTaxRegistration,
            includeBuyerTaxRegistrationSchemeId,
            includeBuyerElectronicAddress,
            includeBuyerElectronicAddressSchemeId,
            includePayableAmount,
            includeLineItem,
            includeLineItemProductName,
            includeLineTradeAgreement,
            includeLinePriceChargeAmount,
            includeLineTradeTax,
            includeLineTradeTaxTypeCode,
            includeLineTradeTaxCategoryCode,
            includeLineTradeTaxRate,
            includeLineTotalAmount,
            includeLineBilledQuantityUnitCode,
            includeInvoiceCurrencyCode,
            includeApplicableTradeTax,
            includeTradeTaxTypeCode,
            includeTradeTaxCategoryCode,
            includeTradeTaxRate,
            includeTradeTaxBasisAmount,
            includeTradeTaxCalculatedAmount,
            includeTaxTotals,
            includeAllowanceTotalAmount,
            includeChargeTotalAmount,
            includeDuePayableAmount,
            includePaidAmount,
            includeRoundingAmount,
            includePaymentMeans,
            includePaymentMeansTypeCode,
            paymentMeansTypeCodeValue,
            includeCreditorAccount,
            includeCreditorAccountId,
            useCreditorProprietaryAccountId,
            includePaymentTerms,
            includePaymentTermsDescription,
            includePaymentTermsDueDate,
            dueDateTimeFormat,
            dueDateTimeValue,
            lineTotalAmount,
            linePriceChargeAmountValue,
            linePriceBasisQuantityValue,
            lineBilledQuantityValue,
            lineBilledQuantityUnitCodeValue,
            taxBasisTotalAmount,
            taxTotalAmount,
            grandTotalAmount,
            allowanceTotalAmount,
            chargeTotalAmount,
            duePayableAmount,
            paidAmount,
            roundingAmount,
            headerTradeTaxCategoryCodeValue,
            headerTradeTaxTypeCodeValue,
            headerTradeTaxRateValue,
            headerTradeTaxBasisAmountValue,
            headerTradeTaxCalculatedAmountValue,
            headerTradeTaxExemptionReasonValue,
            headerTradeTaxExemptionReasonCodeValue,
            lineTradeTaxTypeCodeValue,
            lineTradeTaxCategoryCodeValue,
            lineTradeTaxRateValue,
            invoiceCurrencyCodeValue,
            amountCurrencyId,
            creditorAccountIban,
            creditorProprietaryAccountId,
            sellerCountryIdValue,
            buyerCountryIdValue,
            sellerElectronicAddressValue,
            sellerElectronicAddressSchemeIdValue,
            buyerElectronicAddressValue,
            buyerElectronicAddressSchemeIdValue);
        return Encoding.UTF8.GetBytes(
            "<?xml version=\"1.0\" encoding=\"UTF-8\"?>" +
            "<rsm:CrossIndustryInvoice xmlns:rsm=\"urn:un:unece:uncefact:data:standard:CrossIndustryInvoice:100\" xmlns:ram=\"urn:un:unece:uncefact:data:standard:ReusableAggregateBusinessInformationEntity:100\" xmlns:udt=\"urn:un:unece:uncefact:data:standard:UnqualifiedDataType:100\">" +
            context +
            document +
            transaction +
            "</rsm:CrossIndustryInvoice>");
    }

    private static string CreateSupplyChainTradeTransaction(
        bool includeSupplyChainTradeTransaction,
        bool includeTradeTransactionEssentials,
        bool includeSellerTradeParty,
        bool includeBuyerTradeParty,
        bool includeSellerName,
        bool includeSellerCountryId,
        bool includeSellerTaxRegistration,
        bool includeSellerTaxRegistrationSchemeId,
        bool includeSellerElectronicAddress,
        bool includeSellerElectronicAddressSchemeId,
        bool includeBuyerName,
        bool includeBuyerCountryId,
        bool includeBuyerTaxRegistration,
        bool includeBuyerTaxRegistrationSchemeId,
        bool includeBuyerElectronicAddress,
        bool includeBuyerElectronicAddressSchemeId,
        bool includePayableAmount,
        bool includeLineItem,
        bool includeLineItemProductName,
        bool includeLineTradeAgreement,
        bool includeLinePriceChargeAmount,
        bool includeLineTradeTax,
        bool includeLineTradeTaxTypeCode,
        bool includeLineTradeTaxCategoryCode,
        bool includeLineTradeTaxRate,
        bool includeLineTotalAmount,
        bool includeLineBilledQuantityUnitCode,
        bool includeInvoiceCurrencyCode,
        bool includeApplicableTradeTax,
        bool includeTradeTaxTypeCode,
        bool includeTradeTaxCategoryCode,
        bool includeTradeTaxRate,
        bool includeTradeTaxBasisAmount,
        bool includeTradeTaxCalculatedAmount,
        bool includeTaxTotals,
        bool includeAllowanceTotalAmount,
        bool includeChargeTotalAmount,
        bool includeDuePayableAmount,
        bool includePaidAmount,
        bool includeRoundingAmount,
        bool includePaymentMeans,
        bool includePaymentMeansTypeCode,
        string paymentMeansTypeCodeValue,
        bool includeCreditorAccount,
        bool includeCreditorAccountId,
        bool useCreditorProprietaryAccountId,
        bool includePaymentTerms,
        bool includePaymentTermsDescription,
        bool includePaymentTermsDueDate,
        string dueDateTimeFormat,
        string dueDateTimeValue,
        string lineTotalAmountValue,
        string linePriceChargeAmountValue,
        string? linePriceBasisQuantityValue,
        string lineBilledQuantityValue,
        string lineBilledQuantityUnitCodeValue,
        string taxBasisTotalAmountValue,
        string taxTotalAmountValue,
        string grandTotalAmountValue,
        string allowanceTotalAmountValue,
        string chargeTotalAmountValue,
        string duePayableAmountValue,
        string paidAmountValue,
        string roundingAmountValue,
        string headerTradeTaxCategoryCodeValue,
        string headerTradeTaxTypeCodeValue,
        string headerTradeTaxRateValue,
        string headerTradeTaxBasisAmountValue,
        string headerTradeTaxCalculatedAmountValue,
        string? headerTradeTaxExemptionReasonValue,
        string? headerTradeTaxExemptionReasonCodeValue,
        string lineTradeTaxTypeCodeValue,
        string lineTradeTaxCategoryCodeValue,
        string lineTradeTaxRateValue,
        string invoiceCurrencyCodeValue,
        string? amountCurrencyId,
        string creditorAccountIban,
        string creditorProprietaryAccountId,
        string sellerCountryIdValue,
        string buyerCountryIdValue,
        string sellerElectronicAddressValue,
        string sellerElectronicAddressSchemeIdValue,
        string buyerElectronicAddressValue,
        string buyerElectronicAddressSchemeIdValue) {
        if (!includeSupplyChainTradeTransaction) {
            return string.Empty;
        }

        if (!includeTradeTransactionEssentials) {
            return "<rsm:SupplyChainTradeTransaction />";
        }

        string seller = includeSellerTradeParty
            ? CreateTradeParty("SellerTradeParty", "OfficeIMO Seller", sellerCountryIdValue, "PL1234567890", includeSellerName, includeSellerCountryId, includeSellerTaxRegistration, includeSellerTaxRegistrationSchemeId, includeSellerElectronicAddress, includeSellerElectronicAddressSchemeId, sellerElectronicAddressValue, sellerElectronicAddressSchemeIdValue)
            : string.Empty;
        string buyer = includeBuyerTradeParty
            ? CreateTradeParty("BuyerTradeParty", "OfficeIMO Buyer", buyerCountryIdValue, "DE123456789", includeBuyerName, includeBuyerCountryId, includeBuyerTaxRegistration, includeBuyerTaxRegistrationSchemeId, includeBuyerElectronicAddress, includeBuyerElectronicAddressSchemeId, buyerElectronicAddressValue, buyerElectronicAddressSchemeIdValue)
            : string.Empty;
        string amount = includePayableAmount
            ? "<ram:GrandTotalAmount" + CurrencyAttribute(amountCurrencyId) + ">" + grandTotalAmountValue + "</ram:GrandTotalAmount>"
            : string.Empty;
        string duePayable = includeDuePayableAmount
            ? "<ram:DuePayableAmount" + CurrencyAttribute(amountCurrencyId) + ">" + duePayableAmountValue + "</ram:DuePayableAmount>"
            : string.Empty;
        string paid = includePaidAmount
            ? "<ram:PaidAmount" + CurrencyAttribute(amountCurrencyId) + ">" + paidAmountValue + "</ram:PaidAmount>"
            : string.Empty;
        string rounding = includeRoundingAmount
            ? "<ram:RoundingAmount" + CurrencyAttribute(amountCurrencyId) + ">" + roundingAmountValue + "</ram:RoundingAmount>"
            : string.Empty;
        string lineItem = CreateIncludedSupplyChainTradeLineItem(includeLineItem, includeLineItemProductName, includeLineTradeAgreement, includeLinePriceChargeAmount, includeLineTradeTax, includeLineTradeTaxTypeCode, includeLineTradeTaxCategoryCode, includeLineTradeTaxRate, includeLineTotalAmount, includeLineBilledQuantityUnitCode, lineTotalAmountValue, linePriceChargeAmountValue, linePriceBasisQuantityValue, lineBilledQuantityValue, lineBilledQuantityUnitCodeValue, lineTradeTaxTypeCodeValue, lineTradeTaxCategoryCodeValue, lineTradeTaxRateValue, amountCurrencyId);
        string currencyCode = includeInvoiceCurrencyCode
            ? "<ram:InvoiceCurrencyCode>" + invoiceCurrencyCodeValue + "</ram:InvoiceCurrencyCode>"
            : string.Empty;
        string tradeTax = CreateApplicableTradeTax(
            includeApplicableTradeTax,
            includeTradeTaxTypeCode,
            includeTradeTaxCategoryCode,
            includeTradeTaxRate,
            includeTradeTaxBasisAmount,
            includeTradeTaxCalculatedAmount,
            headerTradeTaxCategoryCodeValue,
            headerTradeTaxTypeCodeValue,
            headerTradeTaxRateValue,
            headerTradeTaxBasisAmountValue,
            headerTradeTaxCalculatedAmountValue,
            headerTradeTaxExemptionReasonValue,
            headerTradeTaxExemptionReasonCodeValue,
            amountCurrencyId);
        string taxTotals = includeTaxTotals
            ? "<ram:TaxBasisTotalAmount" + CurrencyAttribute(amountCurrencyId) + ">" + taxBasisTotalAmountValue + "</ram:TaxBasisTotalAmount>" +
              "<ram:TaxTotalAmount" + CurrencyAttribute(amountCurrencyId) + ">" + taxTotalAmountValue + "</ram:TaxTotalAmount>"
            : string.Empty;
        string allowanceTotal = includeAllowanceTotalAmount
            ? "<ram:AllowanceTotalAmount" + CurrencyAttribute(amountCurrencyId) + ">" + allowanceTotalAmountValue + "</ram:AllowanceTotalAmount>"
            : string.Empty;
        string chargeTotal = includeChargeTotalAmount
            ? "<ram:ChargeTotalAmount" + CurrencyAttribute(amountCurrencyId) + ">" + chargeTotalAmountValue + "</ram:ChargeTotalAmount>"
            : string.Empty;
        string paymentMeans = CreatePaymentMeans(includePaymentMeans, includePaymentMeansTypeCode, paymentMeansTypeCodeValue, includeCreditorAccount, includeCreditorAccountId, useCreditorProprietaryAccountId, creditorAccountIban, creditorProprietaryAccountId);
        string paymentTerms = CreatePaymentTerms(includePaymentTerms, includePaymentTermsDescription, includePaymentTermsDueDate, dueDateTimeFormat, dueDateTimeValue);
        return "<rsm:SupplyChainTradeTransaction>" +
               lineItem +
               "<ram:ApplicableHeaderTradeAgreement>" +
               seller +
               buyer +
               "</ram:ApplicableHeaderTradeAgreement>" +
               "<ram:ApplicableHeaderTradeSettlement>" +
                currencyCode +
                tradeTax +
                paymentMeans +
                paymentTerms +
                "<ram:SpecifiedTradeSettlementHeaderMonetarySummation>" +
               taxTotals +
               allowanceTotal +
               chargeTotal +
               amount +
               paid +
               rounding +
               duePayable +
               "</ram:SpecifiedTradeSettlementHeaderMonetarySummation>" +
               "</ram:ApplicableHeaderTradeSettlement>" +
               "</rsm:SupplyChainTradeTransaction>";
    }

    private static string CreatePaymentMeans(
        bool includePaymentMeans,
        bool includePaymentMeansTypeCode,
        string paymentMeansTypeCodeValue,
        bool includeCreditorAccount,
        bool includeCreditorAccountId,
        bool useCreditorProprietaryAccountId,
        string creditorAccountIban,
        string creditorProprietaryAccountId) {
        if (!includePaymentMeans) {
            return string.Empty;
        }

        string typeCode = includePaymentMeansTypeCode
            ? "<ram:TypeCode>" + paymentMeansTypeCodeValue + "</ram:TypeCode>"
            : string.Empty;
        string accountId = includeCreditorAccountId
            ? useCreditorProprietaryAccountId
                ? "<ram:ProprietaryID>" + creditorProprietaryAccountId + "</ram:ProprietaryID>"
                : "<ram:IBANID>" + creditorAccountIban + "</ram:IBANID>"
            : string.Empty;
        string account = includeCreditorAccount
            ? "<ram:PayeePartyCreditorFinancialAccount>" +
              accountId +
              "</ram:PayeePartyCreditorFinancialAccount>"
            : string.Empty;
        return "<ram:SpecifiedTradeSettlementPaymentMeans>" +
               typeCode +
               account +
               "</ram:SpecifiedTradeSettlementPaymentMeans>";
    }

    private static string CreatePaymentTerms(bool includePaymentTerms, bool includePaymentTermsDescription, bool includePaymentTermsDueDate, string dueDateTimeFormat, string dueDateTimeValue) {
        if (!includePaymentTerms) {
            return string.Empty;
        }

        string description = includePaymentTermsDescription
            ? "<ram:Description>Due within 30 days</ram:Description>"
            : string.Empty;
        string dueDate = includePaymentTermsDueDate
            ? "<ram:DueDateDateTime><udt:DateTimeString format=\"" + dueDateTimeFormat + "\">" + dueDateTimeValue + "</udt:DateTimeString></ram:DueDateDateTime>"
            : string.Empty;

        return "<ram:SpecifiedTradePaymentTerms>" +
               description +
               dueDate +
               "</ram:SpecifiedTradePaymentTerms>";
    }

    private static byte[] CreateTwoLineCiiXmlWithSecondLineMissingProductName() {
        string firstLine = CreateIncludedSupplyChainTradeLineItem(
            includeLineItem: true,
            includeLineItemProductName: true,
            includeLineTradeAgreement: true,
            includeLinePriceChargeAmount: true,
            includeLineTradeTax: true,
            includeLineTradeTaxTypeCode: true,
            includeLineTradeTaxCategoryCode: true,
            includeLineTradeTaxRate: true,
            includeLineTotalAmount: true,
            includeLineBilledQuantityUnitCode: true,
            lineTotalAmountValue: "100.00",
            linePriceChargeAmountValue: "100.00",
            linePriceBasisQuantityValue: null,
            lineBilledQuantityValue: "1",
            lineBilledQuantityUnitCodeValue: "C62",
            lineTradeTaxTypeCodeValue: "VAT",
            lineTradeTaxCategoryCodeValue: "S",
            lineTradeTaxRateValue: "23",
            amountCurrencyId: "EUR");
        string secondLine = CreateIncludedSupplyChainTradeLineItem(
            includeLineItem: true,
            includeLineItemProductName: false,
            includeLineTradeAgreement: true,
            includeLinePriceChargeAmount: true,
            includeLineTradeTax: true,
            includeLineTradeTaxTypeCode: true,
            includeLineTradeTaxCategoryCode: true,
            includeLineTradeTaxRate: true,
            includeLineTotalAmount: true,
            includeLineBilledQuantityUnitCode: true,
            lineTotalAmountValue: "100.00",
            linePriceChargeAmountValue: "100.00",
            linePriceBasisQuantityValue: null,
            lineBilledQuantityValue: "1",
            lineBilledQuantityUnitCodeValue: "C62",
            lineTradeTaxTypeCodeValue: "VAT",
            lineTradeTaxCategoryCodeValue: "S",
            lineTradeTaxRateValue: "23",
            amountCurrencyId: "EUR")
            .Replace("<ram:LineID>1</ram:LineID>", "<ram:LineID>2</ram:LineID>");
        string xml = Encoding.UTF8.GetString(CreateCiiXml())
            .Replace(firstLine, firstLine + secondLine);
        return Encoding.UTF8.GetBytes(xml);
    }

    private static byte[] CreateTwoLineCiiXmlWithSecondLineMissingPriceCharge() {
        string firstLine = CreateIncludedSupplyChainTradeLineItem(
            includeLineItem: true,
            includeLineItemProductName: true,
            includeLineTradeAgreement: true,
            includeLinePriceChargeAmount: true,
            includeLineTradeTax: true,
            includeLineTradeTaxTypeCode: true,
            includeLineTradeTaxCategoryCode: true,
            includeLineTradeTaxRate: true,
            includeLineTotalAmount: true,
            includeLineBilledQuantityUnitCode: true,
            lineTotalAmountValue: "100.00",
            linePriceChargeAmountValue: "100.00",
            linePriceBasisQuantityValue: null,
            lineBilledQuantityValue: "1",
            lineBilledQuantityUnitCodeValue: "C62",
            lineTradeTaxTypeCodeValue: "VAT",
            lineTradeTaxCategoryCodeValue: "S",
            lineTradeTaxRateValue: "23",
            amountCurrencyId: "EUR");
        string secondLine = CreateIncludedSupplyChainTradeLineItem(
            includeLineItem: true,
            includeLineItemProductName: true,
            includeLineTradeAgreement: true,
            includeLinePriceChargeAmount: false,
            includeLineTradeTax: true,
            includeLineTradeTaxTypeCode: true,
            includeLineTradeTaxCategoryCode: true,
            includeLineTradeTaxRate: true,
            includeLineTotalAmount: true,
            includeLineBilledQuantityUnitCode: true,
            lineTotalAmountValue: "100.00",
            linePriceChargeAmountValue: "100.00",
            linePriceBasisQuantityValue: null,
            lineBilledQuantityValue: "1",
            lineBilledQuantityUnitCodeValue: "C62",
            lineTradeTaxTypeCodeValue: "VAT",
            lineTradeTaxCategoryCodeValue: "S",
            lineTradeTaxRateValue: "23",
            amountCurrencyId: "EUR")
            .Replace("<ram:LineID>1</ram:LineID>", "<ram:LineID>2</ram:LineID>");
        string xml = Encoding.UTF8.GetString(CreateCiiXml())
            .Replace(firstLine, firstLine + secondLine);
        return Encoding.UTF8.GetBytes(xml);
    }

    private static byte[] CreateTwoLineCiiXmlWithSecondLineMissingLineTax() {
        string firstLine = CreateIncludedSupplyChainTradeLineItem(
            includeLineItem: true,
            includeLineItemProductName: true,
            includeLineTradeAgreement: true,
            includeLinePriceChargeAmount: true,
            includeLineTradeTax: true,
            includeLineTradeTaxTypeCode: true,
            includeLineTradeTaxCategoryCode: true,
            includeLineTradeTaxRate: true,
            includeLineTotalAmount: true,
            includeLineBilledQuantityUnitCode: true,
            lineTotalAmountValue: "100.00",
            linePriceChargeAmountValue: "100.00",
            linePriceBasisQuantityValue: null,
            lineBilledQuantityValue: "1",
            lineBilledQuantityUnitCodeValue: "C62",
            lineTradeTaxTypeCodeValue: "VAT",
            lineTradeTaxCategoryCodeValue: "S",
            lineTradeTaxRateValue: "23",
            amountCurrencyId: "EUR");
        string secondLine = CreateIncludedSupplyChainTradeLineItem(
            includeLineItem: true,
            includeLineItemProductName: true,
            includeLineTradeAgreement: true,
            includeLinePriceChargeAmount: true,
            includeLineTradeTax: false,
            includeLineTradeTaxTypeCode: true,
            includeLineTradeTaxCategoryCode: true,
            includeLineTradeTaxRate: true,
            includeLineTotalAmount: true,
            includeLineBilledQuantityUnitCode: true,
            lineTotalAmountValue: "100.00",
            linePriceChargeAmountValue: "100.00",
            linePriceBasisQuantityValue: null,
            lineBilledQuantityValue: "1",
            lineBilledQuantityUnitCodeValue: "C62",
            lineTradeTaxTypeCodeValue: "VAT",
            lineTradeTaxCategoryCodeValue: "S",
            lineTradeTaxRateValue: "23",
            amountCurrencyId: "EUR")
            .Replace("<ram:LineID>1</ram:LineID>", "<ram:LineID>2</ram:LineID>");
        string xml = Encoding.UTF8.GetString(CreateCiiXml())
            .Replace(firstLine, firstLine + secondLine);
        return Encoding.UTF8.GetBytes(xml);
    }

    private static byte[] CreateCiiXmlWithSecondPaymentMeansMissingTypeCode() {
        string xml = Encoding.UTF8.GetString(CreateCiiXml());
        string secondPaymentMeans = CreatePaymentMeans(
            true,
            false,
            "58",
            true,
            true,
            false,
            "PL61109010140000071219812874",
            "ACCOUNT-001");
        return Encoding.UTF8.GetBytes(xml.Replace("<ram:SpecifiedTradePaymentTerms>", secondPaymentMeans + "<ram:SpecifiedTradePaymentTerms>"));
    }

    private static byte[] CreateCiiXmlWithTransferPaymentMeansMissingOwnAccount() {
        string xml = Encoding.UTF8.GetString(CreateCiiXml(includeCreditorAccount: false));
        string cashPaymentMeansWithAccount = CreatePaymentMeans(
            true,
            true,
            "10",
            true,
            true,
            false,
            "PL61109010140000071219812874",
            "ACCOUNT-001");
        return Encoding.UTF8.GetBytes(xml.Replace("<ram:SpecifiedTradePaymentTerms>", cashPaymentMeansWithAccount + "<ram:SpecifiedTradePaymentTerms>"));
    }

    private static byte[] CreateCiiXmlWithLineAllowance() {
        string xml = Encoding.UTF8.GetString(CreateCiiXml(
            lineTotalAmount: "95.00",
            linePriceChargeAmountValue: "10.00",
            lineBilledQuantityValue: "10",
            taxBasisTotalAmount: "95.00",
            taxTotalAmount: "21.85",
            grandTotalAmount: "116.85",
            headerTradeTaxBasisAmountValue: "95.00",
            headerTradeTaxCalculatedAmountValue: "21.85"));
        string allowance =
            "<ram:SpecifiedTradeAllowanceCharge>" +
            "<ram:ChargeIndicator><udt:Indicator>false</udt:Indicator></ram:ChargeIndicator>" +
            "<ram:ActualAmount currencyID=\"EUR\">5.00</ram:ActualAmount>" +
            "</ram:SpecifiedTradeAllowanceCharge>";
        return Encoding.UTF8.GetBytes(xml.Replace("<ram:SpecifiedTradeSettlementLineMonetarySummation>", allowance + "<ram:SpecifiedTradeSettlementLineMonetarySummation>"));
    }

    private static string CreateApplicableTradeTax(
        bool includeApplicableTradeTax,
        bool includeTradeTaxTypeCode,
        bool includeTradeTaxCategoryCode,
        bool includeTradeTaxRate,
        bool includeTradeTaxBasisAmount,
        bool includeTradeTaxCalculatedAmount,
        string headerTradeTaxCategoryCodeValue,
        string headerTradeTaxTypeCodeValue,
        string headerTradeTaxRateValue,
        string headerTradeTaxBasisAmountValue,
        string headerTradeTaxCalculatedAmountValue,
        string? headerTradeTaxExemptionReasonValue,
        string? headerTradeTaxExemptionReasonCodeValue,
        string? amountCurrencyId) {
        if (!includeApplicableTradeTax) {
            return string.Empty;
        }

        string calculatedAmount = includeTradeTaxCalculatedAmount
            ? "<ram:CalculatedAmount" + CurrencyAttribute(amountCurrencyId) + ">" + headerTradeTaxCalculatedAmountValue + "</ram:CalculatedAmount>"
            : string.Empty;
        string typeCode = includeTradeTaxTypeCode
            ? "<ram:TypeCode>" + headerTradeTaxTypeCodeValue + "</ram:TypeCode>"
            : string.Empty;
        string basisAmount = includeTradeTaxBasisAmount
            ? "<ram:BasisAmount" + CurrencyAttribute(amountCurrencyId) + ">" + headerTradeTaxBasisAmountValue + "</ram:BasisAmount>"
            : string.Empty;
        string categoryCode = includeTradeTaxCategoryCode
            ? "<ram:CategoryCode>" + headerTradeTaxCategoryCodeValue + "</ram:CategoryCode>"
            : string.Empty;
        string rate = includeTradeTaxRate
            ? "<ram:RateApplicablePercent>" + headerTradeTaxRateValue + "</ram:RateApplicablePercent>"
            : string.Empty;
        string exemptionReason = headerTradeTaxExemptionReasonValue == null
            ? string.Empty
            : "<ram:ExemptionReason>" + headerTradeTaxExemptionReasonValue + "</ram:ExemptionReason>";
        string exemptionReasonCode = headerTradeTaxExemptionReasonCodeValue == null
            ? string.Empty
            : "<ram:ExemptionReasonCode>" + headerTradeTaxExemptionReasonCodeValue + "</ram:ExemptionReasonCode>";

        return "<ram:ApplicableTradeTax>" +
               calculatedAmount +
               typeCode +
               basisAmount +
               categoryCode +
               rate +
               exemptionReason +
               exemptionReasonCode +
               "</ram:ApplicableTradeTax>";
    }

    private static string CreateTradeParty(
        string elementName,
        string nameValue,
        string countryId,
        string taxId,
        bool includeName,
        bool includeCountryId,
        bool includeTaxRegistration,
        bool includeTaxRegistrationSchemeId,
        bool includeElectronicAddress,
        bool includeElectronicAddressSchemeId,
        string electronicAddressValue,
        string electronicAddressSchemeIdValue) {
        string name = includeName
            ? "<ram:Name>" + nameValue + "</ram:Name>"
            : string.Empty;
        string taxRegistrationSchemeId = includeTaxRegistrationSchemeId
            ? " schemeID=\"VA\""
            : string.Empty;
        string taxRegistration = includeTaxRegistration
            ? "<ram:SpecifiedTaxRegistration><ram:ID" + taxRegistrationSchemeId + ">" + taxId + "</ram:ID></ram:SpecifiedTaxRegistration>"
            : string.Empty;
        string electronicAddressSchemeId = includeElectronicAddressSchemeId
            ? " schemeID=\"" + electronicAddressSchemeIdValue + "\""
            : string.Empty;
        string electronicAddress = includeElectronicAddress
            ? "<ram:URIUniversalCommunication><ram:URIID" + electronicAddressSchemeId + ">" + electronicAddressValue + "</ram:URIID></ram:URIUniversalCommunication>"
            : string.Empty;
        string country = includeCountryId
            ? "<ram:CountryID>" + countryId + "</ram:CountryID>"
            : string.Empty;
        string address = "<ram:PostalTradeAddress>" +
                         "<ram:PostcodeCode>00-001</ram:PostcodeCode>" +
                         "<ram:LineOne>Compliance Street 1</ram:LineOne>" +
                         "<ram:CityName>Warsaw</ram:CityName>" +
                         country +
                         "</ram:PostalTradeAddress>";
        return "<ram:" + elementName + ">" +
               name +
               taxRegistration +
               electronicAddress +
               address +
               "</ram:" + elementName + ">";
    }

    private static string CreateIncludedSupplyChainTradeLineItem(
        bool includeLineItem,
        bool includeLineItemProductName,
        bool includeLineTradeAgreement,
        bool includeLinePriceChargeAmount,
        bool includeLineTradeTax,
        bool includeLineTradeTaxTypeCode,
        bool includeLineTradeTaxCategoryCode,
        bool includeLineTradeTaxRate,
        bool includeLineTotalAmount,
        bool includeLineBilledQuantityUnitCode,
        string lineTotalAmountValue,
        string linePriceChargeAmountValue,
        string? linePriceBasisQuantityValue,
        string lineBilledQuantityValue,
        string lineBilledQuantityUnitCodeValue,
        string lineTradeTaxTypeCodeValue,
        string lineTradeTaxCategoryCodeValue,
        string lineTradeTaxRateValue,
        string? amountCurrencyId) {
        if (!includeLineItem) {
            return string.Empty;
        }

        string productName = includeLineItemProductName
            ? "<ram:Name>OfficeIMO PDF compliance work</ram:Name>"
            : string.Empty;
        string linePriceChargeAmount = includeLinePriceChargeAmount
            ? "<ram:ChargeAmount" + CurrencyAttribute(amountCurrencyId) + ">" + linePriceChargeAmountValue + "</ram:ChargeAmount>"
            : string.Empty;
        string linePriceBasisQuantity = linePriceBasisQuantityValue == null
            ? string.Empty
            : "<ram:BasisQuantity>" + linePriceBasisQuantityValue + "</ram:BasisQuantity>";
        string lineTradeAgreement = includeLineTradeAgreement
            ? "<ram:SpecifiedLineTradeAgreement>" +
              "<ram:NetPriceProductTradePrice>" +
              linePriceChargeAmount +
              linePriceBasisQuantity +
              "</ram:NetPriceProductTradePrice>" +
              "</ram:SpecifiedLineTradeAgreement>"
            : string.Empty;
        string lineTotalAmount = includeLineTotalAmount
            ? "<ram:LineTotalAmount" + CurrencyAttribute(amountCurrencyId) + ">" + lineTotalAmountValue + "</ram:LineTotalAmount>"
            : string.Empty;
        string lineTradeTax = CreateLineApplicableTradeTax(includeLineTradeTax, includeLineTradeTaxTypeCode, includeLineTradeTaxCategoryCode, includeLineTradeTaxRate, lineTradeTaxTypeCodeValue, lineTradeTaxCategoryCodeValue, lineTradeTaxRateValue);
        string billedQuantityUnitCode = includeLineBilledQuantityUnitCode
            ? " unitCode=\"" + lineBilledQuantityUnitCodeValue + "\""
            : string.Empty;
        return "<ram:IncludedSupplyChainTradeLineItem>" +
               "<ram:AssociatedDocumentLineDocument><ram:LineID>1</ram:LineID></ram:AssociatedDocumentLineDocument>" +
               "<ram:SpecifiedTradeProduct>" + productName + "</ram:SpecifiedTradeProduct>" +
               lineTradeAgreement +
               "<ram:SpecifiedLineTradeDelivery><ram:BilledQuantity" + billedQuantityUnitCode + ">" + lineBilledQuantityValue + "</ram:BilledQuantity></ram:SpecifiedLineTradeDelivery>" +
               "<ram:SpecifiedLineTradeSettlement>" +
               lineTradeTax +
               "<ram:SpecifiedTradeSettlementLineMonetarySummation>" +
               lineTotalAmount +
               "</ram:SpecifiedTradeSettlementLineMonetarySummation>" +
               "</ram:SpecifiedLineTradeSettlement>" +
               "</ram:IncludedSupplyChainTradeLineItem>";
    }

    private static string CurrencyAttribute(string? amountCurrencyId) {
        return amountCurrencyId == null
            ? string.Empty
            : " currencyID=\"" + amountCurrencyId + "\"";
    }

    private static string CreateLineApplicableTradeTax(
        bool includeLineTradeTax,
        bool includeLineTradeTaxTypeCode,
        bool includeLineTradeTaxCategoryCode,
        bool includeLineTradeTaxRate,
        string lineTradeTaxTypeCodeValue,
        string lineTradeTaxCategoryCodeValue,
        string lineTradeTaxRateValue) {
        if (!includeLineTradeTax) {
            return string.Empty;
        }

        string typeCode = includeLineTradeTaxTypeCode
            ? "<ram:TypeCode>" + lineTradeTaxTypeCodeValue + "</ram:TypeCode>"
            : string.Empty;
        string categoryCode = includeLineTradeTaxCategoryCode
            ? "<ram:CategoryCode>" + lineTradeTaxCategoryCodeValue + "</ram:CategoryCode>"
            : string.Empty;
        string rate = includeLineTradeTaxRate
            ? "<ram:RateApplicablePercent>" + lineTradeTaxRateValue + "</ram:RateApplicablePercent>"
            : string.Empty;

        return "<ram:ApplicableTradeTax>" +
               typeCode +
               categoryCode +
               rate +
               "</ram:ApplicableTradeTax>";
    }


}
