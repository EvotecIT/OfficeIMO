using System.Text;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public partial class PdfComplianceAnalyzerTests {
    private static byte[] AddHeaderTradeTax(byte[] ciiXml, string categoryCode, bool includeRate) {
        string xml = Encoding.UTF8.GetString(ciiXml);
        string tradeTax = CreateApplicableTradeTax(
            true,
            true,
            true,
            includeRate,
            true,
            true,
            categoryCode,
            "VAT",
            includeRate ? "23" : "0",
            "0.00",
            "0.00",
            string.Equals(categoryCode, "O", StringComparison.Ordinal) ? "Not subject to VAT" : null,
            null,
            "EUR");
        return Encoding.UTF8.GetBytes(xml.Replace("</ram:ApplicableHeaderTradeSettlement>", tradeTax + "</ram:ApplicableHeaderTradeSettlement>"));
    }

    private static byte[] AddHeaderTradeTax(byte[] ciiXml, string categoryCode, string rateValue, string basisAmount, string calculatedAmount) {
        string xml = Encoding.UTF8.GetString(ciiXml);
        string tradeTax = CreateApplicableTradeTax(
            true,
            true,
            true,
            true,
            true,
            true,
            categoryCode,
            "VAT",
            rateValue,
            basisAmount,
            calculatedAmount,
            null,
            null,
            "EUR");
        return Encoding.UTF8.GetBytes(xml.Replace("</ram:ApplicableHeaderTradeSettlement>", tradeTax + "</ram:ApplicableHeaderTradeSettlement>"));
    }

    private static byte[] AddHeaderTradeTaxWithoutTypeCode(byte[] ciiXml) {
        string xml = Encoding.UTF8.GetString(ciiXml);
        string tradeTax = CreateApplicableTradeTax(
            true,
            false,
            true,
            true,
            true,
            true,
            "S",
            "VAT",
            "23",
            "0.00",
            "0.00",
            null,
            null,
            "EUR");
        return Encoding.UTF8.GetBytes(xml.Replace("</ram:ApplicableHeaderTradeSettlement>", tradeTax + "</ram:ApplicableHeaderTradeSettlement>"));
    }

    private static byte[] AddHeaderTradeTaxWithoutCategoryCode(byte[] ciiXml) {
        string xml = Encoding.UTF8.GetString(ciiXml);
        string tradeTax = CreateApplicableTradeTax(
            true,
            true,
            false,
            true,
            true,
            true,
            "S",
            "VAT",
            "23",
            "0.00",
            "0.00",
            null,
            null,
            "EUR");
        return Encoding.UTF8.GetBytes(xml.Replace("</ram:ApplicableHeaderTradeSettlement>", tradeTax + "</ram:ApplicableHeaderTradeSettlement>"));
    }

    private static byte[] AddHeaderLineTotalAmount(byte[] ciiXml) {
        string xml = Encoding.UTF8.GetString(ciiXml);
        return Encoding.UTF8.GetBytes(xml.Replace(
            "<ram:SpecifiedTradeSettlementHeaderMonetarySummation>",
            "<ram:SpecifiedTradeSettlementHeaderMonetarySummation><ram:LineTotalAmount currencyID=\"EUR\">100.00</ram:LineTotalAmount>"));
    }

    private static byte[] AddHeaderAllowanceCharge(byte[] ciiXml, bool charge, string categoryCode, string actualAmount, bool includeRate = false, string rateValue = "0", bool includeReason = true) {
        string xml = Encoding.UTF8.GetString(ciiXml);
        string rate = includeRate
            ? "<ram:RateApplicablePercent>" + rateValue + "</ram:RateApplicablePercent>"
            : string.Empty;
        string reason = includeReason
            ? "<ram:Reason>" + (charge ? "Service charge" : "Document allowance") + "</ram:Reason>"
            : string.Empty;
        string allowanceCharge =
            "<ram:SpecifiedTradeAllowanceCharge>" +
            "<ram:ChargeIndicator><udt:Indicator>" + (charge ? "true" : "false") + "</udt:Indicator></ram:ChargeIndicator>" +
            "<ram:ActualAmount currencyID=\"EUR\">" + actualAmount + "</ram:ActualAmount>" +
            reason +
            "<ram:CategoryTradeTax>" +
            "<ram:TypeCode>VAT</ram:TypeCode>" +
            "<ram:CategoryCode>" + categoryCode + "</ram:CategoryCode>" +
            rate +
            "</ram:CategoryTradeTax>" +
            "</ram:SpecifiedTradeAllowanceCharge>";
        return Encoding.UTF8.GetBytes(xml.Replace("<ram:SpecifiedTradeSettlementHeaderMonetarySummation>", allowanceCharge + "<ram:SpecifiedTradeSettlementHeaderMonetarySummation>"));
    }


}
