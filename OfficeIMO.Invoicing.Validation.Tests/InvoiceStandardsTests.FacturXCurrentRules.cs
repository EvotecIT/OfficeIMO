using System.Text;
using System.Xml.Linq;
using OfficeIMO.Invoicing.Tests;

namespace OfficeIMO.Invoicing.Validation.Tests;

public partial class InvoiceStandardsTests {
    [InvoiceStandardsTheory]
    [InlineData(InvoiceProfile.BasicWithoutLines)]
    [InlineData(InvoiceProfile.Basic)]
    [InlineData(InvoiceProfile.En16931)]
    public async Task FacturX1092ReportsTheRenamedCreditorAccountRule(InvoiceProfile profile) {
        XDocument document = CurrentFacturXDocument(profile);
        XNamespace ram = "urn:un:unece:uncefact:data:standard:ReusableAggregateBusinessInformationEntity:100";
        XElement settlement = document.Descendants(ram + "ApplicableHeaderTradeSettlement").Single();
        settlement.Elements(ram + "SpecifiedTradeSettlementPaymentMeans").Remove();
        settlement.Elements(ram + "ApplicableTradeTax").First().AddBeforeSelf(
            new XElement(ram + "SpecifiedTradeSettlementPaymentMeans", new XElement(ram + "TypeCode", "58")));

        InvoiceValidationReport report = await ValidateCurrentFacturX(document);

        Assert.Equal(InvoiceValidationStatus.Passed, report.SchemaStatus);
        Assert.Equal(InvoiceValidationStatus.Invalid, report.BusinessRulesStatus);
        Assert.Contains(report.Diagnostics, diagnostic => diagnostic.Code == "CII-SR-470");
    }

    [InvoiceStandardsTheory]
    [InlineData(true, "BR-FXEXT-BR-54-1")]
    [InlineData(false, "BR-FXEXT-BR-54-2")]
    public async Task FacturX1092ExtendedReportsTheReplacementItemAttributeRules(bool omitName, string expectedRule) {
        XDocument document = CurrentFacturXDocument(InvoiceProfile.Extended);
        XNamespace ram = "urn:un:unece:uncefact:data:standard:ReusableAggregateBusinessInformationEntity:100";
        document.Descendants(ram + "SpecifiedTradeProduct").First().Add(
            new XElement(ram + "ApplicableProductCharacteristic",
                omitName ? new XElement(ram + "Value", "Standard") : new XElement(ram + "Description", "Service tier")));

        InvoiceValidationReport report = await ValidateCurrentFacturX(document);

        Assert.Equal(InvoiceValidationStatus.Passed, report.SchemaStatus);
        Assert.Equal(InvoiceValidationStatus.Invalid, report.BusinessRulesStatus);
        Assert.Contains(report.Diagnostics, diagnostic => diagnostic.Code == expectedRule);
    }

    [InvoiceStandardsTheory]
    [InlineData("0.30", true)]
    [InlineData("0.31", false)]
    public async Task FacturX1092ExtendedRoundsGroupSubtotalsAndAllowsGroupTaxWithoutCategory(string subtotal, bool valid) {
        Invoice invoice = InvoiceFixture.Create();
        invoice.BusinessProcessId = null;
        invoice.Lines[0].UnitPrice = 0.1m;
        invoice.Lines.Add(new InvoiceLine { Id = "2", Name = "Second service", Quantity = 1m,
            UnitPrice = 0.2m, Tax = new InvoiceTaxCategory { Code = "S", Rate = 19m } });
        XDocument document = XDocument.Parse(Encoding.UTF8.GetString(
            InvoiceSerializer.Write(invoice, InvoiceTestContracts.FacturX(InvoiceProfile.Extended))));
        XNamespace ram = "urn:un:unece:uncefact:data:standard:ReusableAggregateBusinessInformationEntity:100";
        XElement[] details = document.Descendants(ram + "IncludedSupplyChainTradeLineItem").ToArray();
        foreach (XElement detail in details)
            detail.Element(ram + "AssociatedDocumentLineDocument")!.Add(
                new XElement(ram + "ParentLineID", "group"), new XElement(ram + "LineStatusReasonCode", "DETAIL"));
        details[0].AddBeforeSelf(new XElement(ram + "IncludedSupplyChainTradeLineItem",
            new XElement(ram + "AssociatedDocumentLineDocument", new XElement(ram + "LineID", "group"),
                new XElement(ram + "LineStatusReasonCode", "GROUP")),
            new XElement(ram + "SpecifiedTradeProduct", new XElement(ram + "Name", "Services subtotal")),
            new XElement(ram + "SpecifiedLineTradeSettlement",
                new XElement(ram + "ApplicableTradeTax", new XElement(ram + "DueDateTypeCode", "5")),
                new XElement(ram + "SpecifiedTradeSettlementLineMonetarySummation", new XElement(ram + "LineTotalAmount", subtotal)))));

        InvoiceValidationReport report = await ValidateCurrentFacturX(document);

        Assert.Equal(InvoiceValidationStatus.Passed, report.SchemaStatus);
        if (valid) Assert.True(report.IsValid, Report(report));
        else {
            Assert.Equal(InvoiceValidationStatus.Invalid, report.BusinessRulesStatus);
            Assert.Contains(report.Diagnostics, diagnostic => diagnostic.Code == "BR-FXEXT-08");
        }
    }

    private static XDocument CurrentFacturXDocument(InvoiceProfile profile) {
        Invoice invoice = InvoiceFixture.Create();
        invoice.BusinessProcessId = null;
        return XDocument.Parse(Encoding.UTF8.GetString(InvoiceSerializer.Write(invoice,
            InvoiceTestContracts.FacturX(profile, InvoiceProjectionPolicy.AllowProfileDefinedDataLoss))));
    }

    private static Task<InvoiceValidationReport> ValidateCurrentFacturX(XDocument document) =>
        new InvoiceValidator(Bundle(), Runner()).ValidateAsync(Encoding.UTF8.GetBytes(document.ToString()),
            InvoiceSpecificationRelease.FacturX_1_09_2_Zugferd_2_5_2);
}
