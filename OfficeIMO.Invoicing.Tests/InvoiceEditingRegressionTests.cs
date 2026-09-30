using System.Xml.Linq;

namespace OfficeIMO.Invoicing.Tests;

public class InvoiceEditingRegressionTests {
    [Theory]
    [InlineData("PaymentReference", "PaymentReference")]
    [InlineData("CreditorReferenceID", "CreditorIdentifier")]
    [InlineData("DirectDebitMandateID", "DirectDebitMandateReference")]
    public void StandaloneCiiPaymentDataDoesNotInventAPaymentInstruction(string element, string field) {
        Invoice invoice = InvoiceFixture.Create();
        invoice.Payments.Clear();
        XNamespace ram = "urn:un:unece:uncefact:data:standard:ReusableAggregateBusinessInformationEntity:100";
        XDocument document = XDocument.Parse(Encoding.UTF8.GetString(InvoiceSerializer.Write(invoice, InvoiceTestContracts.En16931())));
        XElement settlement = document.Descendants(ram + "ApplicableHeaderTradeSettlement").Single();
        XElement parent = element == "DirectDebitMandateID" ? settlement.Element(ram + "SpecifiedTradePaymentTerms")! : settlement;
        parent.Add(new XElement(ram + element, "standalone-value"));
        byte[] source = Encoding.UTF8.GetBytes(document.ToString());

        InvoiceReadResult read = InvoiceParser.Read(source);

        Assert.True(read.HasCompleteMapping);
        Assert.Empty(read.Invoice.Payments);
        Assert.Equal("standalone-value", typeof(Invoice).GetProperty(field)!.GetValue(read.Invoice));
        Assert.True(InvoiceModelValidator.Validate(read.Invoice).IsValid);
        InvoiceConversionResult conversion = InvoiceConverter.Convert(source, InvoiceTestContracts.En16931());
        Assert.True(conversion.Succeeded, string.Join("; ", conversion.Diagnostics.Select(d => d.Message)));
        Assert.Equal("standalone-value", XDocument.Parse(Encoding.UTF8.GetString(conversion.Xml!)).Descendants(ram + element).Single().Value);
        if (element != "CreditorReferenceID") {
            InvoiceConversionResult ubl = InvoiceConverter.Convert(source, InvoiceTestContracts.En16931(InvoiceSyntax.Ubl));
            Assert.False(ubl.Succeeded);
            Assert.Contains(ubl.Diagnostics, diagnostic => diagnostic.Location == field);
        }
    }

    [Fact]
    public void StandaloneUblCreditorIdentifierRoundTripsWithoutPaymentMeans() {
        Invoice invoice = InvoiceFixture.Create();
        invoice.Payments.Clear();
        invoice.CreditorIdentifier = "creditor-1";
        InvoiceXmlOptions options = InvoiceTestContracts.En16931(InvoiceSyntax.Ubl);
        byte[] xml = InvoiceSerializer.Write(invoice, options);
        InvoiceReadResult read = InvoiceParser.Read(xml);
        Assert.Empty(read.Invoice.Payments);
        Assert.Equal("creditor-1", read.Invoice.CreditorIdentifier);
        Assert.Equal(xml, read.Write(options));
        Assert.True(InvoiceConverter.Convert(xml, InvoiceTestContracts.En16931()).Succeeded);
    }

    [Fact]
    public void IndependentHeaderConflictReportsItsExactPaymentOccurrence() {
        Invoice invoice = InvoiceFixture.Create();
        invoice.PaymentReference = "different";
        Assert.Contains(InvoiceSerializer.InspectTarget(invoice, InvoiceTestContracts.En16931()),
            diagnostic => diagnostic.Location == "Payments[0].Reference");
        Assert.Throws<InvalidDataException>(() => InvoiceSerializer.Write(invoice, InvoiceTestContracts.En16931()));
    }

    [Theory]
    [InlineData(InvoiceSyntax.Cii)]
    [InlineData(InvoiceSyntax.Ubl)]
    public void MissingSellerReturnsDiagnosticsAndBlocksOutput(InvoiceSyntax syntax) {
        Invoice invoice = InvoiceFixture.Create();
        invoice.Seller = null!;
        Assert.Contains(InvoiceModelValidator.Validate(invoice).Diagnostics, d => d.Code == "INV-REQUIRED" && d.Location == "Seller");
        Assert.Contains(InvoiceSerializer.InspectTarget(invoice, InvoiceTestContracts.En16931(syntax)), d => d.Location == "Seller");
        Assert.Throws<InvalidDataException>(() => InvoiceSerializer.Write(invoice, InvoiceTestContracts.En16931(syntax)));
    }

    [Fact]
    public void AggregatePrepaymentAndRoundingEditsRebuildPayableWithoutFabricatingLines() {
        Invoice invoice = InvoiceFixture.Create();
        InvoiceXmlOptions options = InvoiceTestContracts.FacturX(InvoiceProfile.BasicWithoutLines, InvoiceProjectionPolicy.AllowProfileDefinedDataLoss);
        Invoice aggregate = InvoiceParser.Read(InvoiceSerializer.Write(invoice, options)).Invoice;
        aggregate.PrepaidAmount = 20m;
        aggregate.RoundingAmount = 0.01m;
        InvoiceCalculation calculation = InvoiceCalculator.UpdateDeclaredAmounts(aggregate);
        Assert.Equal(99.01m, calculation.PayableAmount);
        Assert.Empty(aggregate.Lines);
        Assert.Equal(100m, aggregate.DeclaredTotals!.TaxExclusiveTotal);
        Assert.Equal(19m, aggregate.DeclaredTotals.TaxTotal);
        Assert.Contains(InvoiceSerializer.InspectTarget(aggregate, options), d => d.Location == "RoundingAmount");
        aggregate.RoundingAmount = 0m;
        InvoiceCalculator.UpdateDeclaredAmounts(aggregate);
        Assert.Equal(99m, InvoiceParser.Read(InvoiceSerializer.Write(aggregate, options)).Invoice.DeclaredTotals!.PayableAmount);
    }

    [Fact]
    public void VatEditRequiresARefreshedAccountingCurrencyAmount() {
        Invoice invoice = InvoiceFixture.Create();
        InvoiceCalculator.UpdateDeclaredAmounts(invoice);
        invoice.TaxCurrency = "USD";
        invoice.TaxAmountInAccountingCurrency = 22m;
        InvoiceCalculator.UpdateDeclaredAmounts(invoice);
        Assert.Equal(22m, invoice.TaxAmountInAccountingCurrency);
        invoice.Lines[0].Quantity = 2m;
        InvoiceCalculator.UpdateDeclaredAmounts(invoice);
        Assert.Null(invoice.TaxAmountInAccountingCurrency);
        Assert.Contains(InvoiceModelValidator.Validate(invoice).Diagnostics, d => d.Code == "INV-TAX-CURRENCY");
        Assert.Throws<InvalidDataException>(() => InvoiceSerializer.Write(invoice, InvoiceTestContracts.En16931()));
        invoice.TaxAmountInAccountingCurrency = 44m;
        Assert.True(InvoiceModelValidator.Validate(invoice).IsValid);
    }

    [Fact]
    public void ChangingEitherCurrencyInvalidatesItsAccountingVatAmount() {
        Invoice invoice = InvoiceFixture.Create();
        invoice.TaxCurrency = "USD";
        invoice.TaxAmountInAccountingCurrency = 22m;
        invoice.TaxCurrency = "USD";
        Assert.Equal(22m, invoice.TaxAmountInAccountingCurrency);
        invoice.TaxCurrency = "PLN";
        Assert.Null(invoice.TaxAmountInAccountingCurrency);
        invoice.TaxAmountInAccountingCurrency = 80.75m;
        invoice.Currency = "GBP";
        Assert.Null(invoice.TaxAmountInAccountingCurrency);
    }

    [Fact]
    public void ExplicitExchangeRateRecalculatesAccountingVatAndInvalidRateDoesNotMutateAmounts() {
        Invoice invoice = InvoiceFixture.Create();
        invoice.TaxCurrency = "USD";
        InvoiceCalculator.UpdateDeclaredAmounts(invoice, 1.1578947368421052631578947368m);
        Assert.Equal(22m, invoice.TaxAmountInAccountingCurrency);
        InvoiceDeclaredTotals before = invoice.DeclaredTotals!;
        invoice.Lines[0].Quantity = 2m;
        Assert.Throws<ArgumentException>(() => InvoiceCalculator.UpdateDeclaredAmounts(invoice, 0m));
        Assert.Same(before, invoice.DeclaredTotals);
        Assert.Equal(22m, invoice.TaxAmountInAccountingCurrency);
        InvoiceCalculator.UpdateDeclaredAmounts(invoice, 1.1578947368421052631578947368m);
        Assert.Equal(44m, invoice.TaxAmountInAccountingCurrency);
        Assert.True(InvoiceModelValidator.Validate(invoice).IsValid);
    }

    [Fact]
    public void EditingReportExposesAccountingRefreshAndFailedRecalculationIsAtomic() {
        Invoice invoice = InvoiceFixture.Create();
        InvoiceCalculator.UpdateDeclaredAmounts(invoice);
        invoice.TaxCurrency = "USD";
        invoice.TaxAmountInAccountingCurrency = 22m;
        invoice.Lines[0].Quantity = 2m;
        InvoiceEditResult edited = InvoiceEditor.Recalculate(invoice);
        Assert.True(edited.Succeeded);
        Assert.Contains(edited.Diagnostics, d => d.Code == "INV-ACCOUNTING-VAT-REFRESH" && d.Location == "TaxAmountInAccountingCurrency");
        InvoiceDeclaredTotals before = invoice.DeclaredTotals!;
        invoice.Lines[0].PriceBaseQuantity = 0m;
        InvoiceEditResult failed = InvoiceEditor.Recalculate(invoice);
        Assert.False(failed.Succeeded);
        Assert.Contains(failed.Diagnostics, d => d.Code == "INV-EDIT-CALCULATION");
        Assert.Same(before, invoice.DeclaredTotals);
    }

    [Fact]
    public void AggregateNullTaxCategoryReturnsDiagnosticsWithoutThrowing() {
        Invoice invoice = InvoiceFixture.Create();
        InvoiceXmlOptions options = InvoiceTestContracts.FacturX(InvoiceProfile.BasicWithoutLines, InvoiceProjectionPolicy.AllowProfileDefinedDataLoss);
        Invoice aggregate = InvoiceParser.Read(InvoiceSerializer.Write(invoice, options)).Invoice;
        aggregate.DeclaredTaxes[0].Category = null!;
        Assert.Contains(InvoiceSerializer.InspectTarget(aggregate, options), d => d.Location == "DeclaredTaxes[0].Category");
        Assert.False(InvoiceEditor.Recalculate(aggregate).Succeeded);
    }
}
