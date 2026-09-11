namespace OfficeIMO.Invoicing.Tests;

public class InvoiceTaxGroupingScaleTests {
    [Fact]
    public void ManyDistinctRatesPreserveAmountsAndDeclaredBreakdownMatches() {
        Invoice invoice = InvoiceFixture.Create();
        for (int index = 1; index <= 15000; index++)
            invoice.AllowancesAndCharges.Add(new InvoiceAllowanceCharge {
                IsCharge = true, Amount = 0m, Reason = "Rate group",
                Tax = new InvoiceTaxCategory { Code = "S", Rate = 1m + index / 10000m }
            });
        InvoiceCalculation calculation = InvoiceCalculator.UpdateDeclaredAmounts(invoice);
        Assert.Equal(15001, calculation.Taxes.Count);
        Assert.Equal(119m, calculation.PayableAmount);
        Assert.Equal(15001, invoice.DeclaredTaxes.Count);
        InvoiceModelValidationResult validation = InvoiceModelValidator.Validate(invoice);
        Assert.True(validation.IsValid, string.Join(Environment.NewLine, validation.Diagnostics.Select(d => d.Message)));
        Assert.Equal(calculation.PayableAmount, validation.Calculation!.PayableAmount);
    }
}
