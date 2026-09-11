using System.Globalization;

namespace OfficeIMO.Invoicing.Tests;

public class InvoiceArithmeticBoundaryTests {
    [Theory]
    [InlineData("0.00000000000000000001", "0.0000000001", "0.0000000000000000000000000001", "0.01")]
    [InlineData("-0.00000000000000000001", "0.0000000001", "0.0000000000000000000000000001", "-0.01")]
    [InlineData("1", "1", "200.00000000000000000000000001", "0.00")]
    [InlineData("1", "1", "199.99999999999999999999999999", "0.01")]
    [InlineData("-1", "1", "199.99999999999999999999999999", "-0.01")]
    [InlineData("-1", "1", "200", "0.00")]
    [InlineData("79228162514264337593543950335", "2", "2", "79228162514264337593543950335")]
    [InlineData("-79228162514264337593543950335", "2", "2", "-79228162514264337593543950335")]
    public void LineAmountsRoundTheExactFormulaWithoutIntermediatePrecisionLoss(string quantity, string price, string basis, string expected) {
        Invoice invoice = InvoiceFixture.Create();
        invoice.Lines[0].Quantity = Parse(quantity);
        invoice.Lines[0].UnitPrice = Parse(price);
        invoice.Lines[0].PriceBaseQuantity = Parse(basis);
        invoice.Lines[0].Tax = new InvoiceTaxCategory { Code = "Z", Rate = 0m };
        InvoiceModelValidationResult validation = InvoiceModelValidator.Validate(invoice);
        Assert.True(validation.IsValid, string.Join("; ", validation.Diagnostics.Select(d => d.Message)));
        Assert.Equal(Parse(expected), validation.Calculation!.LineNetTotal);
        Assert.Equal(Parse(expected), validation.Calculation.PayableAmount);
        foreach (InvoiceSyntax syntax in new[] { InvoiceSyntax.Cii, InvoiceSyntax.Ubl }) {
            Invoice parsed = InvoiceParser.Read(InvoiceSerializer.Write(invoice, new InvoiceXmlOptions(syntax))).Invoice;
            Assert.Equal(Parse(expected), parsed.Lines[0].DeclaredNetAmount);
            Assert.Equal(Parse(expected), parsed.DeclaredTotals!.PayableAmount);
        }
    }

    [Fact]
    public void TaxAndAdjustmentPercentagesAvoidIntermediateOverflow() {
        const decimal basis = 10000000000000000000000000000m;
        Invoice invoice = InvoiceFixture.Create();
        invoice.Lines[0].UnitPrice = basis;
        invoice.Lines[0].Tax.Rate = 50m;
        InvoiceModelValidationResult tax = InvoiceModelValidator.Validate(invoice);
        Assert.True(tax.IsValid, string.Join("; ", tax.Diagnostics.Select(d => d.Message)));
        Assert.Equal(5000000000000000000000000000m, tax.Calculation!.TaxTotal);
        invoice.AllowancesAndCharges.Add(new InvoiceAllowanceCharge {
            IsCharge = true, Amount = basis, BaseAmount = basis, Percentage = 100m, Reason = "Additional service",
            Tax = new InvoiceTaxCategory { Code = "S", Rate = 50m }
        });
        InvoiceModelValidationResult adjusted = InvoiceModelValidator.Validate(invoice);
        Assert.True(adjusted.IsValid, string.Join("; ", adjusted.Diagnostics.Select(d => d.Message)));
        Assert.Equal(30000000000000000000000000000m, adjusted.Calculation!.PayableAmount);
    }

    [Fact]
    public void UnrepresentableMonetaryResultReturnsCalculationDiagnostics() {
        Invoice invoice = InvoiceFixture.Create();
        invoice.Lines[0].Quantity = decimal.MaxValue;
        invoice.Lines[0].UnitPrice = 2m;
        Assert.Contains(InvoiceModelValidator.Validate(invoice).Diagnostics, d => d.Code == "INV-CALCULATION");
        Assert.Equal(decimal.MinValue, InvoiceCalculator.RoundAmount(decimal.MinValue));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void TotalsDoNotSilentlyDiscardCentsBeyondDecimalPrecision(bool prepaid) {
        Invoice invoice = InvoiceFixture.Create();
        invoice.Lines[0].UnitPrice = 10000000000000000000000000000m;
        invoice.Lines[0].Tax = new InvoiceTaxCategory { Code = "Z", Rate = 0m };
        if (prepaid) invoice.PrepaidAmount = 0.01m;
        else invoice.Lines.Add(new InvoiceLine { Id = "2", Name = "Small fee", Quantity = 1m, UnitPrice = 0.01m, Tax = invoice.Lines[0].Tax });
        Assert.Contains(InvoiceModelValidator.Validate(invoice).Diagnostics, d => d.Code == "INV-CALCULATION");
    }

    [Fact]
    public void GrossPriceEquationCannotSilentlyDiscardAnUnrepresentableDiscount() {
        Invoice invoice = InvoiceFixture.Create();
        invoice.Lines[0].UnitPrice = 10000000000000000000000000000m;
        invoice.Lines[0].GrossPrice = invoice.Lines[0].UnitPrice;
        invoice.Lines[0].PriceDiscount = 0.01m;
        Assert.Contains(InvoiceModelValidator.Validate(invoice).Diagnostics, d => d.Code == "INV-OVERFLOW");
    }

    private static decimal Parse(string value) => decimal.Parse(value, CultureInfo.InvariantCulture);
}
