using System.Globalization;

namespace OfficeIMO.Invoicing.Tests;

public class InvoiceCalculationTests {
    [Fact]
    public void MixedRatesAdjustmentsBaseQuantitiesAndPrepaymentsShareOneCalculation() {
        Invoice invoice = Example();
        invoice.Lines[0].Quantity = 3m;
        invoice.Lines[0].PriceBaseQuantity = 2m;
        invoice.Lines[0].UnitPrice = 10m;
        invoice.Lines[0].AllowancesAndCharges.Add(new InvoiceAllowanceCharge { Amount = 1m, Reason = "Line discount" });
        invoice.Lines.Add(new InvoiceLine { Id = "2", Name = "Books", Quantity = 2m, UnitPrice = 5m, Tax = new InvoiceTaxCategory { Code = "S", Rate = 7m } });
        invoice.AllowancesAndCharges.Add(new InvoiceAllowanceCharge { IsCharge = true, Amount = 2m, Reason = "Delivery", Tax = new InvoiceTaxCategory { Code = "S", Rate = 19m } });
        invoice.PrepaidAmount = 5m;
        invoice.RoundingAmount = 0.01m;
        InvoiceModelValidationResult validation = InvoiceModelValidator.Validate(invoice);
        Assert.True(validation.IsValid, string.Join("; ", validation.Diagnostics.Select(d => d.Message)));
        InvoiceCalculation result = validation.Calculation!;
        Assert.Equal(14m, result.Lines[0].NetAmount);
        Assert.Equal(24m, result.LineNetTotal);
        Assert.Equal(26m, result.TaxExclusiveTotal);
        Assert.Equal(3.74m, result.TaxTotal);
        Assert.Equal(24.75m, result.PayableAmount);
        Assert.Equal(new[] { 0.70m, 3.04m }, result.Taxes.Select(t => t.TaxAmount));
    }

    [Theory]
    [InlineData("1.005", "1.01")]
    [InlineData("-1.005", "-1.00")]
    [InlineData("-1.006", "-1.01")]
    [InlineData("-1.004", "-1.00")]
    public void RoundingMatchesXPathForPositiveAndNegativeTies(string value, string expected) =>
        Assert.Equal(decimal.Parse(expected, CultureInfo.InvariantCulture), InvoiceCalculator.RoundAmount(decimal.Parse(value, CultureInfo.InvariantCulture)));

    [Fact]
    public void DeclaredAmountsAreCheckedRatherThanSilentlyReplaced() {
        Invoice invoice = Example();
        invoice.DeclaredTotals = new InvoiceDeclaredTotals { PayableAmount = 0m };
        var validation = InvoiceModelValidator.Validate(invoice);
        Assert.False(validation.IsValid);
        Assert.Contains(validation.Diagnostics, d => d.Code == "INV-DECLARED-AMOUNT" && d.Location == "DeclaredTotals.PayableAmount");
        Assert.Equal(0m, invoice.DeclaredTotals.PayableAmount);
        Assert.Throws<InvalidDataException>(() => InvoiceSerializer.Write(invoice));
    }

    [Fact]
    public void InvalidArithmeticAndDateCannotBeSerialized() {
        Invoice invoice = Example();
        invoice.Lines[0].PriceBaseQuantity = 0m;
        invoice.IssueDate = invoice.IssueDate.AddHours(1);
        var validation = InvoiceModelValidator.Validate(invoice);
        Assert.Contains(validation.Diagnostics, d => d.Code == "INV-BASE-QUANTITY");
        Assert.Contains(validation.Diagnostics, d => d.Code == "INV-DATE");
        Assert.Null(validation.Calculation);
    }

    [Fact]
    public void SourceRoundingIsPreservedUntilExplicitRecalculation() {
        Invoice invoice = Example();
        decimal formula = InvoiceCalculator.Calculate(invoice).Lines[0].NetAmount;
        invoice.Lines[0].DeclaredNetAmount = formula + 0.01m;
        InvoiceCalculation initial = InvoiceCalculator.Calculate(invoice);
        invoice.DeclaredTaxes.Add(new InvoiceDeclaredTax {
            Category = new InvoiceTaxCategory { Code = "S", Rate = 19m },
            TaxableAmount = initial.LineNetTotal, TaxAmount = initial.TaxTotal + 0.01m
        });
        Assert.True(InvoiceModelValidator.Validate(invoice).IsValid);
        foreach (InvoiceSyntax syntax in new[] { InvoiceSyntax.Cii, InvoiceSyntax.Ubl }) {
            var parsed = InvoiceParser.Read(InvoiceSerializer.Write(invoice, new InvoiceXmlOptions(syntax, InvoiceProfile.En16931)));
            Assert.Equal(formula + 0.01m, parsed.Invoice.Lines[0].DeclaredNetAmount);
            Assert.Equal(initial.TaxTotal + 0.01m, parsed.Invoice.DeclaredTaxes[0].TaxAmount);
        }
        InvoiceCalculation recalculated = InvoiceCalculator.UpdateDeclaredAmounts(invoice);
        Assert.Equal(formula, recalculated.Lines[0].NetAmount);
        Assert.Equal(formula, invoice.Lines[0].DeclaredNetAmount);
        Assert.Equal(recalculated.Taxes[0].FormulaTaxAmount, invoice.DeclaredTaxes[0].TaxAmount);
    }

    [Fact]
    public void MaterialDeclaredLineDifferenceStillBlocksWriting() {
        Invoice invoice = Example();
        invoice.Lines[0].DeclaredNetAmount = InvoiceCalculator.Calculate(invoice).Lines[0].NetAmount + 0.03m;
        Assert.False(InvoiceModelValidator.Validate(invoice).IsValid);
        Assert.Throws<InvalidDataException>(() => InvoiceSerializer.Write(invoice));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void RecalculationUsesEditedExemptionInsteadOfStaleDeclaredReason(bool documentAdjustment) {
        Invoice invoice = Example();
        var edited = new InvoiceTaxCategory { Code = "E", Rate = 0m, ExemptionReason = "Updated exemption" };
        if (documentAdjustment) invoice.AllowancesAndCharges.Add(new InvoiceAllowanceCharge {
            IsCharge = true, Amount = 10m, Reason = "Service", Tax = edited
        });
        else invoice.Lines[0].Tax = edited;
        invoice.DeclaredTaxes.Add(new InvoiceDeclaredTax {
            Category = new InvoiceTaxCategory { Code = "E", Rate = 0m, ExemptionReason = "Original exemption", ExemptionReasonCode = "VATEX-EU-132" },
            TaxableAmount = 1m, TaxAmount = 0m
        });
        InvoiceCalculator.UpdateDeclaredAmounts(invoice);
        InvoiceDeclaredTax actual = Assert.Single(invoice.DeclaredTaxes.Where(tax => tax.Category.Code == "E"));
        Assert.Equal("Updated exemption", actual.Category.ExemptionReason);
        Assert.Null(actual.Category.ExemptionReasonCode);
        Assert.True(InvoiceModelValidator.Validate(invoice).IsValid);
        foreach (InvoiceSyntax syntax in new[] { InvoiceSyntax.Cii, InvoiceSyntax.Ubl }) {
            Invoice parsed = InvoiceParser.Read(InvoiceSerializer.Write(invoice, new InvoiceXmlOptions(syntax, InvoiceProfile.En16931))).Invoice;
            Assert.Equal("Updated exemption", Assert.Single(parsed.DeclaredTaxes.Where(tax => tax.Category.Code == "E")).Category.ExemptionReason);
        }
    }

    [Fact]
    public void RecalculationPreservesHeaderOnlyImportedExemption() {
        Invoice invoice = Example();
        invoice.Lines[0].Tax = new InvoiceTaxCategory { Code = "E", Rate = 0m };
        invoice.DeclaredTaxes.Add(new InvoiceDeclaredTax {
            Category = new InvoiceTaxCategory { Code = "E", Rate = 0m, ExemptionReason = "Imported exemption", ExemptionReasonCode = "VATEX-EU-132" }
        });
        InvoiceCalculator.UpdateDeclaredAmounts(invoice);
        InvoiceDeclaredTax actual = Assert.Single(invoice.DeclaredTaxes);
        Assert.Equal("Imported exemption", actual.Category.ExemptionReason);
        Assert.Equal("VATEX-EU-132", actual.Category.ExemptionReasonCode);
    }

    internal static Invoice Example() => InvoiceFixture.Create();
}
