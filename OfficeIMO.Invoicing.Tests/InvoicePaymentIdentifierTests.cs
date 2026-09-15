using System.Xml.Linq;

namespace OfficeIMO.Invoicing.Tests;

public class InvoicePaymentIdentifierTests {
    [Fact]
    public void ConflictingUblPaymentReferencesRemainIndependentAndBlockCiiConversion() {
        Invoice invoice = InvoiceFixture.Create();
        invoice.Payments.Add(new InvoicePayment {
            MeansCode = invoice.Payments[0].MeansCode,
            MeansText = "Second route",
            Reference = "second-reference",
            Account = new InvoiceBankAccount { Identifier = "DE79000000001234567890" }
        });
        byte[] xml = InvoiceSerializer.Write(invoice, InvoiceTestContracts.En16931(InvoiceSyntax.Ubl));
        InvoiceReadResult read = InvoiceParser.Read(xml);
        Assert.True(read.HasCompleteMapping);
        Assert.Equal(2, read.Invoice.Payments.Count);
        Assert.Equal(invoice.Payments[0].Reference, read.Invoice.Payments[0].Reference);
        Assert.Equal("second-reference", read.Invoice.Payments[1].Reference);
        Assert.Equal("Second route", read.Invoice.Payments[1].MeansText);
        Assert.Equal(xml, read.Write(InvoiceTestContracts.En16931(InvoiceSyntax.Ubl)));
        InvoiceConversionResult converted = InvoiceConverter.Convert(xml, InvoiceTestContracts.En16931(InvoiceSyntax.Cii));
        Assert.False(converted.Succeeded);
        Assert.Null(converted.Xml);
        Assert.Contains(converted.Diagnostics, d => d.Location == "Payments[1].Reference" && d.Message.IndexOf("second-reference", StringComparison.Ordinal) >= 0);
    }

    [Fact]
    public void UblCardNetworkRoundTripsAndCiiReportsTheExactUnsupportedValue() {
        Invoice invoice = InvoiceFixture.Create();
        InvoicePayment payment = invoice.Payments[0];
        payment.MeansCode = "48";
        payment.Account = null;
        payment.CardNumber = "1234";
        payment.CardHolder = "Card Holder";
        payment.CardNetworkId = "VISA";
        byte[] xml = InvoiceSerializer.Write(invoice, InvoiceTestContracts.En16931(InvoiceSyntax.Ubl));
        InvoiceReadResult read = InvoiceParser.Read(xml);
        Assert.True(read.HasCompleteMapping);
        Assert.Equal("VISA", read.Invoice.Payments[0].CardNetworkId);
        Assert.Equal(xml, read.Write(InvoiceTestContracts.En16931(InvoiceSyntax.Ubl)));
        InvoiceConversionResult converted = InvoiceConverter.Convert(xml, InvoiceTestContracts.En16931());
        Assert.False(converted.Succeeded);
        Assert.Contains(converted.Diagnostics, d => d.Location == "Payments[0].CardNetworkId" && d.Message.IndexOf("VISA", StringComparison.Ordinal) >= 0);
    }

    [Fact]
    public void UblCardAccountRequiresAnExplicitNetworkIdentifier() {
        Invoice invoice = InvoiceFixture.Create();
        InvoicePayment payment = invoice.Payments[0];
        payment.MeansCode = "48";
        payment.Account = null;
        payment.CardNumber = "1234";
        payment.CardNetworkId = null;
        IReadOnlyList<InvoiceDiagnostic> diagnostics = InvoiceSerializer.InspectTarget(invoice, InvoiceTestContracts.En16931(InvoiceSyntax.Ubl));
        Assert.Contains(diagnostics, d => d.Location == "Payments[0].CardNetworkId");
        Assert.Throws<InvalidDataException>(() => InvoiceSerializer.Write(invoice, InvoiceTestContracts.En16931(InvoiceSyntax.Ubl)));
    }

    [Fact]
    public void ConflictingPaymentMeansCodesAreReportedPerOccurrence() {
        Invoice invoice = InvoiceFixture.Create();
        invoice.Payments.Add(new InvoicePayment { MeansCode = "30", Account = new InvoiceBankAccount { Identifier = "DE79000000001234567890" } });
        foreach (InvoiceSyntax syntax in new[] { InvoiceSyntax.Cii, InvoiceSyntax.Ubl }) {
            IReadOnlyList<InvoiceDiagnostic> diagnostics = InvoiceSerializer.InspectTarget(invoice, InvoiceTestContracts.En16931(syntax));
            Assert.Contains(diagnostics, d => d.Location == "Payments[1].MeansCode" && d.Message.IndexOf("30", StringComparison.Ordinal) >= 0);
            Assert.Throws<InvalidDataException>(() => InvoiceSerializer.Write(invoice, InvoiceTestContracts.En16931(syntax)));
        }
    }

    [Theory]
    [InlineData(InvoiceSyntax.Cii)]
    [InlineData(InvoiceSyntax.Ubl)]
    public void MultiplePaymentOccurrencesAndTheirAccountsRoundTrip(InvoiceSyntax syntax) {
        Invoice invoice = InvoiceFixture.Create();
        invoice.Payments.Add(new InvoicePayment {
            MeansCode = invoice.Payments[0].MeansCode,
            Account = new InvoiceBankAccount { Identifier = "DE79000000001234567890", Name = "Second account" }
        });
        byte[] xml = InvoiceSerializer.Write(invoice, InvoiceTestContracts.En16931(syntax));
        InvoiceReadResult read = InvoiceParser.Read(xml);
        Assert.True(read.HasCompleteMapping, string.Join("; ", read.UnmappedData.Select(d => d.Message)));
        Assert.Equal(2, read.Invoice.Payments.Count);
        Assert.Equal(invoice.Payments[0].Account!.Identifier, read.Invoice.Payments[0].Account!.Identifier);
        Assert.Equal("DE79000000001234567890", read.Invoice.Payments[1].Account!.Identifier);
        Assert.Equal("Second account", read.Invoice.Payments[1].Account!.Name);
    }

    [Theory]
    [InlineData(null)]
    [InlineData("")]
    [InlineData(" ")]
    public void StandardItemIdentifiersRequireANonEmptyScheme(string? scheme) {
        Invoice invoice = InvoiceFixture.Create();
        invoice.Lines[0].StandardItemIdentifier = new InvoiceIdentifier("1234567890128", scheme);
        Assert.Contains(InvoiceModelValidator.Validate(invoice).Diagnostics, d => d.Location == "Lines[0].StandardItemIdentifier.SchemeId");
        foreach (InvoiceSyntax syntax in new[] { InvoiceSyntax.Cii, InvoiceSyntax.Ubl })
            Assert.Throws<InvalidDataException>(() => InvoiceSerializer.Write(invoice, InvoiceTestContracts.En16931(syntax)));
    }

    [Fact]
    public void OptionalIdentifiersStillRequireTheirSuppliedValueAndScheme() {
        Invoice invoice = InvoiceFixture.Rich();
        invoice.ObjectIdentifier!.Value = "";
        invoice.Lines[0].ObjectIdentifier!.SchemeId = " ";
        invoice.Delivery!.LocationIdentifier!.Value = "";
        invoice.Payee!.LegalRegistration!.Value = "";
        InvoiceModelValidationResult result = InvoiceModelValidator.Validate(invoice);
        foreach (string path in new[] { "ObjectIdentifier.Value", "Lines[0].ObjectIdentifier.SchemeId", "Delivery.LocationIdentifier.Value", "Payee.LegalRegistration.Value" })
            Assert.Contains(result.Diagnostics, d => d.Location == path && d.Severity == InvoiceDiagnosticSeverity.Error);
    }

    [Theory]
    [InlineData(null, "IB", "Value")]
    [InlineData("", "IB", "Value")]
    [InlineData(" ", "IB", "Value")]
    [InlineData("0721-880X", null, "ListId")]
    [InlineData("0721-880X", "", "ListId")]
    [InlineData("0721-880X", " ", "ListId")]
    public void ClassificationsRequireBothValueAndListIdentifier(string? value, string? list, string missing) {
        Invoice invoice = InvoiceFixture.Create();
        invoice.Lines[0].Classifications.Add(new InvoiceItemClassification { Value = value!, ListId = list! });
        Assert.Contains(InvoiceModelValidator.Validate(invoice).Diagnostics, d => d.Location == "Lines[0].Classifications." + missing);
        foreach (InvoiceSyntax syntax in new[] { InvoiceSyntax.Cii, InvoiceSyntax.Ubl })
            Assert.Throws<InvalidDataException>(() => InvoiceSerializer.Write(invoice, InvoiceTestContracts.En16931(syntax)));
    }
}
