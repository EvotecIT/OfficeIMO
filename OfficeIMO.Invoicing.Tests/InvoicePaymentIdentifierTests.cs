using System.Xml.Linq;

namespace OfficeIMO.Invoicing.Tests;

public class InvoicePaymentIdentifierTests {
    [Theory]
    [InlineData(InvoiceSyntax.Cii, true, 0)]
    [InlineData(InvoiceSyntax.Cii, true, 2)]
    [InlineData(InvoiceSyntax.Ubl, true, 0)]
    [InlineData(InvoiceSyntax.Ubl, true, 2)]
    [InlineData(InvoiceSyntax.Cii, false, 0)]
    [InlineData(InvoiceSyntax.Cii, false, 2)]
    [InlineData(InvoiceSyntax.Ubl, false, 0)]
    [InlineData(InvoiceSyntax.Ubl, false, 2)]
    public void SingletonPaymentDetailsRoundTripWithZeroOrMultipleTransferAccounts(InvoiceSyntax syntax, bool card, int accounts) {
        Invoice invoice = InvoiceFixture.Create();
        InvoicePayment payment = invoice.Payment!;
        payment.Accounts.Clear();
        for (int index = 0; index < accounts; index++) payment.Accounts.Add(new InvoiceBankAccount { Identifier = "DE79000000001234567890", Name = "Account " + index });
        payment.MeansCode = card ? "48" : "59";
        if (card) { payment.CardNumber = "1234"; payment.CardHolder = "Card Holder"; }
        else { payment.MandateReference = "mandate-1"; payment.DebitedAccount = "DE79000000001234567890"; payment.CreditorIdentifier = "DE98ZZZ09999999999"; }
        byte[] xml = InvoiceSerializer.Write(invoice, new InvoiceXmlOptions(syntax));
        XDocument document = XDocument.Parse(System.Text.Encoding.UTF8.GetString(xml));
        string singleton = card ? syntax == InvoiceSyntax.Cii ? "ApplicableTradeSettlementFinancialCard" : "CardAccount"
            : syntax == InvoiceSyntax.Cii ? "PayerPartyDebtorFinancialAccount" : "PaymentMandate";
        Assert.Single(document.Descendants().Where(element => element.Name.LocalName == singleton));
        InvoiceReadResult read = InvoiceParser.Read(xml);
        Assert.True(read.HasCompleteMapping, string.Join("; ", read.UnmappedData.Select(d => d.Message)));
        Assert.Equal(accounts, read.Invoice.Payment!.Accounts.Count);
        Assert.Equal(payment.CardNumber, read.Invoice.Payment.CardNumber);
        Assert.Equal(payment.CardHolder, read.Invoice.Payment.CardHolder);
        Assert.Equal(payment.MandateReference, read.Invoice.Payment.MandateReference);
        Assert.Equal(payment.DebitedAccount, read.Invoice.Payment.DebitedAccount);
        Assert.Equal(payment.CreditorIdentifier, read.Invoice.Payment.CreditorIdentifier);
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
            Assert.Throws<InvalidDataException>(() => InvoiceSerializer.Write(invoice, new InvoiceXmlOptions(syntax)));
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
            Assert.Throws<InvalidDataException>(() => InvoiceSerializer.Write(invoice, new InvoiceXmlOptions(syntax)));
    }
}
