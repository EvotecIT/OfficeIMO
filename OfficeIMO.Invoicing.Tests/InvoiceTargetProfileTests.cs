namespace OfficeIMO.Invoicing.Tests;

public class InvoiceTargetProfileTests {
    [Fact]
    public void CiiBuyerIdentifiersCannotBeSilentlyReducedForUbl() {
        Invoice invoice = InvoiceFixture.Create();
        invoice.Seller.Identifiers.Add(new InvoiceIdentifier("seller-1"));
        invoice.Seller.Identifiers.Add(new InvoiceIdentifier("seller-2"));
        invoice.Buyer.Identifiers.Add(new InvoiceIdentifier("buyer-1"));
        invoice.Buyer.Identifiers.Add(new InvoiceIdentifier("buyer-2"));
        byte[] cii = InvoiceSerializer.Write(invoice);
        InvoiceReadResult read = InvoiceParser.Read(cii);
        Assert.True(read.HasCompleteMapping);
        Assert.Equal(2, read.Invoice.Buyer.Identifiers.Count);
        Assert.Equal(cii, read.Write());
        var ubl = new InvoiceXmlOptions(InvoiceSyntax.Ubl);
        Assert.Contains(InvoiceSerializer.InspectTarget(invoice, ubl), d => d.Location == "Buyer.Identifiers");
        Assert.Throws<InvalidDataException>(() => InvoiceSerializer.Write(invoice, ubl));
        InvoiceConversionResult conversion = InvoiceConverter.Convert(cii, ubl);
        Assert.False(conversion.Succeeded);
        Assert.Null(conversion.Xml);
        invoice.Buyer.Identifiers.RemoveAt(1);
        read = InvoiceParser.Read(InvoiceSerializer.Write(invoice, ubl));
        Assert.Equal(2, read.Invoice.Seller.Identifiers.Count);
        Assert.Single(read.Invoice.Buyer.Identifiers);
    }

    [Theory]
    [InlineData(InvoiceSyntax.Cii, InvoiceProfile.En16931, true)]
    [InlineData(InvoiceSyntax.Ubl, InvoiceProfile.En16931, true)]
    [InlineData(InvoiceSyntax.Cii, InvoiceProfile.XRechnung, true)]
    [InlineData(InvoiceSyntax.Ubl, InvoiceProfile.XRechnung, true)]
    [InlineData(InvoiceSyntax.Ubl, InvoiceProfile.PeppolBis, true)]
    [InlineData(InvoiceSyntax.Ubl, InvoiceProfile.PeppolBis, false)]
    public void DirectDebitFieldsAreCheckedInTheirTargetProfile(InvoiceSyntax syntax, InvoiceProfile profile, bool german) {
        Invoice invoice = DebitInvoice();
        if (!german) { invoice.Seller.Address!.CountryCode = "FR"; invoice.Buyer.Address!.CountryCode = "FR"; }
        var options = new InvoiceXmlOptions(syntax, profile);
        Assert.Empty(InvoiceSerializer.InspectTarget(invoice, options));
        for (int field = 0; field < 3; field++) {
            SetDebitField(invoice.Payment!, field, null);
            bool required = profile != InvoiceProfile.En16931 && (field == 0 || profile == InvoiceProfile.XRechnung || german);
            var diagnostics = InvoiceSerializer.InspectTarget(invoice, options);
            Assert.Equal(required, diagnostics.Any(d => d.Severity == InvoiceDiagnosticSeverity.Error));
            if (required) Assert.Throws<InvalidDataException>(() => InvoiceSerializer.Write(invoice, options));
            else Assert.True(InvoiceParser.Read(InvoiceSerializer.Write(invoice, options)).HasCompleteMapping);
            SetDebitField(invoice.Payment!, field, field == 0 ? "mandate-1" : field == 1 ? "DE98ZZZ09999999999" : "DE89370400440532013000");
        }
    }

    [Theory]
    [InlineData(InvoiceSyntax.Cii, InvoiceProfile.XRechnung)]
    [InlineData(InvoiceSyntax.Ubl, InvoiceProfile.XRechnung)]
    [InlineData(InvoiceSyntax.Ubl, InvoiceProfile.PeppolBis)]
    public void OtherDirectDebitCodeAlsoRequiresMandate(InvoiceSyntax syntax, InvoiceProfile profile) {
        Invoice invoice = DebitInvoice();
        invoice.Payment!.MeansCode = "49";
        invoice.Payment.MandateReference = null;
        Assert.Contains(InvoiceSerializer.InspectTarget(invoice, new InvoiceXmlOptions(syntax, profile)), d => d.Location == "Payment.MandateReference");
        Assert.Empty(InvoiceSerializer.InspectTarget(invoice, new InvoiceXmlOptions(syntax, InvoiceProfile.En16931)));
    }

    [Theory]
    [InlineData(InvoiceSyntax.Cii, InvoiceProfile.XRechnung)]
    [InlineData(InvoiceSyntax.Ubl, InvoiceProfile.XRechnung)]
    [InlineData(InvoiceSyntax.Ubl, InvoiceProfile.PeppolBis)]
    public void GermanSepaDirectDebitCannotIncludeOtherPaymentGroups(InvoiceSyntax syntax, InvoiceProfile profile) {
        Invoice invoice = DebitInvoice();
        invoice.Payment!.Accounts.Add(new InvoiceBankAccount { Identifier = "DE89370400440532013000" });
        invoice.Payment.CardNumber = "1234";
        var diagnostics = InvoiceSerializer.InspectTarget(invoice, new InvoiceXmlOptions(syntax, profile));
        Assert.Contains(diagnostics, d => d.Location == "Payment.Accounts");
        Assert.Contains(diagnostics, d => d.Location == "Payment.CardNumber");
        Assert.Empty(InvoiceSerializer.InspectTarget(invoice, new InvoiceXmlOptions(syntax, InvoiceProfile.En16931)));
    }

    private static Invoice DebitInvoice() {
        Invoice invoice = InvoiceFixture.Create();
        invoice.Payment!.Accounts.Clear();
        invoice.Payment.MeansCode = "59";
        invoice.Payment.MandateReference = "mandate-1";
        invoice.Payment.CreditorIdentifier = "DE98ZZZ09999999999";
        invoice.Payment.DebitedAccount = "DE89370400440532013000";
        return invoice;
    }

    private static void SetDebitField(InvoicePayment payment, int field, string? value) {
        if (field == 0) payment.MandateReference = value;
        else if (field == 1) payment.CreditorIdentifier = value;
        else payment.DebitedAccount = value;
    }
}
