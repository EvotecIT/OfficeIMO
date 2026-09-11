namespace OfficeIMO.Invoicing.Tests;

public class InvoiceConversionBoundaryTests {
    [Theory]
    [InlineData("")]
    [InlineData("<bad>")]
    [InlineData("<not-an-invoice />")]
    public void InvalidSourceReturnsDiagnosticsWithoutOutput(string source) {
        InvoiceConversionResult result = InvoiceConverter.Convert(System.Text.Encoding.UTF8.GetBytes(source), new InvoiceXmlOptions());
        Assert.False(result.Succeeded);
        Assert.Null(result.Xml);
        Assert.Contains(result.Diagnostics, d => d.Code == "INV-CONVERSION-INPUT" && d.Severity == InvoiceDiagnosticSeverity.Error);
    }

    [Theory]
    [InlineData("LOCAL-123")]
    [InlineData("DE12345678901234567890")]
    public void GenericCreditorAccountRetainsItsIdentityAcrossBothSyntaxes(string identifier) {
        Invoice invoice = InvoiceFixture.Create();
        invoice.Payment!.Accounts[0].Identifier = identifier;
        invoice.Payment.Accounts[0].IsIban = false;
        byte[] source = InvoiceSerializer.Write(invoice, new InvoiceXmlOptions(InvoiceSyntax.Cii));
        InvoiceConversionResult ubl = InvoiceConverter.Convert(source, new InvoiceXmlOptions(InvoiceSyntax.Ubl));
        Assert.True(ubl.Succeeded);
        InvoiceConversionResult cii = InvoiceConverter.Convert(ubl.Xml!, new InvoiceXmlOptions(InvoiceSyntax.Cii));
        Assert.True(cii.Succeeded);
        InvoiceBankAccount account = InvoiceParser.Read(cii.Xml!).Invoice.Payment!.Accounts[0];
        Assert.False(account.IsIban);
        Assert.Equal(identifier, account.Identifier);
    }

    [Fact]
    public void ExplicitProprietaryAccountCannotLoseItsClassificationInUbl() {
        Invoice invoice = InvoiceFixture.Create();
        invoice.Payment!.Accounts[0].IsIban = false;
        InvoiceConversionResult result = InvoiceConverter.Convert(InvoiceSerializer.Write(invoice), new InvoiceXmlOptions(InvoiceSyntax.Ubl));
        Assert.False(result.Succeeded);
        Assert.Null(result.Xml);
        Assert.Contains(result.Diagnostics, d => d.Location == "Payment.Accounts[0]");
    }

    [Fact]
    public void ValidDebitedIbanRoundTripsAcrossBothSyntaxes() {
        Invoice invoice = InvoiceFixture.Create();
        invoice.Payment!.MeansCode = "49";
        invoice.Payment.DebitedAccount = "DE89370400440532013000";
        InvoiceConversionResult result = InvoiceConverter.Convert(InvoiceSerializer.Write(invoice, new InvoiceXmlOptions(InvoiceSyntax.Ubl)), new InvoiceXmlOptions());
        Assert.True(result.Succeeded);
        Assert.Equal(invoice.Payment.DebitedAccount, InvoiceParser.Read(result.Xml!).Invoice.Payment!.DebitedAccount);
    }

    [Fact]
    public void TargetXmlExpansionReturnsDiagnosticsWithoutPartialOutput() {
        Invoice invoice = InvoiceFixture.Create();
        invoice.SupportingDocuments.Add(new InvoiceSupportingDocument {
            Reference = "large", FileName = "large.bin", MimeType = "application/octet-stream", Data = new byte[8 * 1024 * 1024]
        });
        for (int index = 0; index < 4; index++) invoice.Notes.Add(new InvoiceNote(new string('A', 1000000)));
        for (int index = 0; index < 30000; index++) invoice.Notes.Add(new InvoiceNote("N"));
        byte[] source = InvoiceSerializer.Write(invoice, new InvoiceXmlOptions(InvoiceSyntax.Ubl));
        Assert.InRange(source.Length, 15 * 1024 * 1024, InvoiceProfileDeclaration.MaximumXmlBytes);
        Assert.True(InvoiceParser.Read(source).HasCompleteMapping);
        Assert.Throws<InvalidDataException>(() => InvoiceSerializer.Write(invoice));
        InvoiceConversionResult result = InvoiceConverter.Convert(source, new InvoiceXmlOptions(InvoiceSyntax.Cii));
        Assert.False(result.Succeeded);
        Assert.Null(result.Xml);
        Assert.Contains(result.Diagnostics, d => d.Code == "INV-CONVERSION-OUTPUT" && d.Severity == InvoiceDiagnosticSeverity.Error);
    }

    [Theory]
    [InlineData("LOCAL-123")]
    [InlineData("DE12345678901234567890")]
    public void GenericUblDebitedAccountCannotBeRelabeledAsAnIban(string account) {
        Invoice invoice = InvoiceFixture.Create();
        invoice.Payment!.MeansCode = "49";
        invoice.Payment.DebitedAccount = account;
        byte[] source = InvoiceSerializer.Write(invoice, new InvoiceXmlOptions(InvoiceSyntax.Ubl));
        InvoiceReadResult read = InvoiceParser.Read(source);
        Assert.True(read.HasCompleteMapping);
        Assert.Equal(account, read.Invoice.Payment!.DebitedAccount);
        Assert.Equal(account, InvoiceParser.Read(read.Write()).Invoice.Payment!.DebitedAccount);
        InvoiceConversionResult result = InvoiceConverter.Convert(source, new InvoiceXmlOptions(InvoiceSyntax.Cii));
        Assert.False(result.Succeeded);
        Assert.Null(result.Xml);
        Assert.Contains(result.Diagnostics, d => d.Location == "Payment.DebitedAccount");
    }
}
