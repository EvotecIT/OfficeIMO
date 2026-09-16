namespace OfficeIMO.Invoicing.Tests;

public class InvoiceConversionBoundaryTests {
    [Theory]
    [InlineData("LOCAL-123")]
    [InlineData("DE12345678901234567890")]
    [InlineData("DE5112345678901")]
    public void InvalidSourceDebtorIbanCannotBeRecastAsAGenericUblAccount(string identifier) {
        Invoice invoice = InvoiceFixture.Create();
        invoice.Payments[0].MeansCode = "49";
        invoice.Payments[0].DebitedAccount = "DE89370400440532013000";
        string source = System.Text.Encoding.UTF8.GetString(InvoiceSerializer.Write(invoice, InvoiceTestContracts.En16931())).Replace(invoice.Payments[0].DebitedAccount!, identifier);
        InvoiceConversionResult result = InvoiceConverter.Convert(System.Text.Encoding.UTF8.GetBytes(source), InvoiceTestContracts.En16931(InvoiceSyntax.Ubl));
        Assert.False(result.Succeeded);
        Assert.Null(result.Xml);
        Assert.Contains(result.Diagnostics, d => d.Severity == InvoiceDiagnosticSeverity.Error);
    }

    [Theory]
    [InlineData("")]
    [InlineData("<bad>")]
    [InlineData("<not-an-invoice />")]
    public void InvalidSourceReturnsDiagnosticsWithoutOutput(string source) {
        InvoiceConversionResult result = InvoiceConverter.Convert(System.Text.Encoding.UTF8.GetBytes(source), InvoiceTestContracts.En16931());
        Assert.False(result.Succeeded);
        Assert.Null(result.Xml);
        Assert.Contains(result.Diagnostics, d => d.Code == "INV-CONVERSION-INPUT" && d.Severity == InvoiceDiagnosticSeverity.Error);
    }

    [Theory]
    [InlineData("LOCAL-123")]
    [InlineData("DE12345678901234567890")]
    public void GenericCreditorAccountRetainsItsIdentityAcrossBothSyntaxes(string identifier) {
        Invoice invoice = InvoiceFixture.Create();
        invoice.Payments[0].Account!.Identifier = identifier;
        invoice.Payments[0].Account!.IsIban = false;
        byte[] source = InvoiceSerializer.Write(invoice, InvoiceTestContracts.En16931(InvoiceSyntax.Cii));
        InvoiceConversionResult ubl = InvoiceConverter.Convert(source, InvoiceTestContracts.En16931(InvoiceSyntax.Ubl));
        Assert.True(ubl.Succeeded);
        InvoiceConversionResult cii = InvoiceConverter.Convert(ubl.Xml!, InvoiceTestContracts.En16931(InvoiceSyntax.Cii));
        Assert.True(cii.Succeeded);
        InvoiceBankAccount account = InvoiceParser.Read(cii.Xml!).Invoice.Payments[0].Account!;
        Assert.False(account.IsIban);
        Assert.Equal(identifier, account.Identifier);
    }

    [Fact]
    public void ExplicitProprietaryAccountCannotLoseItsClassificationInUbl() {
        Invoice invoice = InvoiceFixture.Create();
        invoice.Payments[0].Account!.IsIban = false;
        InvoiceConversionResult result = InvoiceConverter.Convert(InvoiceSerializer.Write(invoice, InvoiceTestContracts.En16931()), InvoiceTestContracts.En16931(InvoiceSyntax.Ubl));
        Assert.False(result.Succeeded);
        Assert.Null(result.Xml);
        Assert.Contains(result.Diagnostics, d => d.Location == "Payments[0].Account");
    }

    [Fact]
    public void ValidDebitedIbanRoundTripsAcrossBothSyntaxes() {
        Invoice invoice = InvoiceFixture.Create();
        invoice.Payments[0].MeansCode = "49";
        invoice.Payments[0].DebitedAccount = "DE89370400440532013000";
        InvoiceConversionResult result = InvoiceConverter.Convert(InvoiceSerializer.Write(invoice, InvoiceTestContracts.En16931(InvoiceSyntax.Ubl)), InvoiceTestContracts.En16931());
        Assert.True(result.Succeeded);
        Assert.Equal(invoice.Payments[0].DebitedAccount, InvoiceParser.Read(result.Xml!).Invoice.Payments[0].DebitedAccount);
    }

    [Fact]
    public void TargetXmlExpansionReturnsDiagnosticsWithoutPartialOutput() {
        Invoice invoice = InvoiceFixture.Create();
        invoice.SupportingDocuments.Add(new InvoiceSupportingDocument {
            Reference = "large", FileName = "large.bin", MimeType = "application/octet-stream", Data = new byte[8 * 1024 * 1024]
        });
        for (int index = 0; index < 4; index++) invoice.Notes.Add(new InvoiceNote(new string('A', 1000000)));
        for (int index = 0; index < 30000; index++) invoice.Notes.Add(new InvoiceNote("N"));
        byte[] source = InvoiceSerializer.Write(invoice, InvoiceTestContracts.En16931(InvoiceSyntax.Ubl));
        Assert.InRange(source.Length, 15 * 1024 * 1024, InvoiceProfileDeclaration.MaximumXmlBytes);
        Assert.True(InvoiceParser.Read(source).HasCompleteMapping);
        Assert.Throws<InvalidDataException>(() => InvoiceSerializer.Write(invoice, InvoiceTestContracts.En16931()));
        InvoiceConversionResult result = InvoiceConverter.Convert(source, InvoiceTestContracts.En16931(InvoiceSyntax.Cii));
        Assert.False(result.Succeeded);
        Assert.Null(result.Xml);
        Assert.Contains(result.Diagnostics, d => d.Code == "INV-CONVERSION-OUTPUT" && d.Severity == InvoiceDiagnosticSeverity.Error);
    }

    [Theory]
    [InlineData("LOCAL-123")]
    [InlineData("DE12345678901234567890")]
    public void GenericUblDebitedAccountCannotBeRelabeledAsAnIban(string account) {
        Invoice invoice = InvoiceFixture.Create();
        invoice.Payments[0].MeansCode = "49";
        invoice.Payments[0].DebitedAccount = account;
        byte[] source = InvoiceSerializer.Write(invoice, InvoiceTestContracts.En16931(InvoiceSyntax.Ubl));
        InvoiceReadResult read = InvoiceParser.Read(source);
        Assert.True(read.HasCompleteMapping);
        Assert.Equal(account, read.Invoice.Payments[0].DebitedAccount);
        Assert.Equal(account, InvoiceParser.Read(read.Write(InvoiceTestContracts.En16931(InvoiceSyntax.Ubl))).Invoice.Payments[0].DebitedAccount);
        InvoiceConversionResult result = InvoiceConverter.Convert(source, InvoiceTestContracts.En16931(InvoiceSyntax.Cii));
        Assert.False(result.Succeeded);
        Assert.Null(result.Xml);
        Assert.Contains(result.Diagnostics, d => d.Location == "Payments[0].DebitedAccount");
    }
}
