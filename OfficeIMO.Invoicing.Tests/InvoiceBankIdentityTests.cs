namespace OfficeIMO.Invoicing.Tests;

public class InvoiceBankIdentityTests {
    public static IEnumerable<object[]> RegisteredIbans() =>
        File.ReadLines(Path.Combine(AppContext.BaseDirectory, "Fixtures", "Swift", "iban-registry-102.csv"))
            .Skip(1).Where(line => line.Length > 0).Select(line => new object[] { line.Split(',')[2] });

    [Theory]
    [MemberData(nameof(RegisteredIbans))]
    public void RegisteredCountryExamplesRetainIbanIdentityAcrossSyntaxes(string identifier) {
        Invoice invoice = InvoiceFixture.Create();
        invoice.Payment!.Accounts[0].Identifier = identifier;
        invoice.Payment.Accounts[0].IsIban = true;
        foreach (InvoiceSyntax syntax in new[] { InvoiceSyntax.Ubl, InvoiceSyntax.Cii }) {
            InvoiceReadResult result = InvoiceParser.Read(InvoiceSerializer.Write(invoice, new InvoiceXmlOptions(syntax)));
            Assert.True(result.HasCompleteMapping);
            InvoiceBankAccount account = Assert.Single(result.Invoice.Payment!.Accounts);
            Assert.True(account.IsIban);
            Assert.Equal(identifier, account.Identifier);
        }
    }

    [Theory]
    [InlineData("DE", "12345678901")]
    [InlineData("DE", "1234567890123456789")]
    [InlineData("DE", "A23456789012345678")]
    [InlineData("GB", "123460161331926819")]
    [InlineData("US", "123456789012345678")]
    [InlineData("ZZ", "123456789012345678")]
    public void ChecksumCannotOverrideCountryFormatOrAccountKind(string country, string bban) {
        string identifier = WithChecksum(country, bban);
        Invoice invoice = InvoiceFixture.Create();
        invoice.Payment!.Accounts[0].Identifier = identifier;
        invoice.Payment.Accounts[0].IsIban = false;
        byte[] source = InvoiceSerializer.Write(invoice, new InvoiceXmlOptions(InvoiceSyntax.Ubl));
        InvoiceReadResult read = InvoiceParser.Read(source);
        Assert.True(read.HasCompleteMapping);
        Assert.False(Assert.Single(read.Invoice.Payment!.Accounts).IsIban);
        InvoiceConversionResult cii = InvoiceConverter.Convert(source, new InvoiceXmlOptions());
        Assert.True(cii.Succeeded);
        Assert.False(Assert.Single(InvoiceParser.Read(cii.Xml!).Invoice.Payment!.Accounts).IsIban);
        Assert.Contains("<ram:ProprietaryID>" + identifier + "</ram:ProprietaryID>", Encoding.UTF8.GetString(cii.Xml!));

        invoice.Payment.Accounts[0].IsIban = true;
        Assert.Contains("Payment.Accounts[0]", Assert.Throws<InvalidDataException>(() => InvoiceSerializer.Write(invoice)).Message);
        invoice.Payment.Accounts[0].IsIban = false;
        invoice.Payment.MeansCode = "49";
        invoice.Payment.DebitedAccount = identifier;
        source = InvoiceSerializer.Write(invoice, new InvoiceXmlOptions(InvoiceSyntax.Ubl));
        cii = InvoiceConverter.Convert(source, new InvoiceXmlOptions());
        Assert.False(cii.Succeeded);
        Assert.Null(cii.Xml);
        Assert.Contains(cii.Diagnostics, diagnostic => diagnostic.Location == "Payment.DebitedAccount");
    }

    [Theory]
    [InlineData("de89 3704 0044 0532 0130 00")]
    [InlineData("DE89\t370400440532013000")]
    public void AsciiCaseAndSpacingRemainSupported(string identifier) {
        Invoice invoice = InvoiceFixture.Create();
        invoice.Payment!.Accounts[0].Identifier = identifier;
        Assert.NotEmpty(InvoiceSerializer.Write(invoice));
        Assert.True(Assert.Single(InvoiceParser.Read(InvoiceSerializer.Write(invoice,
            new InvoiceXmlOptions(InvoiceSyntax.Ubl))).Invoice.Payment!.Accounts).IsIban);
    }

    [Theory]
    [InlineData("GB29NWＢK60161331926819")]
    [InlineData("GB29NWBK6016133192681９")]
    [InlineData("GB00NWBK60161331926819")]
    [InlineData("DE99000000000000000030")]
    [InlineData("DE01000000000000000048")]
    [InlineData("DE00000000000000000066")]
    [InlineData("GB71ſABC60161331926819")]
    public void NonAsciiCharactersAndInvalidCheckDigitsAreNotIbans(string identifier) {
        Invoice invoice = InvoiceFixture.Create();
        invoice.Payment!.Accounts[0].Identifier = identifier;
        Assert.Contains("Payment.Accounts[0]", Assert.Throws<InvalidDataException>(() => InvoiceSerializer.Write(invoice)).Message);
    }

    private static string WithChecksum(string country, string bban) {
        // Independent full decimal representation, rather than the production streaming remainder.
        string rearranged = bban + country + "00";
        string digits = string.Concat(rearranged.Select(character => character >= 'A'
            ? ((int)character - 'A' + 10).ToString(System.Globalization.CultureInfo.InvariantCulture)
            : character.ToString()));
        int checksum = 98 - (int)(System.Numerics.BigInteger.Parse(digits, System.Globalization.CultureInfo.InvariantCulture) % 97);
        return country + checksum.ToString("00", System.Globalization.CultureInfo.InvariantCulture) + bban;
    }
}
