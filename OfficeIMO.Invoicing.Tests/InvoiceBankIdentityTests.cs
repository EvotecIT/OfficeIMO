namespace OfficeIMO.Invoicing.Tests;

public class InvoiceBankIdentityTests {
    public static IEnumerable<object[]> RegisteredIbans() =>
        File.ReadLines(Path.Combine(AppContext.BaseDirectory, "Fixtures", "Swift", "iban-registry-102.csv"))
            .Skip(1).Where(line => line.Length > 0).Select(line => new object[] { line.Split(',')[2] });

    [Theory]
    [MemberData(nameof(RegisteredIbans))]
    public void RegisteredCountryExamplesRetainIbanIdentityAcrossSyntaxes(string identifier) {
        Invoice invoice = InvoiceFixture.Create();
        invoice.Payments[0].Account!.Identifier = identifier;
        invoice.Payments[0].Account!.IsIban = true;
        foreach (InvoiceSyntax syntax in new[] { InvoiceSyntax.Ubl, InvoiceSyntax.Cii }) {
            InvoiceReadResult result = InvoiceParser.Read(InvoiceSerializer.Write(invoice, InvoiceTestContracts.En16931(syntax)));
            Assert.True(result.HasCompleteMapping);
            InvoiceBankAccount account = Assert.IsType<InvoiceBankAccount>(result.Invoice.Payments[0].Account);
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
        invoice.Payments[0].Account!.Identifier = identifier;
        invoice.Payments[0].Account!.IsIban = false;
        byte[] source = InvoiceSerializer.Write(invoice, InvoiceTestContracts.En16931(InvoiceSyntax.Ubl));
        InvoiceReadResult read = InvoiceParser.Read(source);
        Assert.True(read.HasCompleteMapping);
        Assert.False(read.Invoice.Payments[0].Account!.IsIban);
        InvoiceConversionResult cii = InvoiceConverter.Convert(source, InvoiceTestContracts.En16931());
        Assert.True(cii.Succeeded);
        Assert.False(InvoiceParser.Read(cii.Xml!).Invoice.Payments[0].Account!.IsIban);
        Assert.Contains("<ram:ProprietaryID>" + identifier + "</ram:ProprietaryID>", Encoding.UTF8.GetString(cii.Xml!));

        invoice.Payments[0].Account!.IsIban = true;
        Assert.Contains("Payments[0].Account", Assert.Throws<InvalidDataException>(() => InvoiceSerializer.Write(invoice, InvoiceTestContracts.En16931())).Message);
        invoice.Payments[0].Account!.IsIban = false;
        invoice.Payments[0].MeansCode = "49";
        invoice.Payments[0].DebitedAccount = identifier;
        source = InvoiceSerializer.Write(invoice, InvoiceTestContracts.En16931(InvoiceSyntax.Ubl));
        cii = InvoiceConverter.Convert(source, InvoiceTestContracts.En16931());
        Assert.False(cii.Succeeded);
        Assert.Null(cii.Xml);
        Assert.Contains(cii.Diagnostics, diagnostic => diagnostic.Location == "Payments[0].DebitedAccount");
    }

    [Theory]
    [InlineData("de89 3704 0044 0532 0130 00")]
    [InlineData("DE89\t370400440532013000")]
    public void AsciiCaseAndSpacingRemainSupported(string identifier) {
        Invoice invoice = InvoiceFixture.Create();
        invoice.Payments[0].Account!.Identifier = identifier;
        Assert.NotEmpty(InvoiceSerializer.Write(invoice, InvoiceTestContracts.En16931()));
        Assert.True(InvoiceParser.Read(InvoiceSerializer.Write(invoice,
            InvoiceTestContracts.En16931(InvoiceSyntax.Ubl))).Invoice.Payments[0].Account!.IsIban);
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
        invoice.Payments[0].Account!.Identifier = identifier;
        Assert.Contains("Payments[0].Account", Assert.Throws<InvalidDataException>(() => InvoiceSerializer.Write(invoice, InvoiceTestContracts.En16931())).Message);
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
