using System.Xml.Linq;

namespace OfficeIMO.Invoicing.Tests;

public class Fa3InvoiceReaderTests {
    public static IEnumerable<object[]> OfficialExamples => Enumerable.Range(1, 26).Select(number => new object[] { number });
    private static byte[] Example(int number) => File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "Fixtures", "FA3", "example-" + number.ToString("D2") + ".xml"));
    private static readonly XNamespace Ns = Fa3InvoiceReader.NamespaceUri;

    [Theory]
    [MemberData(nameof(OfficialExamples))]
    public void IndependentNationalExamplesRetainExactIdentityAmountsAndRows(int number) {
        byte[] xml = Example(number);
        XElement root = XDocument.Parse(Encoding.UTF8.GetString(xml)).Root!;
        XElement fa = root.Element(Ns + "Fa")!;
        Fa3InvoiceReadResult result = Fa3InvoiceReader.Read(xml);
        Assert.Equal(fa.Element(Ns + "P_2")!.Value, result.Invoice.Number);
        Assert.Equal(fa.Element(Ns + "KodWaluty")!.Value, result.Invoice.Currency);
        Assert.Equal(decimal.Parse(fa.Element(Ns + "P_15")!.Value, System.Globalization.CultureInfo.InvariantCulture), result.DeclaredTotal);
        Assert.Equal(fa.Elements(Ns + "FaWiersz").Count(), result.Lines.Count);
        Assert.Equal(xml, result.GetOriginalBytes());
        Assert.Contains(result.CommonMappingDiagnostics, item => item.Code == "FA3-NATIONAL-TOTALS");
        Assert.Throws<InvalidDataException>(() => result.WriteCommonInvoice(new InvoiceXmlOptions(InvoiceSpecificationRelease.En16931_1_3_16, InvoiceSyntax.Cii, InvoiceProfile.En16931)));
        // These national fixtures carry fields beyond the supported projection. The report must expose them.
        Assert.NotEmpty(result.UnmappedData);
    }

    [Fact]
    public void CorrectionPreviousStateIsNotAddedAsAnotherBilledLine() {
        Fa3InvoiceReadResult result = Fa3InvoiceReader.Read(Example(2));
        Assert.Equal(Fa3InvoiceKind.Correction, result.Kind);
        Assert.Equal(-200m, result.DeclaredTotal);
        Assert.Equal(-162.60m, Assert.Single(result.TaxSummaries).TaxableAmount);
        Assert.Equal(-37.40m, result.TaxSummaries[0].TaxAmount);
        Assert.True(result.Lines[0].IsPreviousState);
        Assert.False(result.Lines[1].IsPreviousState);
        Assert.Equal(1626.01m, result.Lines[0].NetAmount);
        Assert.Equal(1463.41m, result.Lines[1].NetAmount);
        Assert.Empty(result.Invoice.Lines);
        Assert.Null(result.Invoice.DeclaredTotals);
    }

    [Fact]
    public void SimplifiedAndForeignCurrencyAmountsRemainAbsentOrExplicit() {
        Fa3InvoiceReadResult simplified = Fa3InvoiceReader.Read(Example(16));
        Fa3InvoiceLineData line = Assert.Single(simplified.Lines);
        Assert.Null(line.Quantity); Assert.Null(line.UnitPrice); Assert.Null(line.UnitLabel);
        Assert.Empty(simplified.Invoice.Lines);
        Assert.Equal(string.Empty, simplified.Invoice.Buyer.Name);
        Fa3InvoiceReadResult foreign = Fa3InvoiceReader.Read(Example(20));
        Assert.Equal("EUR", foreign.Invoice.Currency);
        Assert.NotNull(Assert.Single(foreign.TaxSummaries).TaxAmountInPln);
        Assert.Null(foreign.Invoice.TaxAmountInAccountingCurrency);
        Assert.Null(foreign.Invoice.TaxCurrency);
        Assert.Null(foreign.Invoice.Seller.Address.City);
        Assert.Equal("00-001 Warszawa", foreign.Invoice.Seller.Address.Line2);
    }

    [Fact]
    public void SourceBytesAreDefensiveAndRemainIndependentOfCommonModelEdits() {
        byte[] source = Example(1); byte[] expected = (byte[])source.Clone();
        Fa3InvoiceReadResult result = Fa3InvoiceReader.Read(source);
        source[0] = 0; result.GetOriginalBytes()[1] = 0;
        result.Invoice.Number = "edited";
        Assert.Equal(expected, result.GetOriginalBytes());
        InvoiceTaxRegistration nip = Assert.Single(result.Invoice.Seller.TaxRegistrations);
        Assert.Equal("9999999999", nip.Identifier); Assert.Equal("NIP", nip.SchemeId); Assert.Equal(InvoiceTaxRegistrationKind.Fiscal, nip.Kind);
        Assert.Contains(result.UnmappedData, item => item.Location.Contains("NrKlienta[1]"));
    }

    [Theory]
    [InlineData("duplicate")]
    [InlineData("nested")]
    [InlineData("declaration")]
    [InlineData("timezone")]
    public void AmbiguousOrInconsistentSourceCannotBeMisidentified(string mutation) {
        XDocument document = XDocument.Parse(Encoding.UTF8.GetString(Example(1)));
        XElement fa = document.Root!.Element(Ns + "Fa")!;
        switch (mutation) {
            case "duplicate": fa.Add(new XElement(fa.Element(Ns + "P_2")!)); break;
            case "nested": fa.Element(Ns + "P_2")!.Add(new XElement(Ns + "nested", "value")); break;
            case "declaration": document.Descendants(Ns + "KodFormularza").Single().SetAttributeValue("wersjaSchemy", "2-0E"); break;
            case "timezone": document.Descendants(Ns + "DataWytworzeniaFa").Single().Value = "2026-02-01T00:00:00"; break;
        }
        Assert.Throws<InvalidDataException>(() => Fa3InvoiceReader.Read(Encoding.UTF8.GetBytes(document.ToString())));
    }
}
