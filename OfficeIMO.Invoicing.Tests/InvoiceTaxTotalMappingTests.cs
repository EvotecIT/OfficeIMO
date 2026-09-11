using System.Text;
using System.Xml.Linq;

namespace OfficeIMO.Invoicing.Tests;

public class InvoiceTaxTotalMappingTests {
    [Theory]
    [InlineData(false, false)]
    [InlineData(true, false)]
    [InlineData(false, true)]
    [InlineData(true, true)]
    public void UblSelectsTheFirstCoherentTaxGroupWhenAnEarlierTotalHasNoBreakdown(bool creditNote, bool differentEarlierTotal) {
        Invoice invoice = InvoiceFixture.Create();
        invoice.TypeCode = creditNote ? "381" : "380";
        if (creditNote) invoice.DueDate = null;
        var options = new InvoiceXmlOptions(InvoiceSyntax.Ubl);
        byte[] original = InvoiceSerializer.Write(invoice, options);
        InvoiceReadResult expected = InvoiceParser.Read(original);
        XDocument document = XDocument.Parse(Encoding.UTF8.GetString(original));
        XElement coherent = document.Root!.Elements().Single(e => e.Name.LocalName == "TaxTotal");
        XElement earlier = new XElement(coherent);
        earlier.Elements().Where(e => e.Name.LocalName == "TaxSubtotal").Remove();
        if (differentEarlierTotal) earlier.Elements().Single(e => e.Name.LocalName == "TaxAmount").Value = "999";
        coherent.AddBeforeSelf(earlier);
        InvoiceReadResult read = InvoiceParser.Read(Encoding.UTF8.GetBytes(document.ToString()));
        Assert.False(read.HasCompleteMapping);
        Assert.Contains(read.UnmappedData, d => d.Message.Contains("duplicated"));
        Assert.Equal(expected.Invoice.DeclaredTotals!.TaxTotal, read.Invoice.DeclaredTotals!.TaxTotal);
        Assert.Equal(expected.Invoice.DeclaredTaxes.Count, read.Invoice.DeclaredTaxes.Count);
        Assert.Throws<InvalidDataException>(() => read.Write());
        foreach (InvoiceSyntax target in new[] { InvoiceSyntax.Cii, InvoiceSyntax.Ubl }) {
            InvoiceReadResult rewritten = InvoiceParser.Read(read.Write(new InvoiceXmlOptions(target), allowUnmappedDataLoss: true));
            Assert.True(rewritten.HasCompleteMapping);
            Assert.Equal(expected.Invoice.DeclaredTotals.TaxTotal, rewritten.Invoice.DeclaredTotals!.TaxTotal);
            Assert.Equal(expected.Invoice.DeclaredTaxes.Count, rewritten.Invoice.DeclaredTaxes.Count);
        }
    }

    [Theory]
    [InlineData(InvoiceSyntax.Cii, false, false)]
    [InlineData(InvoiceSyntax.Cii, true, false)]
    [InlineData(InvoiceSyntax.Cii, false, true)]
    [InlineData(InvoiceSyntax.Cii, true, true)]
    [InlineData(InvoiceSyntax.Ubl, false, false)]
    [InlineData(InvoiceSyntax.Ubl, true, false)]
    [InlineData(InvoiceSyntax.Ubl, false, true)]
    [InlineData(InvoiceSyntax.Ubl, true, true)]
    public void DuplicateVatTotalsRetainTheFirstCurrencyGroupForExplicitLossyRewrite(InvoiceSyntax syntax, bool creditNote, bool accountingCurrency) {
        VerifyDuplicateTotal(syntax, creditNote, accountingCurrency, changeDuplicate: true);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void IdenticalUblTaxGroupsDoNotDuplicateTheRetainedBreakdown(bool creditNote) {
        VerifyDuplicateTotal(InvoiceSyntax.Ubl, creditNote, accountingCurrency: false, changeDuplicate: false);
    }

    private static void VerifyDuplicateTotal(InvoiceSyntax syntax, bool creditNote, bool accountingCurrency, bool changeDuplicate) {
        Invoice invoice = InvoiceFixture.Create();
        invoice.TypeCode = creditNote ? "381" : "380";
        if (creditNote) invoice.DueDate = null;
        invoice.TaxCurrency = "USD";
        invoice.TaxAmountInAccountingCurrency = 25m;
        var options = new InvoiceXmlOptions(syntax);
        byte[] original = InvoiceSerializer.Write(invoice, options);
        InvoiceReadResult expected = InvoiceParser.Read(original);
        XDocument document = XDocument.Parse(Encoding.UTF8.GetString(original));
        string currency = accountingCurrency ? invoice.TaxCurrency : invoice.Currency;
        XElement primary = syntax == InvoiceSyntax.Ubl
            ? document.Root!.Elements().Single(e => e.Name.LocalName == "TaxTotal" && e.Elements().Any(amount => amount.Name.LocalName == "TaxAmount" && (string?)amount.Attribute("currencyID") == currency))
            : document.Descendants().Single(e => e.Name.LocalName == "TaxTotalAmount" && (string?)e.Attribute("currencyID") == currency);
        XElement duplicate = new XElement(primary);
        if (changeDuplicate) {
            XElement amount = syntax == InvoiceSyntax.Ubl ? duplicate.Elements().Single(e => e.Name.LocalName == "TaxAmount") : duplicate;
            amount.Value = "999";
        }
        primary.AddAfterSelf(duplicate);
        byte[] source = Encoding.UTF8.GetBytes(document.ToString());
        InvoiceReadResult read = InvoiceParser.Read(source);

        Assert.False(read.HasCompleteMapping);
        Assert.Contains(read.UnmappedData, d => d.Code == "INV-UNMAPPED" && d.Message.Contains("duplicated"));
        Assert.Equal(source, read.GetOriginalBytes());
        Assert.Equal(expected.Invoice.DeclaredTotals!.TaxTotal, read.Invoice.DeclaredTotals!.TaxTotal);
        Assert.Equal(expected.Invoice.TaxAmountInAccountingCurrency, read.Invoice.TaxAmountInAccountingCurrency);
        Assert.Equal(expected.Invoice.DeclaredTaxes.Count, read.Invoice.DeclaredTaxes.Count);
        Assert.True(InvoiceModelValidator.Validate(read.Invoice).IsValid);
        Assert.Throws<InvalidDataException>(() => read.Write());
        foreach (InvoiceSyntax target in new[] { InvoiceSyntax.Cii, InvoiceSyntax.Ubl }) {
            var targetOptions = new InvoiceXmlOptions(target);
            InvoiceConversionResult conversion = InvoiceConverter.Convert(source, targetOptions);
            Assert.False(conversion.Succeeded);
            Assert.Null(conversion.Xml);
            InvoiceReadResult rewritten = InvoiceParser.Read(read.Write(targetOptions, allowUnmappedDataLoss: true));
            Assert.True(rewritten.HasCompleteMapping);
            Assert.Equal(expected.Invoice.DeclaredTotals.TaxTotal, rewritten.Invoice.DeclaredTotals!.TaxTotal);
            Assert.Equal(expected.Invoice.TaxAmountInAccountingCurrency, rewritten.Invoice.TaxAmountInAccountingCurrency);
            Assert.Equal(expected.Invoice.DeclaredTaxes.Count, rewritten.Invoice.DeclaredTaxes.Count);
            Assert.Equal(expected.Invoice.TypeCode, rewritten.Invoice.TypeCode);
        }
    }
}
