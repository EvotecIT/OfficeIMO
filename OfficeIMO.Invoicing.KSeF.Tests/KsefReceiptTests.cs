using System.Security.Cryptography;
using System.Text;
using System.Xml.Linq;
using OfficeIMO.Invoicing.KSeF;
using OfficeIMO.Invoicing.Validation;

namespace OfficeIMO.Invoicing.KSeF.Tests;

public class KsefReceiptTests {
    private const string Session = "36950822-93-9D5A28BFDA-47C899773E-5C";
    private const string Number = "5265877635-20250916-0200A0D6723E-C2";
    private const string Hash = "GZMGNVzs3krF6URKgvaw77OOeG3nJ+WGziT5xguliQ8=";
    [Theory]
    [InlineData("upo-invoice-nip.xml", "DA5688B6DB93181E945ACAEDFAF55F8BBBA9197B5F35751327501B1A87BD4422")]
    [InlineData("upo-session-nip.xml", "AB42D4090C3E453935E48FFB450F6FD10911110DA99EB9500A75E296A4C1FDBB")]
    public void IndependentTestReceiptBindingAndSchemaDisagreementRemainSeparate(string fixture, string digest) {
        byte[] xml = File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "Fixtures", fixture));
        Assert.Equal(digest, Convert.ToHexString(SHA256.HashData(xml)));
        KsefReceipt receipt = KsefReceipt.Inspect(xml, new KsefContext(KsefContextKind.Nip, "5265877635"), Session, Number, Hash);
        Assert.True(receipt.IsBound); Assert.False(receipt.RetrievedThroughAuthenticatedApi);
        Assert.Equal(InvoiceValidationStatus.Invalid, receipt.SchemaStatus);
        Assert.Contains(receipt.Diagnostics, diagnostic => diagnostic.Code == "XSD" && diagnostic.Message.Contains("NazwaPodmiotuPrzyjmujacego", StringComparison.Ordinal));
        Assert.Equal(xml, receipt.GetBytes());
    }
    [Theory]
    [InlineData(KsefContextKind.Nip, "5265877635", "Nip")]
    [InlineData(KsefContextKind.InternalId, "5265877635-00001", "IdWewnetrzny")]
    [InlineData(KsefContextKind.NipVatUe, "5265877635-DE123456789", "IdZlozonyVatUE")]
    [InlineData(KsefContextKind.PeppolId, "PPL000001", "IdDostawcyUslugPeppol")]
    public void NativeReceiptContextNamesAreBoundWithoutRelabelingIdentifiers(KsefContextKind kind, string value, string tag) {
        XDocument document = XDocument.Load(Path.Combine(AppContext.BaseDirectory, "Fixtures", "upo-invoice-nip.xml"));
        XElement identifier = document.Descendants().Single(element => element.Name.LocalName == "IdKontekstu");
        identifier.ReplaceNodes(new XElement(identifier.Name.Namespace + tag, value));
        document.Descendants().Single(element => element.Name.LocalName == "NazwaPodmiotuPrzyjmujacego").Value = "Ministerstwo Finansów";
        byte[] xml = Encoding.UTF8.GetBytes(document.ToString());
        KsefReceipt receipt = KsefReceipt.Inspect(xml, new KsefContext(kind, value), Session, Number, Hash);
        Assert.True(receipt.IsBound); Assert.Equal(InvoiceValidationStatus.Passed, receipt.SchemaStatus);
        identifier.Add(new XElement(identifier.Name.Namespace + tag, value));
        Assert.False(KsefReceipt.Inspect(Encoding.UTF8.GetBytes(document.ToString()), new KsefContext(kind, value), Session, Number, Hash).IsBound);
    }
    [Fact]
    public void WrongDocumentBindingDtdAndBoundsNeverProduceQualifiedReceipts() {
        byte[] xml = File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "Fixtures", "upo-invoice-nip.xml")); var context = new KsefContext(KsefContextKind.Nip, "5265877635");
        Assert.False(KsefReceipt.Inspect(xml, context, Session, Number, Convert.ToBase64String(new byte[32])).IsBound);
        Assert.False(KsefReceipt.Inspect(xml, new KsefContext(KsefContextKind.Nip, "9999999999"), Session, Number, Hash).IsBound);
        Assert.False(KsefReceipt.Inspect(Encoding.UTF8.GetBytes("<!DOCTYPE x [<!ENTITY value 'untrusted'>]><x>&value;</x>"), context, Session, Number, Hash).IsBound);
        Assert.Throws<InvalidDataException>(() => KsefReceipt.Inspect(new byte[2 * 1024 * 1024 + 1], context, Session, Number, Hash));
        Assert.Throws<ArgumentException>(() => new KsefContext(KsefContextKind.Nip, "0000000000"));
    }
}
