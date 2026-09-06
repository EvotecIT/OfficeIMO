using System.Xml;
using System.Xml.Linq;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Pdf.Tests;

public class PdfCiiInvoiceDocumentTests {
    private const string Rsm = "urn:un:unece:uncefact:data:standard:CrossIndustryInvoice:100";
    private const string Ram = "urn:un:unece:uncefact:data:standard:ReusableAggregateBusinessInformationEntity:100";
    private const string Udt = "urn:un:unece:uncefact:data:standard:UnqualifiedDataType:100";

    [Fact]
    public void StreamRoundTripHonorsCurrentPositionAndLeavesStreamsOpen() {
        byte[] bytes = Encoding.UTF8.GetBytes(Invoice("INV-1"));
        using var input = new MemoryStream(new byte[] { 1, 2, 3 }.Concat(bytes).ToArray());
        input.Position = 3;
        var invoice = PdfCiiInvoiceDocument.Load(input);
        Assert.True(input.CanRead);
        Assert.Equal(input.Length, input.Position);
        using var output = new MemoryStream();
        output.WriteByte(42);
        invoice.Save(output);
        Assert.True(output.CanWrite);
        Assert.Equal(new byte[] { 42 }.Concat(bytes).ToArray(), output.ToArray());
    }

    [Fact]
    public void FileRoundTripPreservesBytesAndCanSaveEditedXml() {
        string path = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xml");
        try {
            byte[] bytes = Encoding.UTF8.GetBytes(Invoice("INV-1"));
            File.WriteAllBytes(path, bytes);
            var invoice = PdfCiiInvoiceDocument.Load(path);
            invoice.Save(path);
            Assert.Equal(bytes, File.ReadAllBytes(path));
            invoice.WithDocumentId("INV-2").Save(path);
            Assert.Equal("INV-2", PdfCiiInvoiceDocument.Load(path).DocumentId);
        } finally {
            File.Delete(path);
        }
    }

    [Fact]
    public void SavingToACallerStreamCannotExposeTheDocumentBuffer() {
        byte[] bytes = Encoding.UTF8.GetBytes(Invoice("INV-1"));
        var invoice = PdfCiiInvoiceDocument.Load(bytes);
        using var stream = new MutatingInvoiceStream();
        invoice.Save(stream);
        Assert.Equal(bytes, invoice.ToBytes());
        stream.Captured![1] ^= 1;
        Assert.Equal(bytes, invoice.ToBytes());
        Assert.Equal("INV-1", invoice.DocumentId);
        Assert.Equal(bytes, new PdfOptions().UseFacturXDocument(invoice).EmbeddedFiles.Single().Data);
    }

    [Fact]
    public void NonSeekableInputIsBoundedWithoutReadingPastTheLimit() {
        using var valid = new NonSeekableInvoiceStream(Encoding.UTF8.GetBytes(Invoice("INV-1")));
        Assert.Equal("INV-1", PdfCiiInvoiceDocument.Load(valid).DocumentId);
        Assert.True(valid.CanRead);
        using var oversized = new NonSeekableInvoiceStream(new byte[PdfCiiInvoiceDocument.MaximumXmlBytes + 100]);
        Assert.Throws<InvalidDataException>(() => PdfCiiInvoiceDocument.Load(oversized));
        Assert.Equal(PdfCiiInvoiceDocument.MaximumXmlBytes + 1, oversized.BytesRead);
        Assert.True(oversized.CanRead);
    }

    [Fact]
    public void UneditedRoundTripPreservesOriginalBytesAndDefensiveCopies() {
        string xml = Invoice("INV-1");
        byte[] bytes = Encoding.Unicode.GetPreamble().Concat(Encoding.Unicode.GetBytes(xml)).ToArray();
        byte[] expected = (byte[])bytes.Clone();
        var document = PdfCiiInvoiceDocument.Load(bytes);
        bytes[20] ^= 1;
        byte[] result = document.ToBytes();
        Assert.Equal(expected, result);
        result[20] ^= 1;
        Assert.Equal(expected, document.ToBytes());
        Assert.Equal("INV-1", document.DocumentId);
        Assert.Equal("380", document.TypeCode);
        Assert.Equal("EUR", document.CurrencyCode);
        Assert.Equal(new DateTime(2026, 9, 5), document.IssueDate);
    }

    [Fact]
    public void EditingRetainsUnknownXmlAndLeavesOriginalAndOtherFieldsAlone() {
        var original = Load(Invoice("INV-1"));
        var edited = original.WithDocumentId("INV<&2").WithIssueDate(new DateTime(2028, 2, 29));
        Assert.Equal("INV-1", original.DocumentId);
        Assert.Equal("INV<&2", edited.DocumentId);
        Assert.Equal(new DateTime(2028, 2, 29), edited.IssueDate);
        XDocument before = XDocument.Parse(Invoice("INV-1"), LoadOptions.PreserveWhitespace);
        XDocument after = XDocument.Parse(Encoding.UTF8.GetString(edited.ToBytes()), LoadOptions.PreserveWhitespace);
        Assert.True(XNode.DeepEquals(before.Root!.Element(XName.Get("Extension", "urn:custom")),
            after.Root!.Element(XName.Get("Extension", "urn:custom"))));
        Assert.Equal("keep", after.Root.Attribute(XName.Get("flag", "urn:custom"))!.Value);
        Assert.Equal("INV-1", after.Descendants(XName.Get("PaymentReference", Ram)).Single().Value);
        Assert.Contains("<!--retain-->", Encoding.UTF8.GetString(edited.ToBytes()));
        Assert.Equal(edited.ToBytes(), PdfCiiInvoiceDocument.Load(edited.ToBytes()).ToBytes());
    }

    [Fact]
    public void NamespacesAndDirectPathsPreventSpoofedFieldSelection() {
        string xml = Invoice("INV-1").Replace("<ram:ID>INV-1</ram:ID>", "<x:ID>spoof</x:ID><ram:ID>INV-1</ram:ID>");
        var document = Load(xml.Replace("rsm:", "root:").Replace("xmlns:rsm", "xmlns:root"));
        Assert.Equal("INV-1", document.DocumentId);
        Assert.Equal("INV-2", document.WithDocumentId("INV-2").DocumentId);
        Assert.Throws<InvalidDataException>(() => Load(xml.Replace(Rsm, Rsm + "0")));
    }

    [Theory]
    [InlineData("<ram:ID>one</ram:ID><ram:ID>two</ram:ID>", typeof(InvalidDataException))]
    [InlineData("<ram:ID><x:value>one</x:value></ram:ID>", typeof(InvalidOperationException))]
    public void AmbiguousOrStructuredFieldsAreNotReadAsInvoiceIdentifiers(string field, Type editException) {
        var document = Load(Invoice("INV-1").Replace("<ram:ID>INV-1</ram:ID>", field));
        Assert.Throws<InvalidDataException>(() => document.DocumentId);
        Assert.Throws(editException, () => document.WithDocumentId("replacement"));
    }

    [Theory]
    [InlineData("")]
    [InlineData("<ram:ID><!--keep-->one</ram:ID>")]
    [InlineData("<ram:ID><?keep data?>one</ram:ID>")]
    public void MissingOrAnnotatedFieldsCannotLoseContentDuringEditing(string field) {
        byte[] bytes = Encoding.UTF8.GetBytes(Invoice("INV-1").Replace("<ram:ID>INV-1</ram:ID>", field));
        var document = PdfCiiInvoiceDocument.Load(bytes);
        Assert.Throws<InvalidOperationException>(() => document.WithDocumentId("replacement"));
        Assert.Equal(bytes, document.ToBytes());
    }

    [Fact]
    public void SignedXmlCanBePassedThroughButCannotBeEdited() {
        string xml = Invoice("INV-1").Replace("</rsm:CrossIndustryInvoice>",
            "<Signature xmlns='http://www.w3.org/2000/09/xmldsig#'/></rsm:CrossIndustryInvoice>");
        var document = Load(xml);
        Assert.True(document.HasXmlSignature);
        Assert.Equal(Encoding.UTF8.GetBytes(xml), document.ToBytes());
        Assert.Throws<InvalidOperationException>(() => document.WithDocumentId("INV-2"));
        Assert.Throws<InvalidOperationException>(() => document.WithIssueDate(DateTime.Today));
    }

    [Fact]
    public void DtdMalformedAndExcessiveInputsAreRejected() {
        Assert.Throws<XmlException>(() => Load("<!DOCTYPE rsm:CrossIndustryInvoice [<!ENTITY x 'expanded'>]>" + Invoice("&x;")));
        Assert.Throws<XmlException>(() => Load(Invoice("INV-1") + "<other/>"));
        Assert.Throws<InvalidDataException>(() => PdfCiiInvoiceDocument.Load(new byte[PdfCiiInvoiceDocument.MaximumXmlBytes + 1]));
        string nested = string.Concat(Enumerable.Repeat("<x:a>", 130)) + string.Concat(Enumerable.Repeat("</x:a>", 130));
        Assert.Throws<InvalidDataException>(() => Load(Invoice("INV-1").Replace("<!--retain-->", nested)));
    }

    [Fact]
    public void UnsupportedDateRepresentationIsPreservedWithoutReinterpretation() {
        var document = Load(Invoice("INV-1").Replace("format='102'", "format='610'"));
        Assert.Null(document.IssueDate);
        Assert.Throws<InvalidOperationException>(() => document.WithIssueDate(DateTime.Today));
    }

    [Fact]
    public void InvalidOrOversizedIdentifierEditsPreserveTheOriginal() {
        var document = Load(Invoice("INV-1"));
        Assert.Throws<ArgumentException>(() => document.WithDocumentId(" "));
        Assert.Throws<XmlException>(() => document.WithDocumentId("invalid\0identifier"));
        Assert.Throws<ArgumentException>(() => document.WithDocumentId(new string('x', PdfCiiInvoiceDocument.MaximumXmlBytes + 1)));
        Assert.Equal("INV-1", document.DocumentId);
    }

    [Fact]
    public void PdfBridgeUsesTheSameSnapshotAndGroundworkAsByteAttachment() {
        var invoice = Load(Invoice("INV-1")).WithDocumentId("INV-2");
        var expected = new PdfOptions().UseFacturX(invoice.ToBytes());
        var actual = new PdfOptions().UseFacturXDocument(invoice);
        Assert.Equal(expected.EmbeddedFiles.Single().Data, actual.EmbeddedFiles.Single().Data);
        Assert.Equal(expected.FileVersion, actual.FileVersion);
        Assert.Equal(expected.ElectronicInvoiceMetadata!.ConformanceLevel, actual.ElectronicInvoiceMetadata!.ConformanceLevel);
    }

    private static PdfCiiInvoiceDocument Load(string xml) => PdfCiiInvoiceDocument.Load(Encoding.UTF8.GetBytes(xml));

    private sealed class NonSeekableInvoiceStream : MemoryStream {
        internal NonSeekableInvoiceStream(byte[] bytes) : base(bytes, writable: false) { }
        internal int BytesRead { get; private set; }
        public override bool CanSeek => false;
        public override long Length => throw new NotSupportedException();
        public override long Position {
            get => throw new NotSupportedException();
            set => throw new NotSupportedException();
        }
        public override int Read(byte[] buffer, int offset, int count) {
            int read = base.Read(buffer, offset, count);
            BytesRead += read;
            return read;
        }
    }

    private sealed class MutatingInvoiceStream : MemoryStream {
        internal byte[]? Captured { get; private set; }
        public override void Write(byte[] buffer, int offset, int count) {
            Captured = buffer;
            buffer[offset] ^= 1;
            base.Write(buffer, offset, count);
        }
    }

    private static string Invoice(string id) =>
        $"<rsm:CrossIndustryInvoice xmlns:rsm='{Rsm}' xmlns:ram='{Ram}' xmlns:udt='{Udt}' xmlns:x='urn:custom' x:flag='keep'>\n" +
        "<!--retain--><rsm:ExchangedDocument><ram:ID>" + id + "</ram:ID><ram:TypeCode>380</ram:TypeCode>" +
        "<ram:IssueDateTime><udt:DateTimeString format='102'>20260905</udt:DateTimeString></ram:IssueDateTime></rsm:ExchangedDocument>" +
        "<x:Extension><x:ID>extension-id</x:ID><x:Text><![CDATA[unknown <content>]]></x:Text></x:Extension>" +
        "<rsm:SupplyChainTradeTransaction><ram:ApplicableHeaderTradeSettlement><ram:PaymentReference>INV-1</ram:PaymentReference>" +
        "<ram:InvoiceCurrencyCode>EUR</ram:InvoiceCurrencyCode></ram:ApplicableHeaderTradeSettlement></rsm:SupplyChainTradeTransaction>" +
        "</rsm:CrossIndustryInvoice>";
}
